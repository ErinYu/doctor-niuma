import fs from 'fs';
import * as lark from '@larksuiteoapi/node-sdk';

import { ASSISTANT_NAME } from '../config.js';
import { readEnvFile } from '../env.js';
import { logger } from '../logger.js';
import { registerChannel, ChannelOpts } from './registry.js';
import {
  CardContent,
  Channel,
  OnChatMetadata,
  OnInboundMessage,
  RegisteredGroup,
} from '../types.js';

interface FeishuChannelOpts {
  appId: string;
  appSecret: string;
  onMessage: OnInboundMessage;
  onChatMetadata: OnChatMetadata;
  registeredGroups: () => Record<string, RegisteredGroup>;
}

// Inbound event shape from @larksuiteoapi/node-sdk
interface FeishuMessageEvent {
  sender: {
    sender_id?: { open_id?: string };
    sender_type: string;
  };
  message: {
    message_id: string;
    create_time: string;
    chat_id: string;
    chat_type: string;
    message_type: string;
    content: string;
  };
}

export class FeishuChannel implements Channel {
  name = 'feishu';

  private client: lark.Client;
  private wsClient: lark.WSClient;
  private opts: FeishuChannelOpts;
  private _connected = false;

  constructor(opts: FeishuChannelOpts) {
    this.opts = opts;
    this.client = new lark.Client({
      appId: opts.appId,
      appSecret: opts.appSecret,
    });
    this.wsClient = new lark.WSClient({
      appId: opts.appId,
      appSecret: opts.appSecret,
      loggerLevel: lark.LoggerLevel.warn,
    });
  }

  async connect(): Promise<void> {
    const eventDispatcher = new lark.EventDispatcher({}).register({
      'im.message.receive_v1': async (data: FeishuMessageEvent) => {
        logger.debug({ data }, 'Feishu: raw event received');
        try {
          await this.handleInbound(data);
        } catch (err) {
          logger.error({ err }, 'Feishu: error handling inbound message');
        }
      },
    });

    // start() is async — do not await (it blocks indefinitely keeping the WS alive)
    // but register connected state after a short delay to let the handshake finish
    this.wsClient.start({ eventDispatcher }).catch((err: unknown) => {
      logger.error({ err }, 'Feishu: WebSocket error');
      this._connected = false;
    });

    // Give the SDK a moment to complete the handshake before marking connected
    await new Promise((resolve) => setTimeout(resolve, 500));
    this._connected = true;
    logger.info('Feishu connected via WebSocket long-connection');
  }

  private async handleInbound(data: FeishuMessageEvent): Promise<void> {
    const { sender, message: msg } = data;

    logger.info(
      {
        sender_type: sender.sender_type,
        chat_type: msg.chat_type,
        chat_id: msg.chat_id,
        open_id: sender.sender_id?.open_id,
      },
      'Feishu: inbound event',
    );

    // Skip non-user messages (bots, system)
    if (sender.sender_type !== 'user') return;

    const isDirect = msg.chat_type === 'p2p';
    const openId = sender.sender_id?.open_id ?? '';
    const chatId = msg.chat_id;

    const chatJid = isDirect ? `fs:user:${openId}` : `fs:group:${chatId}`;

    // Parse message content
    let content = '';
    if (msg.message_type === 'text') {
      try {
        const parsed = JSON.parse(msg.content) as { text?: string };
        content = parsed.text ?? '';
        // Strip Feishu @mention tags (format: @_user_xxxxx or <at user_id="...">name</at>)
        content = content
          .replace(/<at[^>]*>[^<]*<\/at>/g, '')
          .replace(/@_[a-z]+_[a-zA-Z0-9]+/g, '')
          .trim();
      } catch {
        content = '';
      }
    } else if (msg.message_type === 'image') {
      content = '[Image]';
    } else if (msg.message_type === 'audio') {
      content = '[Audio]';
    } else if (msg.message_type === 'file') {
      content = '[File]';
    } else if (msg.message_type === 'video') {
      content = '[Video]';
    }

    // Prepend trigger for group messages
    if (!isDirect) {
      content = `@${ASSISTANT_NAME} ${content}`;
    }

    const timestamp = new Date(Number(msg.create_time)).toISOString();

    this.opts.onChatMetadata(chatJid, timestamp, chatJid, 'feishu', !isDirect);

    const group = this.opts.registeredGroups()[chatJid];
    if (!group) {
      logger.debug({ chatJid }, 'Feishu: message from unregistered chat');
      return;
    }

    this.opts.onMessage(chatJid, {
      id: msg.message_id,
      chat_jid: chatJid,
      sender: openId,
      sender_name: openId,
      content,
      timestamp,
      is_from_me: false,
    });

    logger.info({ chatJid, openId }, 'Feishu message stored');
  }

  async sendMessage(jid: string, text: string): Promise<void> {
    try {
      const isDirect = jid.startsWith('fs:user:');
      const id = jid.replace(/^fs:(user|group):/, '');
      const receiveIdType = isDirect ? 'open_id' : 'chat_id';

      await this.client.im.message.create({
        params: { receive_id_type: receiveIdType },
        data: {
          receive_id: id,
          msg_type: 'text',
          content: JSON.stringify({ text }),
        },
      });

      logger.info({ jid, length: text.length }, 'Feishu message sent');
    } catch (err) {
      logger.error({ jid, err }, 'Feishu: failed to send message');
    }
  }

  async sendFile(
    jid: string,
    filePath: string,
    fileName: string,
  ): Promise<void> {
    try {
      const isDirect = jid.startsWith('fs:user:');
      const id = jid.replace(/^fs:(user|group):/, '');
      const receiveIdType = isDirect ? 'open_id' : 'chat_id';

      // Step 1: Upload file to Feishu to get file_key
      const fileStream = fs.createReadStream(filePath);
      const uploadRes = await this.client.im.file.create({
        data: {
          file_type: 'stream',
          file_name: fileName,
          file: fileStream,
        },
      });
      const fileKey = (uploadRes as { data?: { file_key?: string } }).data
        ?.file_key;
      if (!fileKey) throw new Error('Feishu file upload returned no file_key');

      // Step 2: Send file message
      await this.client.im.message.create({
        params: { receive_id_type: receiveIdType },
        data: {
          receive_id: id,
          msg_type: 'file',
          content: JSON.stringify({ file_key: fileKey }),
        },
      });

      logger.info({ jid, fileName }, 'Feishu file sent');
    } catch (err) {
      logger.error({ jid, fileName, err }, 'Feishu: failed to send file');
    }
  }

  async sendCard(jid: string, card: CardContent): Promise<void> {
    try {
      const isDirect = jid.startsWith('fs:user:');
      const id = jid.replace(/^fs:(user|group):/, '');
      const receiveIdType = isDirect ? 'open_id' : 'chat_id';

      // Build Feishu interactive card format
      const cardContent = buildFeishuCard(card);

      await this.client.im.message.create({
        params: { receive_id_type: receiveIdType },
        data: {
          receive_id: id,
          msg_type: 'interactive',
          content: JSON.stringify(cardContent),
        },
      });

      logger.info({ jid }, 'Feishu card sent');
    } catch (err) {
      logger.error({ jid, err }, 'Feishu: failed to send card');
    }
  }

  isConnected(): boolean {
    return this._connected;
  }

  ownsJid(jid: string): boolean {
    return jid.startsWith('fs:');
  }

  async disconnect(): Promise<void> {
    this._connected = false;
    logger.info('Feishu disconnected');
  }
}

/**
 * Convert generic CardContent to Feishu interactive card format
 * @see https://open.larksuite.com/document/home/interactive-cards-card
 */
function buildFeishuCard(card: CardContent): Record<string, unknown> {
  const elements = card.elements || [];

  // Build header if present
  let header: Record<string, unknown> | undefined;
  if (card.header?.title?.content) {
    header = {
      title: {
        tag: 'plain_text',
        content: card.header.title.content,
      },
      subtitle: card.header.subtitle
        ? {
            tag: 'plain_text',
            content: card.header.subtitle.content,
          }
        : undefined,
    };
  }

  // Build card config
  return {
    config: {
      wide_screen_mode: true,
    },
    header,
    elements,
  };
}

registerChannel('feishu', (opts: ChannelOpts) => {
  const env = readEnvFile([
    'ENABLE_FEISHU',
    'FEISHU_APP_ID',
    'FEISHU_APP_SECRET',
  ]);
  const enabled =
    process.env.ENABLE_FEISHU === 'true' || env.ENABLE_FEISHU === 'true';
  if (!enabled) return null;

  const appId = env.FEISHU_APP_ID;
  const appSecret = env.FEISHU_APP_SECRET;
  if (!appId || !appSecret) {
    logger.warn(
      'Feishu enabled but FEISHU_APP_ID or FEISHU_APP_SECRET missing in .env',
    );
    return null;
  }

  return new FeishuChannel({ ...opts, appId, appSecret });
});
