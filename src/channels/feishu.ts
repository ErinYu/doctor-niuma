import fs from 'fs';
import * as lark from '@larksuiteoapi/node-sdk';

import { ASSISTANT_NAME } from '../config.js';
import { readEnvFile } from '../env.js';
import { resolveGroupFolderPath } from '../group-folder.js';
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
      // Download image and provide path to agent
      try {
        const imageData = JSON.parse(msg.content) as { image_key?: string };
        const imageKey = imageData.image_key;
        if (imageKey) {
          const imagePath = await this.downloadImage(msg.message_id, imageKey, chatJid);
          if (imagePath) {
            content = `[Image: ${imagePath}]`;
          } else {
            content = '[Image]';
          }
        } else {
          content = '[Image]';
        }
      } catch {
        content = '[Image]';
      }
    } else if (msg.message_type === 'audio') {
      content = '[Audio]';
    } else if (msg.message_type === 'file') {
      // Download file and provide path to agent
      try {
        const fileData = JSON.parse(msg.content) as { file_key?: string; file_name?: string };
        const fileKey = fileData.file_key;
        if (fileKey) {
          const filePath = await this.downloadFile(msg.message_id, fileKey, chatJid, fileData.file_name);
          if (filePath) {
            content = `[File: ${filePath}]`;
          } else {
            content = '[File]';
          }
        } else {
          content = '[File]';
        }
      } catch {
        content = '[File]';
      }
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
      logger.info({ chatJid }, 'Feishu: message from unregistered chat');
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

      // Use interactive card with markdown tag for proper table rendering
      await this.client.im.message.create({
        params: { receive_id_type: receiveIdType },
        data: {
          receive_id: id,
          msg_type: 'interactive',
          content: JSON.stringify({
            config: {
              wide_screen_mode: true,
            },
            elements: [
              {
                tag: 'markdown',
                content: text,
              },
            ],
          }),
        },
      });

      logger.info(
        { jid, length: text.length },
        'Feishu message sent (markdown)',
      );
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
      // Map extension to Feishu file_type
      const ext = fileName.split('.').pop()?.toLowerCase() || '';
      const fileTypeMap: Record<string, string> = {
        pdf: 'pdf', doc: 'doc', docx: 'doc',
        xls: 'xls', xlsx: 'xls',
        ppt: 'ppt', pptx: 'ppt',
        mp4: 'mp4', opus: 'opus',
      };
      const fileType = fileTypeMap[ext] || 'stream';

      const fileStream = fs.createReadStream(filePath);
      const uploadRes = await this.client.im.file.create({
        data: {
          file_type: fileType,
          file_name: fileName,
          file: fileStream,
        },
      });
      // SDK may return file_key at top level or nested under .data
      const res = uploadRes as Record<string, unknown>;
      const fileKey =
        (res.data as { file_key?: string })?.file_key
        || (res as { file_key?: string }).file_key;
      if (!fileKey) {
        logger.error({ uploadRes: JSON.stringify(uploadRes), fileName, fileType }, 'Feishu file upload returned no file_key');
        throw new Error('Feishu file upload returned no file_key');
      }

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

  async downloadFile(
    messageId: string,
    fileKey: string,
    chatJid: string,
    fileName?: string,
  ): Promise<string | null> {
    try {
      const group = this.opts.registeredGroups()[chatJid];
      if (!group) {
        logger.warn({ chatJid }, 'Cannot download file: group not registered');
        return null;
      }

      const safeName = fileName || `file_${Date.now()}`;

      // Create uploads directory on host (maps to /workspace/group/ in container)
      const groupDir = resolveGroupFolderPath(group.folder);
      const uploadsDir = `${groupDir}/uploads`;
      fs.mkdirSync(uploadsDir, { recursive: true });

      const filePath = `${uploadsDir}/${safeName}`;
      const containerPath = `/workspace/group/${group.folder}/uploads/${safeName}`;

      // Use correct Feishu API: download message resource by message_id + file_key
      const response = await fetch(
        `https://open.feishu.cn/open-apis/im/v1/messages/${messageId}/resources/${fileKey}?type=file`,
        {
          headers: {
            Authorization: `Bearer ${await this.getTenantToken()}`,
          },
        },
      );

      if (!response.ok) {
        logger.warn(
          { fileKey, messageId, status: response.status },
          'Failed to download file from Feishu',
        );
        return null;
      }

      const arrayBuffer = await response.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);
      fs.writeFileSync(filePath, buffer);

      logger.info({ fileKey, filePath, safeName }, 'Feishu file downloaded');
      return containerPath;
    } catch (err) {
      logger.error({ fileKey, err }, 'Feishu: failed to download file');
      return null;
    }
  }

  async downloadImage(messageId: string, imageKey: string, chatJid: string): Promise<string | null> {
    try {
      const group = this.opts.registeredGroups()[chatJid];
      if (!group) {
        logger.warn({ chatJid }, 'Cannot download image: group not registered');
        return null;
      }

      // Create uploads directory on host (maps to /workspace/group/ in container)
      const groupDir = resolveGroupFolderPath(group.folder);
      const uploadsDir = `${groupDir}/uploads`;
      fs.mkdirSync(uploadsDir, { recursive: true });

      const imageFileName = `image_${Date.now()}.png`;
      const filePath = `${uploadsDir}/${imageFileName}`;
      const containerPath = `/workspace/group/${group.folder}/uploads/${imageFileName}`;

      // Use correct Feishu API: download message resource by message_id + image_key
      const response = await fetch(
        `https://open.feishu.cn/open-apis/im/v1/messages/${messageId}/resources/${imageKey}?type=image`,
        {
          headers: {
            Authorization: `Bearer ${await this.getTenantToken()}`,
          },
        },
      );

      if (!response.ok) {
        logger.warn(
          { imageKey, messageId, status: response.status },
          'Failed to download image from Feishu',
        );
        return null;
      }

      const arrayBuffer = await response.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);
      fs.writeFileSync(filePath, buffer);

      logger.info({ imageKey, filePath }, 'Feishu image downloaded');
      return containerPath;
    } catch (err) {
      logger.error({ imageKey, err }, 'Feishu: failed to download image');
      return null;
    }
  }

  private async getTenantToken(): Promise<string> {
    // Get tenant access token using app credentials
    const response = await fetch(
      'https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal',
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          app_id: this.opts.appId,
          app_secret: this.opts.appSecret,
        }),
      },
    );

    const data = (await response.json()) as { tenant_access_token?: string };
    return data.tenant_access_token || '';
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
