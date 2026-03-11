import axios from 'axios';
import { DWClient, TOPIC_ROBOT } from 'dingtalk-stream';

import { ASSISTANT_NAME } from '../config.js';
import { readEnvFile } from '../env.js';
import { logger } from '../logger.js';
import { registerChannel, ChannelOpts } from './registry.js';
import {
  Channel,
  OnChatMetadata,
  OnInboundMessage,
  RegisteredGroup,
} from '../types.js';

interface DingTalkChannelOpts {
  clientId: string;
  clientSecret: string;
  onMessage: OnInboundMessage;
  onChatMetadata: OnChatMetadata;
  registeredGroups: () => Record<string, RegisteredGroup>;
}

interface DingTalkMessageData {
  msgId: string;
  conversationId: string;
  conversationType: string; // '1' for direct, '2' for group
  senderId: string;
  senderStaffId?: string;
  senderNick: string;
  chatbotUserId: string;
  msgtype: string;
  text?: { content: string };
  createAt: number;
  conversationTitle?: string;
}

interface DingTalkWebSocketResponse {
  headers: { messageId: string };
  data: string;
}

export class DingTalkChannel implements Channel {
  name = 'dingtalk';

  private client: DWClient;
  private opts: DingTalkChannelOpts;
  private _connected = false;
  private accessToken: string | null = null;
  private tokenExpiry = 0;

  constructor(opts: DingTalkChannelOpts) {
    this.opts = opts;
    this.client = new DWClient({
      clientId: opts.clientId,
      clientSecret: opts.clientSecret,
      debug: false,
    });
  }

  async connect(): Promise<void> {
    this.client.registerCallbackListener(TOPIC_ROBOT, (res: DingTalkWebSocketResponse) => {
      this.handleInbound(res).catch((err) =>
        logger.error({ err }, 'DingTalk: error handling inbound message'),
      );
    });

    await this.client.connect();
    this._connected = true;
    logger.info('DingTalk connected');
  }

  private async handleInbound(res: DingTalkWebSocketResponse): Promise<void> {
    let data: DingTalkMessageData;
    try {
      data = JSON.parse(res.data) as DingTalkMessageData;
    } catch (err) {
      logger.error({ err }, 'DingTalk: failed to parse inbound payload');
      return;
    }

    // Skip self-messages (bot loop prevention)
    if (data.senderId === data.chatbotUserId) return;

    // Acknowledge to prevent server retries
    this.client.socketCallBackResponse(res.headers.messageId, {
      status: 'SUCCESS',
    });

    const isDirect = data.conversationType === '1';
    const chatJid = isDirect
      ? `dt:user:${data.senderId}`
      : `dt:group:${data.conversationId}`;

    // Extract text content
    let content = data.msgtype === 'text' && data.text ? data.text.content : '';

    // Strip @bot mentions from content (DingTalk includes them in text)
    content = content.replace(/@\S+/g, '').trim();

    // Prepend trigger so NanoClaw's TRIGGER_PATTERN matches group messages
    // DingTalk only delivers to robot when @mentioned — every group message is a trigger
    if (!isDirect) {
      content = `@${ASSISTANT_NAME} ${content}`;
    }

    // Media type prefixes
    if (data.msgtype === 'picture') content = `[Image]\n${content}`;
    else if (data.msgtype === 'audio') content = `[Audio]\n${content}`;
    else if (data.msgtype === 'file') content = `[File]\n${content}`;
    else if (data.msgtype === 'video') content = `[Video]\n${content}`;

    const timestamp = new Date(data.createAt).toISOString();
    const senderName = data.senderNick;
    const sender = data.senderStaffId || data.senderId;
    const chatName = isDirect
      ? senderName
      : (data.conversationTitle || chatJid);

    this.opts.onChatMetadata(chatJid, timestamp, chatName, 'dingtalk', !isDirect);

    const group = this.opts.registeredGroups()[chatJid];
    if (!group) {
      logger.debug({ chatJid, chatName }, 'DingTalk: message from unregistered chat');
      return;
    }

    this.opts.onMessage(chatJid, {
      id: data.msgId,
      chat_jid: chatJid,
      sender,
      sender_name: senderName,
      content,
      timestamp,
      is_from_me: false,
    });

    logger.info({ chatJid, chatName, sender: senderName }, 'DingTalk message stored');
  }

  async sendMessage(jid: string, text: string): Promise<void> {
    try {
      const token = await this.getAccessToken();
      const isDirect = jid.startsWith('dt:user:');
      const id = jid.replace(/^dt:(user|group):/, '');

      if (isDirect) {
        await axios.post(
          'https://api.dingtalk.com/v1.0/robot/oToMessages/batchSend',
          {
            robotCode: this.opts.clientId,
            userIds: [id],
            msgKey: 'sampleMarkdown',
            msgParam: JSON.stringify({ title: ASSISTANT_NAME, text }),
          },
          { headers: { 'x-acs-dingtalk-access-token': token } },
        );
      } else {
        await axios.post(
          'https://api.dingtalk.com/v1.0/robot/groupMessages/send',
          {
            robotCode: this.opts.clientId,
            openConversationId: id,
            msgKey: 'sampleMarkdown',
            msgParam: JSON.stringify({ title: ASSISTANT_NAME, text }),
          },
          { headers: { 'x-acs-dingtalk-access-token': token } },
        );
      }

      logger.info({ jid, length: text.length }, 'DingTalk message sent');
    } catch (err) {
      logger.error({ jid, err }, 'DingTalk: failed to send message');
    }
  }

  isConnected(): boolean {
    return this._connected && this.client.connected;
  }

  ownsJid(jid: string): boolean {
    return jid.startsWith('dt:');
  }

  async disconnect(): Promise<void> {
    this.client.disconnect();
    this._connected = false;
    logger.info('DingTalk disconnected');
  }

  private async getAccessToken(): Promise<string> {
    if (this.accessToken && Date.now() < this.tokenExpiry) {
      return this.accessToken;
    }

    const res = await axios.post(
      'https://api.dingtalk.com/v1.0/oauth2/accessToken',
      {
        appKey: this.opts.clientId,
        appSecret: this.opts.clientSecret,
      },
    );

    this.accessToken = res.data.accessToken;
    // Refresh 60 seconds before expiry
    this.tokenExpiry = Date.now() + (res.data.expireIn - 60) * 1000;
    return this.accessToken!;
  }
}

registerChannel('dingtalk', (opts: ChannelOpts) => {
  const env = readEnvFile([
    'ENABLE_DINGTALK',
    'DINGTALK_CLIENT_ID',
    'DINGTALK_CLIENT_SECRET',
  ]);
  const enabled =
    process.env.ENABLE_DINGTALK === 'true' || env.ENABLE_DINGTALK === 'true';
  if (!enabled) return null;

  const clientId = env.DINGTALK_CLIENT_ID;
  const clientSecret = env.DINGTALK_CLIENT_SECRET;
  if (!clientId || !clientSecret) {
    logger.warn(
      'DingTalk enabled but DINGTALK_CLIENT_ID or DINGTALK_CLIENT_SECRET missing in .env',
    );
    return null;
  }

  return new DingTalkChannel({ ...opts, clientId, clientSecret });
});
