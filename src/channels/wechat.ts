import { WechatyBuilder, Wechaty, Message } from 'wechaty';
import qrcodeTerminal from 'qrcode-terminal';

import { ASSISTANT_NAME, TRIGGER_PATTERN } from '../config.js';
import { readEnvFile } from '../env.js';
import { logger } from '../logger.js';
import { registerChannel, ChannelOpts } from './registry.js';
import {
  Channel,
  OnChatMetadata,
  OnInboundMessage,
  RegisteredGroup,
} from '../types.js';

export interface WeChatChannelOpts {
  onMessage: OnInboundMessage;
  onChatMetadata: OnChatMetadata;
  registeredGroups: () => Record<string, RegisteredGroup>;
}

export class WeChatChannel implements Channel {
  name = 'wechat';

  private bot: Wechaty;
  private opts: WeChatChannelOpts;

  constructor(opts: WeChatChannelOpts) {
    this.opts = opts;
    this.bot = WechatyBuilder.build({
      name: 'wechat-bot',
    });
  }

  async connect(): Promise<void> {
    this.bot.on('scan', (qrcode, status) => {
      logger.info({ status }, 'WeChat scan needed');
      qrcodeTerminal.generate(qrcode, { small: true });
    });

    this.bot.on('login', (user) => {
      logger.info({ user: user.name() }, 'WeChat connected');
    });

    this.bot.on('logout', (user) => {
      logger.info({ user: user.name() }, 'WeChat disconnected');
    });

    this.bot.on('message', async (message: Message) => {
      if (message.self()) return;

      const contact = message.talker();
      const room = message.room();

      const chatJid = room ? `wc:room:${room.id}` : `wc:user:${contact.id}`;
      let content = message.text();
      const timestamp = message.date().toISOString();
      const senderName = contact.name();
      const sender = contact.id;
      const msgId = message.id;

      let chatName = senderName;
      if (room) {
        chatName = (await room.topic()) || chatName;
      }

      // Prepend trigger pattern for rooms if bot is mentioned
      if (room && (await message.mentionSelf())) {
        content = `@${ASSISTANT_NAME} ${content}`;
      }

      // Handle attachments
      if (message.type() === this.bot.Message.Type.Image) {
        content = `[Image]\n${content}`;
      } else if (message.type() === this.bot.Message.Type.Attachment) {
        content = `[File]\n${content}`;
      } else if (message.type() === this.bot.Message.Type.Audio) {
        content = `[Audio]\n${content}`;
      } else if (message.type() === this.bot.Message.Type.Video) {
        content = `[Video]\n${content}`;
      }

      const isGroup = !!room;
      this.opts.onChatMetadata(chatJid, timestamp, chatName, 'wechat', isGroup);

      const group = this.opts.registeredGroups()[chatJid];
      if (!group) {
        logger.debug(
          { chatJid, chatName },
          'Message from unregistered WeChat channel',
        );
        return;
      }

      this.opts.onMessage(chatJid, {
        id: msgId,
        chat_jid: chatJid,
        sender,
        sender_name: senderName,
        content,
        timestamp,
        is_from_me: false,
      });

      logger.info(
        { chatJid, chatName, sender: senderName },
        'WeChat message stored',
      );
    });

    await this.bot.start();
  }

  async sendMessage(jid: string, text: string): Promise<void> {
    try {
      if (jid.startsWith('wc:room:')) {
        const roomId = jid.replace('wc:room:', '');
        const room = await this.bot.Room.find({ id: roomId });
        if (room) await room.say(text);
      } else if (jid.startsWith('wc:user:')) {
        const userId = jid.replace('wc:user:', '');
        const contact = await this.bot.Contact.find({ id: userId });
        if (contact) await contact.say(text);
      }
      logger.info({ jid, length: text.length }, 'WeChat message sent');
    } catch (err) {
      logger.error({ jid, err }, 'Failed to send WeChat message');
    }
  }

  isConnected(): boolean {
    return this.bot.isLoggedIn;
  }

  ownsJid(jid: string): boolean {
    return jid.startsWith('wc:');
  }

  async disconnect(): Promise<void> {
    await this.bot.stop();
    logger.info('WeChat bot stopped');
  }

  async setTyping(jid: string, isTyping: boolean): Promise<void> {
    // WeChat doesn't have an explicit API for setting typing status in Wechaty
  }
}

registerChannel('wechat', (opts: ChannelOpts) => {
  const envVars = readEnvFile(['ENABLE_WECHAT']);
  if (
    process.env.ENABLE_WECHAT === 'true' ||
    envVars.ENABLE_WECHAT === 'true'
  ) {
    return new WeChatChannel(opts);
  }
  return null;
});
