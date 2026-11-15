// whatsapp_service.js - Official WhatsApp Web Multi-Device Gateway using Baileys
const { default: makeWASocket, useMultiFileAuthState, DisconnectReason } = require('@whiskeysockets/baileys');
const QRCode = require('qrcode');
const path = require('path');
const fs = require('fs');

const AUTH_DIR = path.join(__dirname, 'auth_baileys');

let sock = null;
let currentQR = null;
let connectionState = 'close';
let linkedPhone = null;
let isStarting = false;

async function startWhatsAppService() {
  if (isStarting) return;
  if (sock && connectionState === 'open') return;

  isStarting = true;

  try {
    if (!fs.existsSync(AUTH_DIR)) {
      fs.mkdirSync(AUTH_DIR, { recursive: true });
    }

    const { state, saveCreds } = await useMultiFileAuthState(AUTH_DIR);

    sock = makeWASocket({
      auth: state,
      printQRInTerminal: false,
      browser: ['BDO Bani Portal', 'Chrome', '1.0.0'],
      syncFullHistory: false
    });

    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', async (update) => {
      const { connection, lastDisconnect, qr } = update;

      if (qr) {
        try {
          currentQR = await QRCode.toDataURL(qr, {
            scale: 8,
            margin: 2,
            color: { dark: '#020617', light: '#ffffff' }
          });
          /* console.log('[WhatsApp-Service] Fresh Live QR Code Generated!'); */
        } catch (e) {
          console.error('[WhatsApp-Service] QR generation error:', e);
        }
      }

      if (connection === 'close') {
        connectionState = 'close';
        currentQR = null;
        const statusCode = lastDisconnect?.error?.output?.statusCode;
        const shouldReconnect = statusCode !== DisconnectReason.loggedOut;
        /* console.log('[WhatsApp-Service] Connection closed. Reason code:', statusCode, 'Reconnect:', shouldReconnect); */
        if (shouldReconnect) {
          isStarting = false;
          setTimeout(startWhatsAppService, 4000);
        } else {
          isStarting = false;
          disconnect();
        }
      } else if (connection === 'open') {
        connectionState = 'open';
        currentQR = null;
        isStarting = false;
        try {
          const userJid = sock.user?.id || '';
          linkedPhone = userJid.split(':')[0] || 'Admin Phone';
        } catch (e) {
          linkedPhone = 'Admin Phone';
        }
        /* console.log('🎉 [WhatsApp-Service] DEVICE LINKED SUCCESSFULLY! Linked Phone:', linkedPhone); */
      } else if (connection === 'connecting') {
        connectionState = 'connecting';
      }
    });

  } catch (err) {
    console.error('[WhatsApp-Service] Initialization error:', err);
    isStarting = false;
  }
}

async function getStatus() {
  if (!sock && connectionState !== 'connecting' && connectionState !== 'open') {
    startWhatsAppService();
  }
  return {
    connected: connectionState === 'open',
    status: connectionState,
    phone: linkedPhone,
    qr: currentQR
  };
}

async function sendOtp(phone, otpCode, targetUser) {
  if (!sock || connectionState !== 'open') {
    return { success: false, error: 'WhatsApp is not linked yet. Please scan the QR code to link.' };
  }

  const cleanPhone = (phone || '').replace(/[^0-9]/g, '');
  if (!cleanPhone) {
    return { success: false, error: 'Invalid phone number provided.' };
  }

  const jid = cleanPhone + '@s.whatsapp.net';
  const textMsg = `🏛️ *OFFICIAL BDO BANI MONITORING PORTAL*\n\nYour 6-Digit WhatsApp Login OTP is: *${otpCode}*\n\nOfficer: ${targetUser}\nStatus: Live Verified\nValid for 10 minutes. Do not share.`;

  try {
    await sock.sendMessage(jid, { text: textMsg });
    /* console.log(`[WhatsApp-Service] ✅ OTP (${otpCode}) sent successfully to ${jid}`); */
    return { success: true };
  } catch (err) {
    console.error(`[WhatsApp-Service] ❌ Failed to send OTP to ${jid}:`, err);
    return { success: false, error: err.message };
  }
}

async function disconnect() {
  try {
    if (sock) {
      sock.logout().catch(() => {});
      sock = null;
    }
  } catch (e) {}
  connectionState = 'close';
  currentQR = null;
  linkedPhone = null;
  isStarting = false;

  // Clear auth files so fresh QR can be scanned next time
  try {
    if (fs.existsSync(AUTH_DIR)) {
      fs.rmSync(AUTH_DIR, { recursive: true, force: true });
    }
  } catch (e) {}
  return { success: true };
}

module.exports = {
  startWhatsAppService,
  getStatus,
  sendOtp,
  disconnect
};
