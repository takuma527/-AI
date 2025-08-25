/**
 * 💬 WebSocket サービス
 * リアルタイムチャット機能の実装
 */

const jwt = require('jsonwebtoken');
const User = require('../models/User');
const { logUserAction, logSecurityEvent } = require('../utils/logger');

/**
 * Socket.IO 接続とイベントハンドリング
 */
module.exports = (io) => {
  // Socket認証ミドルウェア
  io.use(async (socket, next) => {
    try {
      const token = socket.handshake.auth.token || socket.handshake.headers.authorization?.replace('Bearer ', '');
      
      if (!token) {
        logSecurityEvent('WEBSOCKET_NO_TOKEN', {
          socketId: socket.id,
          ip: socket.handshake.address,
          userAgent: socket.handshake.headers['user-agent']
        });
        return next(new Error('Authentication error: No token provided'));
      }
      
      const decoded = jwt.verify(token, process.env.JWT_SECRET);
      const user = await User.findById(decoded.id).select('-password');
      
      if (!user || !user.isActive || user.isLocked) {
        logSecurityEvent('WEBSOCKET_INVALID_USER', {
          socketId: socket.id,
          userId: decoded.id,
          ip: socket.handshake.address
        });
        return next(new Error('Authentication error: Invalid user'));
      }
      
      socket.userId = user._id.toString();
      socket.username = user.username;
      socket.role = user.role;
      
      logUserAction(
        user._id,
        'WEBSOCKET_CONNECT',
        'websocket',
        socket.handshake.address,
        socket.handshake.headers['user-agent'],
        { socketId: socket.id }
      );
      
      next();
    } catch (error) {
      logSecurityEvent('WEBSOCKET_AUTH_ERROR', {
        socketId: socket.id,
        error: error.message,
        ip: socket.handshake.address
      });
      next(new Error('Authentication error'));
    }
  });
  
  // Socket接続処理
  io.on('connection', (socket) => {
    console.log(`✅ WebSocket接続: ${socket.username} (${socket.id})`);
    
    // ユーザーを専用ルームに参加させる
    socket.join(`user:${socket.userId}`);
    
    // チャットルーム参加
    socket.on('join-chat', (data) => {
      const { conversationId } = data;
      
      if (conversationId) {
        socket.join(`chat:${conversationId}`);
        
        logUserAction(
          socket.userId,
          'WEBSOCKET_JOIN_CHAT',
          `chat:${conversationId}`,
          socket.handshake.address,
          socket.handshake.headers['user-agent'],
          { conversationId }
        );
        
        socket.emit('joined-chat', { conversationId });
      }
    });
    
    // チャットメッセージ送信
    socket.on('send-message', async (data) => {
      try {
        const { conversationId, message, type = 'text' } = data;
        
        // 基本的な検証
        if (!message || typeof message !== 'string' || message.trim().length === 0) {
          socket.emit('error', { message: 'Invalid message format' });
          return;
        }
        
        if (message.length > 5000) {
          socket.emit('error', { message: 'Message too long' });
          return;
        }
        
        // ユーザーの使用制限チェック
        const user = await User.findById(socket.userId);
        if (!user.canAskQuestion()) {
          socket.emit('error', { 
            message: 'Daily limit exceeded',
            code: 'DAILY_LIMIT_EXCEEDED'
          });
          return;
        }
        
        // メッセージをチャットルームに送信
        const messageData = {
          id: `msg_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
          conversationId,
          userId: socket.userId,
          username: socket.username,
          message: message.trim(),
          type,
          timestamp: new Date().toISOString(),
          status: 'sent'
        };
        
        // チャットルームの他の参加者に送信
        socket.to(`chat:${conversationId}`).emit('new-message', messageData);
        
        // 送信者に確認を返す
        socket.emit('message-sent', messageData);
        
        logUserAction(
          socket.userId,
          'WEBSOCKET_SEND_MESSAGE',
          `chat:${conversationId}`,
          socket.handshake.address,
          socket.handshake.headers['user-agent'],
          {
            conversationId,
            messageLength: message.length,
            type
          }
        );
        
      } catch (error) {
        logSecurityEvent('WEBSOCKET_MESSAGE_ERROR', {
          socketId: socket.id,
          userId: socket.userId,
          error: error.message,
          ip: socket.handshake.address
        });
        
        socket.emit('error', { 
          message: 'Failed to send message',
          code: 'MESSAGE_SEND_ERROR'
        });
      }
    });
    
    // タイピング状態の通知
    socket.on('typing', (data) => {
      const { conversationId, isTyping } = data;
      
      if (conversationId) {
        socket.to(`chat:${conversationId}`).emit('user-typing', {
          userId: socket.userId,
          username: socket.username,
          isTyping: !!isTyping
        });
      }
    });
    
    // オンライン状態の更新
    socket.on('update-status', (data) => {
      const { status } = data;
      
      if (['online', 'away', 'busy'].includes(status)) {
        socket.broadcast.emit('user-status-changed', {
          userId: socket.userId,
          username: socket.username,
          status
        });
      }
    });
    
    // チャットルーム退出
    socket.on('leave-chat', (data) => {
      const { conversationId } = data;
      
      if (conversationId) {
        socket.leave(`chat:${conversationId}`);
        
        socket.to(`chat:${conversationId}`).emit('user-left', {
          userId: socket.userId,
          username: socket.username
        });
        
        logUserAction(
          socket.userId,
          'WEBSOCKET_LEAVE_CHAT',
          `chat:${conversationId}`,
          socket.handshake.address,
          socket.handshake.headers['user-agent'],
          { conversationId }
        );
      }
    });
    
    // 接続切断処理
    socket.on('disconnect', (reason) => {
      console.log(`❌ WebSocket切断: ${socket.username} (${socket.id}) - ${reason}`);
      
      // オフライン状態を他のユーザーに通知
      socket.broadcast.emit('user-status-changed', {
        userId: socket.userId,
        username: socket.username,
        status: 'offline'
      });
      
      logUserAction(
        socket.userId,
        'WEBSOCKET_DISCONNECT',
        'websocket',
        socket.handshake.address,
        socket.handshake.headers['user-agent'],
        { 
          socketId: socket.id,
          reason 
        }
      );
    });
    
    // エラーハンドリング
    socket.on('error', (error) => {
      logSecurityEvent('WEBSOCKET_ERROR', {
        socketId: socket.id,
        userId: socket.userId,
        error: error.message,
        ip: socket.handshake.address
      });
    });
    
    // 接続成功の通知
    socket.emit('connected', {
      message: 'Successfully connected to Excel ChatBot',
      userId: socket.userId,
      username: socket.username,
      timestamp: new Date().toISOString()
    });
  });
  
  // 定期的な接続状態チェック
  setInterval(() => {
    const connectedSockets = io.sockets.sockets.size;
    
    if (connectedSockets > 100) { // 同時接続数の監視
      logSecurityEvent('HIGH_WEBSOCKET_CONNECTIONS', {
        connectedSockets,
        timestamp: new Date().toISOString()
      });
    }
  }, 60000); // 1分ごと
};