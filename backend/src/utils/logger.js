/**
 * 🔐 ログ記録ユーティリティ
 * セキュリティイベントと一般ログの記録
 */

const winston = require('winston');
const path = require('path');

// ログフォーマット
const logFormat = winston.format.combine(
  winston.format.timestamp({
    format: 'YYYY-MM-DD HH:mm:ss'
  }),
  winston.format.errors({ stack: true }),
  winston.format.json(),
  winston.format.printf(({ level, message, timestamp, ...meta }) => {
    return JSON.stringify({
      timestamp,
      level,
      message,
      ...meta
    });
  })
);

// メインロガー
const logger = winston.createLogger({
  level: process.env.LOG_LEVEL || 'info',
  format: logFormat,
  defaultMeta: { service: 'excel-chatbot' },
  transports: [
    // エラーログファイル
    new winston.transports.File({
      filename: path.join(process.cwd(), 'logs', 'error.log'),
      level: 'error',
      maxsize: 5242880, // 5MB
      maxFiles: 5
    }),
    
    // 統合ログファイル
    new winston.transports.File({
      filename: path.join(process.cwd(), 'logs', 'combined.log'),
      maxsize: 5242880, // 5MB
      maxFiles: 10
    })
  ]
});

// 開発環境ではコンソール出力も追加
if (process.env.NODE_ENV !== 'production') {
  logger.add(new winston.transports.Console({
    format: winston.format.combine(
      winston.format.colorize(),
      winston.format.simple()
    )
  }));
}

// セキュリティ専用ロガー
const securityLogger = winston.createLogger({
  level: 'warn',
  format: logFormat,
  defaultMeta: { service: 'security' },
  transports: [
    new winston.transports.File({
      filename: path.join(process.cwd(), 'logs', 'security.log'),
      maxsize: 5242880, // 5MB
      maxFiles: 20 // セキュリティログは長期保存
    })
  ]
});

// 監査専用ロガー
const auditLogger = winston.createLogger({
  level: 'info',
  format: logFormat,
  defaultMeta: { service: 'audit' },
  transports: [
    new winston.transports.File({
      filename: path.join(process.cwd(), 'logs', 'audit.log'),
      maxsize: 10485760, // 10MB
      maxFiles: 30 // 監査ログは最長保存
    })
  ]
});

/**
 * セキュリティイベントのログ記録
 */
const logSecurityEvent = (event, details = {}) => {
  securityLogger.warn('Security Event', {
    event,
    timestamp: new Date().toISOString(),
    ...details
  });
};

/**
 * 監査イベントのログ記録
 */
const logAuditEvent = (action, user, details = {}) => {
  auditLogger.info('Audit Event', {
    action,
    user,
    timestamp: new Date().toISOString(),
    ...details
  });
};

/**
 * ユーザーアクションのログ記録
 */
const logUserAction = (userId, action, resource, ip, userAgent, details = {}) => {
  auditLogger.info('User Action', {
    userId,
    action,
    resource,
    ip,
    userAgent,
    timestamp: new Date().toISOString(),
    ...details
  });
};

/**
 * API呼び出しのログ記録
 */
const logAPICall = (method, path, statusCode, responseTime, ip, userId = null) => {
  logger.info('API Call', {
    method,
    path,
    statusCode,
    responseTime,
    ip,
    userId,
    timestamp: new Date().toISOString()
  });
};

/**
 * データベース操作のログ記録
 */
const logDBOperation = (operation, collection, userId, result) => {
  auditLogger.info('Database Operation', {
    operation,
    collection,
    userId,
    result: result ? 'success' : 'failure',
    timestamp: new Date().toISOString()
  });
};

/**
 * 認証イベントのログ記録
 */
const logAuthEvent = (event, username, ip, userAgent, success = true, reason = null) => {
  const level = success ? 'info' : 'warn';
  
  securityLogger.log(level, 'Authentication Event', {
    event,
    username,
    ip,
    userAgent,
    success,
    reason,
    timestamp: new Date().toISOString()
  });
};

/**
 * エラーのログ記録（スタックトレース付き）
 */
const logError = (error, context = {}) => {
  logger.error('Application Error', {
    message: error.message,
    stack: error.stack,
    ...context,
    timestamp: new Date().toISOString()
  });
};

/**
 * パフォーマンス監視ログ
 */
const logPerformance = (operation, duration, details = {}) => {
  logger.info('Performance Metrics', {
    operation,
    duration,
    ...details,
    timestamp: new Date().toISOString()
  });
};

/**
 * チャット会話のログ記録
 */
const logChatInteraction = (userId, conversationId, messageLength, responseTime, model) => {
  auditLogger.info('Chat Interaction', {
    userId,
    conversationId,
    messageLength,
    responseTime,
    model,
    timestamp: new Date().toISOString()
  });
};

module.exports = {
  logger,
  securityLogger,
  auditLogger,
  logSecurityEvent,
  logAuditEvent,
  logUserAction,
  logAPICall,
  logDBOperation,
  logAuthEvent,
  logError,
  logPerformance,
  logChatInteraction
};