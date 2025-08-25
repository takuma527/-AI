/**
 * 🔐 監査ミドルウェア
 * 全てのAPIアクセスとユーザーアクションを記録
 */

const { logUserAction, logAPICall, logSecurityEvent } = require('../utils/logger');

/**
 * 監査ログミドルウェア
 */
const auditLogger = (req, res, next) => {
  const startTime = Date.now();
  const originalSend = res.send;
  
  // レスポンスをインターセプトして記録
  res.send = function(body) {
    const responseTime = Date.now() - startTime;
    const statusCode = res.statusCode;
    
    // API呼び出しログ
    logAPICall(
      req.method,
      req.path,
      statusCode,
      responseTime,
      req.ip,
      req.user?.id || null
    );
    
    // ユーザーアクションログ
    if (req.user) {
      logUserAction(
        req.user.id,
        `${req.method} ${req.path}`,
        req.originalUrl,
        req.ip,
        req.headers['user-agent'],
        {
          statusCode,
          responseTime,
          bodySize: body ? Buffer.byteLength(body, 'utf8') : 0
        }
      );
    }
    
    // 異常なレスポンスタイムの記録
    if (responseTime > 5000) { // 5秒以上
      logSecurityEvent('SLOW_RESPONSE', {
        path: req.path,
        method: req.method,
        responseTime,
        ip: req.ip,
        userId: req.user?.id
      });
    }
    
    // 4xx, 5xxエラーの記録
    if (statusCode >= 400) {
      logSecurityEvent('ERROR_RESPONSE', {
        path: req.path,
        method: req.method,
        statusCode,
        ip: req.ip,
        userId: req.user?.id,
        userAgent: req.headers['user-agent']
      });
    }
    
    originalSend.call(this, body);
  };
  
  next();
};

/**
 * 重要なアクション専用監査ログ
 */
const criticalActionAudit = (action) => {
  return (req, res, next) => {
    const originalSend = res.send;
    
    res.send = function(body) {
      // 成功した重要なアクションのみ記録
      if (res.statusCode < 400) {
        logUserAction(
          req.user?.id || 'anonymous',
          action,
          req.originalUrl,
          req.ip,
          req.headers['user-agent'],
          {
            timestamp: new Date().toISOString(),
            requestBody: sanitizeForLog(req.body),
            success: true
          }
        );
      } else {
        // 失敗した重要なアクションはセキュリティイベントとして記録
        logSecurityEvent('CRITICAL_ACTION_FAILED', {
          action,
          path: req.path,
          statusCode: res.statusCode,
          ip: req.ip,
          userId: req.user?.id,
          requestBody: sanitizeForLog(req.body)
        });
      }
      
      originalSend.call(this, body);
    };
    
    next();
  };
};

/**
 * ファイルアクセス監査
 */
const fileAccessAudit = (req, res, next) => {
  const originalSend = res.send;
  
  res.send = function(body) {
    if (req.files && req.files.length > 0) {
      req.files.forEach((file, index) => {
        logUserAction(
          req.user?.id || 'anonymous',
          'FILE_UPLOAD',
          `${req.originalUrl}[${index}]`,
          req.ip,
          req.headers['user-agent'],
          {
            filename: file.originalname,
            mimetype: file.mimetype,
            size: file.size,
            fieldname: file.fieldname
          }
        );
      });
    }
    
    originalSend.call(this, body);
  };
  
  next();
};

/**
 * 認証関連の監査
 */
const authAudit = (req, res, next) => {
  const originalSend = res.send;
  
  res.send = function(body) {
    const isLoginAttempt = req.path.includes('/login');
    const isRegistration = req.path.includes('/register');
    const isLogout = req.path.includes('/logout');
    
    if (isLoginAttempt) {
      const success = res.statusCode === 200;
      logUserAction(
        req.body.username || 'unknown',
        'LOGIN_ATTEMPT',
        req.originalUrl,
        req.ip,
        req.headers['user-agent'],
        {
          success,
          statusCode: res.statusCode,
          timestamp: new Date().toISOString()
        }
      );
      
      if (!success) {
        logSecurityEvent('LOGIN_FAILED', {
          username: req.body.username,
          ip: req.ip,
          userAgent: req.headers['user-agent'],
          reason: res.statusCode === 401 ? 'Invalid credentials' : 'Other error'
        });
      }
    }
    
    if (isRegistration) {
      const success = res.statusCode === 201;
      logUserAction(
        req.body.username || 'unknown',
        'REGISTRATION_ATTEMPT',
        req.originalUrl,
        req.ip,
        req.headers['user-agent'],
        {
          success,
          email: req.body.email,
          statusCode: res.statusCode
        }
      );
    }
    
    if (isLogout) {
      logUserAction(
        req.user?.id || req.user?.username || 'unknown',
        'LOGOUT',
        req.originalUrl,
        req.ip,
        req.headers['user-agent'],
        {
          timestamp: new Date().toISOString()
        }
      );
    }
    
    originalSend.call(this, body);
  };
  
  next();
};

/**
 * データ変更操作の監査
 */
const dataChangeAudit = (operation) => {
  return (req, res, next) => {
    const originalSend = res.send;
    
    res.send = function(body) {
      if (res.statusCode < 400) {
        logUserAction(
          req.user?.id || 'system',
          operation,
          req.originalUrl,
          req.ip,
          req.headers['user-agent'],
          {
            resourceId: req.params.id || 'bulk',
            changes: sanitizeForLog(req.body),
            success: true
          }
        );
      }
      
      originalSend.call(this, body);
    };
    
    next();
  };
};

/**
 * セキュリティ制御の監査
 */
const securityControlAudit = (req, res, next) => {
  // セキュリティ関連のヘッダーチェック
  const securityHeaders = [
    'authorization',
    'x-csrf-token',
    'x-requested-with'
  ];
  
  const presentHeaders = securityHeaders.filter(header => req.headers[header]);
  const missingHeaders = securityHeaders.filter(header => !req.headers[header]);
  
  if (missingHeaders.length > 0 && req.method !== 'GET') {
    logSecurityEvent('MISSING_SECURITY_HEADERS', {
      missingHeaders,
      path: req.path,
      method: req.method,
      ip: req.ip,
      userAgent: req.headers['user-agent']
    });
  }
  
  next();
};

/**
 * ログ用データのサニタイゼーション
 */
const sanitizeForLog = (data) => {
  if (!data || typeof data !== 'object') {
    return data;
  }
  
  const sanitized = { ...data };
  
  // パスワードなどの機密情報を除去
  const sensitiveFields = ['password', 'token', 'secret', 'key', 'authorization'];
  
  for (const field of sensitiveFields) {
    if (sanitized[field]) {
      sanitized[field] = '[REDACTED]';
    }
  }
  
  return sanitized;
};

module.exports = {
  auditLogger,
  criticalActionAudit,
  fileAccessAudit,
  authAudit,
  dataChangeAudit,
  securityControlAudit
};