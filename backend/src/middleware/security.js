/**
 * 🔐 セキュリティミドルウェア
 * 追加のセキュリティヘッダーとチェック機能
 */

const logger = require('../utils/logger');

/**
 * セキュリティヘッダーの設定
 */
const securityHeaders = (req, res, next) => {
  // X-Frame-Options (既にhelmetで設定されているが念のため)
  res.setHeader('X-Frame-Options', 'DENY');
  
  // X-Content-Type-Options
  res.setHeader('X-Content-Type-Options', 'nosniff');
  
  // Referrer-Policy
  res.setHeader('Referrer-Policy', 'strict-origin-when-cross-origin');
  
  // Permissions Policy
  res.setHeader('Permissions-Policy', 'geolocation=(), microphone=(), camera=()');
  
  // X-XSS-Protection (古いブラウザ用)
  res.setHeader('X-XSS-Protection', '1; mode=block');
  
  // Custom security header
  res.setHeader('X-Security-Policy', 'Excel-Chatbot-Secure');
  
  next();
};

/**
 * リクエストサイズ制限チェック
 */
const requestSizeLimit = (maxSize = 10 * 1024 * 1024) => { // 10MB
  return (req, res, next) => {
    const contentLength = parseInt(req.headers['content-length']);
    
    if (contentLength && contentLength > maxSize) {
      logger.warn(`Request size limit exceeded: ${contentLength} bytes from IP ${req.ip}`);
      return res.status(413).json({
        error: 'Payload Too Large',
        message: 'Request size exceeds maximum allowed size'
      });
    }
    
    next();
  };
};

/**
 * 疑わしいヘッダーのチェック
 */
const suspiciousHeaderCheck = (req, res, next) => {
  const suspiciousHeaders = [
    'x-forwarded-host',
    'x-forwarded-server',
    'x-real-ip'
  ];
  
  // プロキシ経由でない場合、これらのヘッダーは疑わしい
  if (!req.headers['x-forwarded-for']) {
    for (const header of suspiciousHeaders) {
      if (req.headers[header]) {
        logger.warn(`Suspicious header detected: ${header} from IP ${req.ip}`);
      }
    }
  }
  
  next();
};

/**
 * User-Agentの検証
 */
const userAgentValidation = (req, res, next) => {
  const userAgent = req.headers['user-agent'];
  
  // User-Agentが存在しない場合
  if (!userAgent) {
    logger.warn(`Missing User-Agent from IP ${req.ip}`);
    return res.status(400).json({
      error: 'Bad Request',
      message: 'User-Agent header is required'
    });
  }
  
  // 明らかに自動化ツールのUser-Agent
  const suspiciousPatterns = [
    /bot/i,
    /crawler/i,
    /spider/i,
    /scraper/i,
    /python/i,
    /curl/i,
    /wget/i,
    /postman/i
  ];
  
  const isSuspicious = suspiciousPatterns.some(pattern => pattern.test(userAgent));
  
  if (isSuspicious && !req.path.startsWith('/api/public')) {
    logger.warn(`Suspicious User-Agent: ${userAgent} from IP ${req.ip}`);
    // 完全にブロックするのではなく、より厳しいレート制限を適用
    req.suspiciousAgent = true;
  }
  
  next();
};

/**
 * IPホワイトリスト（管理者機能用）
 */
const ipWhitelist = (allowedIPs = []) => {
  return (req, res, next) => {
    if (allowedIPs.length === 0) {
      return next(); // ホワイトリストが空の場合はスキップ
    }
    
    const clientIP = req.ip || req.connection.remoteAddress;
    
    if (!allowedIPs.includes(clientIP)) {
      logger.warn(`Access denied for IP: ${clientIP}`);
      return res.status(403).json({
        error: 'Forbidden',
        message: 'Access denied from this IP address'
      });
    }
    
    next();
  };
};

/**
 * セキュリティイベントの検出と記録
 */
const securityEventDetector = (req, res, next) => {
  // SQLインジェクション試行の検出
  const sqlInjectionPatterns = [
    /(\bOR\b|\bAND\b).*=.*=/i,
    /UNION.*SELECT/i,
    /DROP.*TABLE/i,
    /INSERT.*INTO/i,
    /DELETE.*FROM/i,
    /UPDATE.*SET/i,
    /'.*OR.*'.*='/i,
    /--/,
    /\/\*.*\*\//
  ];
  
  // XSS試行の検出
  const xssPatterns = [
    /<script[^>]*>.*?<\/script>/gi,
    /javascript:/gi,
    /onload\s*=/gi,
    /onclick\s*=/gi,
    /onerror\s*=/gi,
    /<iframe[^>]*>.*?<\/iframe>/gi
  ];
  
  // パストラバーサル試行の検出
  const pathTraversalPatterns = [
    /\.\.\//g,
    /\.\.\\\/g,
    /%2e%2e%2f/gi,
    /%2e%2e\\/gi
  ];
  
  const checkPatterns = (text, patterns, type) => {
    if (typeof text !== 'string') return false;
    
    return patterns.some(pattern => {
      if (pattern.test(text)) {
        logger.warn(`${type} attempt detected from IP ${req.ip}: ${text}`);
        return true;
      }
      return false;
    });
  };
  
  // URL、クエリパラメータ、ボディをチェック
  const textToCheck = [
    req.url,
    JSON.stringify(req.query),
    JSON.stringify(req.body)
  ].join(' ');
  
  if (checkPatterns(textToCheck, sqlInjectionPatterns, 'SQL Injection') ||
      checkPatterns(textToCheck, xssPatterns, 'XSS') ||
      checkPatterns(textToCheck, pathTraversalPatterns, 'Path Traversal')) {
    
    return res.status(400).json({
      error: 'Bad Request',
      message: 'Malicious content detected'
    });
  }
  
  next();
};

module.exports = {
  securityHeaders,
  requestSizeLimit,
  suspiciousHeaderCheck,
  userAgentValidation,
  ipWhitelist,
  securityEventDetector
};