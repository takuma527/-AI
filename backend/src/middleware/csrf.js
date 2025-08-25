/**
 * 🔐 CSRF保護ミドルウェア
 * Cross-Site Request Forgery対策
 */

const csrf = require('csrf');
const logger = require('../utils/logger');

// CSRFトークン生成器
const tokens = new csrf();

/**
 * CSRF保護ミドルウェア
 */
const csrfProtection = (req, res, next) => {
  // セーフメソッドはCSRF保護をスキップ
  if (['GET', 'HEAD', 'OPTIONS'].includes(req.method)) {
    return next();
  }
  
  // CSRFシークレットの生成または取得
  if (!req.session.csrfSecret) {
    req.session.csrfSecret = tokens.secretSync();
  }
  
  const secret = req.session.csrfSecret;
  const token = req.headers['x-csrf-token'] || 
                req.body._csrf || 
                req.query._csrf;
  
  // トークンの検証
  if (!token || !tokens.verify(secret, token)) {
    logger.warn(`CSRF token validation failed from IP ${req.ip}, User-Agent: ${req.headers['user-agent']}`);
    
    return res.status(403).json({
      error: 'Forbidden',
      message: 'Invalid CSRF token',
      code: 'CSRF_TOKEN_INVALID'
    });
  }
  
  next();
};

/**
 * CSRFトークンを生成して返すルート
 */
const generateCSRFToken = (req, res) => {
  // CSRFシークレットの生成または取得
  if (!req.session.csrfSecret) {
    req.session.csrfSecret = tokens.secretSync();
  }
  
  const token = tokens.create(req.session.csrfSecret);
  
  res.json({
    csrfToken: token,
    timestamp: Date.now()
  });
};

/**
 * SameSiteクッキーベースのCSRF対策
 */
const sameSiteProtection = (req, res, next) => {
  // SameSite属性がStrictに設定されているクッキーをチェック
  const sessionCookie = req.headers.cookie;
  
  if (sessionCookie && req.method !== 'GET') {
    // Refererヘッダーのチェック
    const referer = req.headers.referer;
    const origin = req.headers.origin;
    const host = req.headers.host;
    
    const allowedOrigins = process.env.ALLOWED_ORIGINS?.split(',') || [`http://localhost:3000`];
    const isValidOrigin = allowedOrigins.some(allowedOrigin => 
      referer?.startsWith(allowedOrigin) || origin === allowedOrigin
    );
    
    if (!isValidOrigin && !req.path.startsWith('/api/public')) {
      logger.warn(`Invalid origin/referer: origin=${origin}, referer=${referer} from IP ${req.ip}`);
      
      return res.status(403).json({
        error: 'Forbidden',
        message: 'Invalid request origin'
      });
    }
  }
  
  next();
};

/**
 * ダブルサブミット・クッキーパターン
 */
const doubleSubmitCookie = (req, res, next) => {
  if (['GET', 'HEAD', 'OPTIONS'].includes(req.method)) {
    return next();
  }
  
  // クッキーからCSRFトークンを取得
  const cookieToken = req.cookies.csrfToken;
  
  // ヘッダーまたはボディからCSRFトークンを取得
  const headerToken = req.headers['x-csrf-token'] || 
                     req.body._csrf;
  
  // 両方のトークンが存在し、一致することを確認
  if (!cookieToken || !headerToken || cookieToken !== headerToken) {
    logger.warn(`Double submit cookie validation failed from IP ${req.ip}`);
    
    return res.status(403).json({
      error: 'Forbidden',
      message: 'CSRF token mismatch'
    });
  }
  
  next();
};

/**
 * CSRFトークンをクッキーに設定
 */
const setCSRFCookie = (req, res, next) => {
  if (!req.session.csrfSecret) {
    req.session.csrfSecret = tokens.secretSync();
  }
  
  const token = tokens.create(req.session.csrfSecret);
  
  res.cookie('csrfToken', token, {
    httpOnly: false, // フロントエンドからアクセス可能
    secure: process.env.NODE_ENV === 'production',
    sameSite: 'strict',
    maxAge: 24 * 60 * 60 * 1000 // 24時間
  });
  
  // レスポンスヘッダーにも設定
  res.setHeader('X-CSRF-Token', token);
  
  next();
};

/**
 * カスタムヘッダーベースの検証
 */
const customHeaderProtection = (req, res, next) => {
  if (['GET', 'HEAD', 'OPTIONS'].includes(req.method)) {
    return next();
  }
  
  // カスタムヘッダーの存在をチェック
  const customHeader = req.headers['x-requested-with'];
  
  if (customHeader !== 'XMLHttpRequest' && !req.path.startsWith('/api/public')) {
    logger.warn(`Missing or invalid X-Requested-With header from IP ${req.ip}`);
    
    return res.status(400).json({
      error: 'Bad Request',
      message: 'Invalid request headers'
    });
  }
  
  next();
};

module.exports = {
  csrfProtection,
  generateCSRFToken,
  sameSiteProtection,
  doubleSubmitCookie,
  setCSRFCookie,
  customHeaderProtection
};