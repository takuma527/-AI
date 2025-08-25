/**
 * 🔐 認証ミドルウェア
 * JWT認証とアクセス制御
 */

const jwt = require('jsonwebtoken');
const User = require('../models/User');
const { logSecurityEvent, logAuthEvent } = require('../utils/logger');

/**
 * JWTトークン検証ミドルウェア
 */
const authMiddleware = async (req, res, next) => {
  try {
    const authHeader = req.headers.authorization;
    
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      logSecurityEvent('MISSING_AUTH_TOKEN', {
        path: req.path,
        method: req.method,
        ip: req.ip,
        userAgent: req.headers['user-agent']
      });
      
      return res.status(401).json({
        error: 'Access denied',
        message: 'No valid authorization token provided'
      });
    }
    
    const token = authHeader.substring(7);
    
    // トークン検証
    const decoded = jwt.verify(token, process.env.JWT_SECRET);
    
    // ユーザー存在確認
    const user = await User.findById(decoded.id).select('-password');
    
    if (!user || !user.isActive) {
      logSecurityEvent('INVALID_USER_TOKEN', {
        userId: decoded.id,
        path: req.path,
        ip: req.ip
      });
      
      return res.status(401).json({
        error: 'Access denied',
        message: 'Invalid or expired token'
      });
    }
    
    // アカウントロック状態チェック
    if (user.isLocked) {
      logSecurityEvent('ACCESS_ATTEMPT_LOCKED_ACCOUNT', {
        userId: user._id,
        username: user.username,
        path: req.path,
        ip: req.ip
      });
      
      return res.status(423).json({
        error: 'Account locked',
        message: 'Account is temporarily locked'
      });
    }
    
    req.user = user;
    next();
    
  } catch (error) {
    if (error.name === 'TokenExpiredError') {
      return res.status(401).json({
        error: 'Token expired',
        message: 'Access token has expired'
      });
    }
    
    if (error.name === 'JsonWebTokenError') {
      logSecurityEvent('INVALID_JWT_TOKEN', {
        error: error.message,
        path: req.path,
        ip: req.ip,
        userAgent: req.headers['user-agent']
      });
      
      return res.status(401).json({
        error: 'Invalid token',
        message: 'Invalid access token'
      });
    }
    
    return res.status(500).json({
      error: 'Authentication error',
      message: 'An error occurred during authentication'
    });
  }
};

/**
 * ロール別アクセス制御
 */
const requireRole = (roles) => {
  return (req, res, next) => {
    if (!req.user) {
      return res.status(401).json({
        error: 'Authentication required',
        message: 'Please log in to access this resource'
      });
    }
    
    const userRole = req.user.role;
    
    if (!roles.includes(userRole)) {
      logSecurityEvent('UNAUTHORIZED_ACCESS_ATTEMPT', {
        userId: req.user._id,
        username: req.user.username,
        requiredRoles: roles,
        userRole,
        path: req.path,
        ip: req.ip
      });
      
      return res.status(403).json({
        error: 'Access forbidden',
        message: 'Insufficient permissions to access this resource'
      });
    }
    
    next();
  };
};

/**
 * 管理者専用アクセス
 */
const requireAdmin = requireRole(['admin']);

/**
 * プレミアム以上のアクセス
 */
const requirePremium = requireRole(['premium', 'admin']);

/**
 * セッション検証ミドルウェア
 */
const sessionValidation = (req, res, next) => {
  // セッションが存在し、ユーザーIDが一致することを確認
  if (req.session && req.session.userId && req.user) {
    if (req.session.userId.toString() !== req.user._id.toString()) {
      logSecurityEvent('SESSION_USER_MISMATCH', {
        sessionUserId: req.session.userId,
        tokenUserId: req.user._id,
        ip: req.ip
      });
      
      return res.status(401).json({
        error: 'Session invalid',
        message: 'Session does not match authenticated user'
      });
    }
  }
  
  next();
};

/**
 * オプショナル認証ミドルウェア（ログインしていなくてもOK）
 */
const optionalAuth = async (req, res, next) => {
  try {
    const authHeader = req.headers.authorization;
    
    if (authHeader && authHeader.startsWith('Bearer ')) {
      const token = authHeader.substring(7);
      const decoded = jwt.verify(token, process.env.JWT_SECRET);
      const user = await User.findById(decoded.id).select('-password');
      
      if (user && user.isActive && !user.isLocked) {
        req.user = user;
      }
    }
    
    next();
  } catch (error) {
    // オプショナル認証なのでエラーは無視
    next();
  }
};

/**
 * API使用量制限チェック
 */
const usageLimitCheck = (req, res, next) => {
  if (!req.user) {
    return next();
  }
  
  if (!req.user.canAskQuestion()) {
    logSecurityEvent('USAGE_LIMIT_EXCEEDED', {
      userId: req.user._id,
      username: req.user.username,
      questionsAsked: req.user.usage.questionsAsked,
      dailyLimit: req.user.usage.dailyLimit,
      ip: req.ip
    });
    
    return res.status(429).json({
      error: 'Usage limit exceeded',
      message: 'Daily question limit reached',
      dailyLimit: req.user.usage.dailyLimit,
      questionsAsked: req.user.usage.questionsAsked,
      resetTime: new Date(req.user.usage.lastResetDate.getTime() + 24 * 60 * 60 * 1000)
    });
  }
  
  next();
};

/**
 * IP別追加制限
 */
const ipBasedLimit = (maxRequestsPerIP = 1000) => {
  const ipRequestCounts = new Map();
  
  return (req, res, next) => {
    const clientIP = req.ip;
    const currentTime = Date.now();
    const windowStart = currentTime - (24 * 60 * 60 * 1000); // 24時間
    
    // 古いエントリをクリーンアップ
    for (const [ip, data] of ipRequestCounts.entries()) {
      if (data.lastRequest < windowStart) {
        ipRequestCounts.delete(ip);
      }
    }
    
    const ipData = ipRequestCounts.get(clientIP) || { count: 0, lastRequest: currentTime };
    
    if (ipData.lastRequest < windowStart) {
      ipData.count = 1;
    } else {
      ipData.count += 1;
    }
    
    ipData.lastRequest = currentTime;
    ipRequestCounts.set(clientIP, ipData);
    
    if (ipData.count > maxRequestsPerIP) {
      logSecurityEvent('IP_RATE_LIMIT_EXCEEDED', {
        ip: clientIP,
        requestCount: ipData.count,
        maxAllowed: maxRequestsPerIP,
        path: req.path
      });
      
      return res.status(429).json({
        error: 'IP rate limit exceeded',
        message: 'Too many requests from this IP address'
      });
    }
    
    next();
  };
};

module.exports = {
  authMiddleware,
  requireRole,
  requireAdmin,
  requirePremium,
  sessionValidation,
  optionalAuth,
  usageLimitCheck,
  ipBasedLimit
};