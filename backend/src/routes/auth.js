/**
 * 🔐 認証関連ルート
 * セキュアなユーザー認証・認可システム
 */

const express = require('express');
const rateLimit = require('express-rate-limit');
const User = require('../models/User');
const { authAudit } = require('../middleware/audit');
const { 
  userRegistrationValidation, 
  loginValidation, 
  handleValidationErrors 
} = require('../middleware/validation');
const { generateCSRFToken } = require('../middleware/csrf');
const { logAuthEvent, logSecurityEvent } = require('../utils/logger');

const router = express.Router();

// 認証試行のレート制限（より厳しい制限）
const authLimiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15分
  max: 5, // 15分間で5回まで
  message: {
    error: 'Too many authentication attempts, please try again later',
    code: 'AUTH_RATE_LIMIT_EXCEEDED'
  },
  standardHeaders: true,
  legacyHeaders: false,
  keyGenerator: (req) => {
    return `auth:${req.ip}:${req.body.username || req.body.email}`;
  }
});

// パスワードリセットのレート制限
const passwordResetLimiter = rateLimit({
  windowMs: 60 * 60 * 1000, // 1時間
  max: 3, // 1時間に3回まで
  message: {
    error: 'Too many password reset attempts, please try again later',
    code: 'PASSWORD_RESET_LIMIT_EXCEEDED'
  }
});

/**
 * CSRFトークン取得
 */
router.get('/csrf-token', generateCSRFToken);

/**
 * ユーザー登録
 */
router.post('/register', 
  authLimiter,
  userRegistrationValidation,
  handleValidationErrors,
  authAudit,
  async (req, res) => {
    try {
      const { username, email, password, firstName, lastName } = req.body;
      
      // 既存ユーザーチェック
      const existingUser = await User.findOne({
        $or: [{ email }, { username }]
      });
      
      if (existingUser) {
        logSecurityEvent('DUPLICATE_REGISTRATION_ATTEMPT', {
          attemptedEmail: email,
          attemptedUsername: username,
          existingField: existingUser.email === email ? 'email' : 'username',
          ip: req.ip
        });
        
        return res.status(409).json({
          error: 'User already exists',
          message: 'A user with this email or username already exists'
        });
      }
      
      // 新しいユーザー作成
      const user = new User({
        username,
        email,
        password,
        firstName,
        lastName
      });
      
      // メール検証トークン生成
      const verificationToken = user.generateEmailVerificationToken();
      
      await user.save();
      
      logAuthEvent('USER_REGISTERED', username, req.ip, req.headers['user-agent'], true);
      
      // パスワードを除いてレスポンス
      const { password: _, ...userResponse } = user.toObject();
      
      res.status(201).json({
        message: 'User registered successfully',
        user: userResponse,
        verificationRequired: true
      });
      
    } catch (error) {
      logAuthEvent('REGISTRATION_ERROR', req.body.username, req.ip, req.headers['user-agent'], false, error.message);
      
      res.status(500).json({
        error: 'Registration failed',
        message: 'An error occurred during registration'
      });
    }
  }
);

/**
 * ユーザーログイン
 */
router.post('/login',
  authLimiter,
  loginValidation,
  handleValidationErrors,
  authAudit,
  async (req, res) => {
    try {
      const { username, password } = req.body;
      
      // ユーザー検索（usernameまたはemail）
      const user = await User.findOne({
        $or: [{ username }, { email: username }],
        isActive: true
      });
      
      if (!user) {
        logAuthEvent('LOGIN_FAILED', username, req.ip, req.headers['user-agent'], false, 'User not found');
        
        return res.status(401).json({
          error: 'Authentication failed',
          message: 'Invalid credentials'
        });
      }
      
      // アカウントロック状態チェック
      if (user.isLocked) {
        logSecurityEvent('LOGIN_ATTEMPT_LOCKED_ACCOUNT', {
          username,
          userId: user._id,
          ip: req.ip,
          lockUntil: user.lockUntil
        });
        
        return res.status(423).json({
          error: 'Account locked',
          message: 'Account is temporarily locked due to multiple failed login attempts'
        });
      }
      
      // パスワード検証
      const isValidPassword = await user.comparePassword(password);
      
      if (!isValidPassword) {
        await user.incLoginAttempts();
        logAuthEvent('LOGIN_FAILED', username, req.ip, req.headers['user-agent'], false, 'Invalid password');
        
        return res.status(401).json({
          error: 'Authentication failed',
          message: 'Invalid credentials'
        });
      }
      
      // メール検証チェック
      if (!user.isEmailVerified) {
        return res.status(403).json({
          error: 'Email not verified',
          message: 'Please verify your email address before logging in'
        });
      }
      
      // ログイン成功処理
      user.currentIP = req.ip;
      await user.resetLoginAttempts();
      
      // トークン生成
      const accessToken = user.generateAccessToken();
      const refreshToken = user.generateRefreshToken();
      
      // セッション設定
      req.session.userId = user._id;
      req.session.username = user.username;
      req.session.role = user.role;
      
      logAuthEvent('LOGIN_SUCCESS', username, req.ip, req.headers['user-agent'], true);
      
      // セキュアクッキーとしてリフレッシュトークンを設定
      res.cookie('refreshToken', refreshToken, {
        httpOnly: true,
        secure: process.env.NODE_ENV === 'production',
        sameSite: 'strict',
        maxAge: 7 * 24 * 60 * 60 * 1000 // 7日
      });
      
      const { password: _, ...userResponse } = user.toObject();
      
      res.json({
        message: 'Login successful',
        user: userResponse,
        accessToken,
        expiresIn: '15m'
      });
      
    } catch (error) {
      logAuthEvent('LOGIN_ERROR', req.body.username, req.ip, req.headers['user-agent'], false, error.message);
      
      res.status(500).json({
        error: 'Login failed',
        message: 'An error occurred during login'
      });
    }
  }
);

/**
 * ログアウト
 */
router.post('/logout', authAudit, (req, res) => {
  try {
    // セッション破棄
    req.session.destroy((err) => {
      if (err) {
        return res.status(500).json({
          error: 'Logout failed',
          message: 'Failed to destroy session'
        });
      }
      
      // リフレッシュトークンクッキー削除
      res.clearCookie('refreshToken');
      res.clearCookie('sessionId');
      
      logAuthEvent('LOGOUT_SUCCESS', req.user?.username || 'unknown', req.ip, req.headers['user-agent'], true);
      
      res.json({
        message: 'Logout successful'
      });
    });
  } catch (error) {
    res.status(500).json({
      error: 'Logout failed',
      message: 'An error occurred during logout'
    });
  }
});

/**
 * トークンリフレッシュ
 */
router.post('/refresh-token', async (req, res) => {
  try {
    const refreshToken = req.cookies.refreshToken;
    
    if (!refreshToken) {
      return res.status(401).json({
        error: 'No refresh token',
        message: 'Refresh token is required'
      });
    }
    
    // トークン検証
    const decoded = jwt.verify(refreshToken, process.env.JWT_REFRESH_SECRET);
    
    const user = await User.findById(decoded.id);
    if (!user || !user.isActive) {
      return res.status(401).json({
        error: 'Invalid refresh token',
        message: 'User not found or inactive'
      });
    }
    
    // 新しいアクセストークン生成
    const newAccessToken = user.generateAccessToken();
    
    res.json({
      accessToken: newAccessToken,
      expiresIn: '15m'
    });
    
  } catch (error) {
    res.status(401).json({
      error: 'Invalid refresh token',
      message: 'Refresh token is invalid or expired'
    });
  }
});

/**
 * パスワードリセット要求
 */
router.post('/forgot-password',
  passwordResetLimiter,
  async (req, res) => {
    try {
      const { email } = req.body;
      
      const user = await User.findOne({ email, isActive: true });
      
      if (!user) {
        // セキュリティ：存在しないメールでも成功レスポンス
        return res.json({
          message: 'Password reset instructions sent to email if account exists'
        });
      }
      
      const resetToken = user.generatePasswordResetToken();
      await user.save();
      
      logSecurityEvent('PASSWORD_RESET_REQUESTED', {
        email,
        userId: user._id,
        ip: req.ip
      });
      
      // ここで実際のメール送信処理を実装
      // await sendPasswordResetEmail(user.email, resetToken);
      
      res.json({
        message: 'Password reset instructions sent to email if account exists'
      });
      
    } catch (error) {
      res.status(500).json({
        error: 'Password reset failed',
        message: 'An error occurred while processing password reset request'
      });
    }
  }
);

/**
 * メール検証
 */
router.get('/verify-email/:token', async (req, res) => {
  try {
    const { token } = req.params;
    
    const user = await User.findOne({
      emailVerificationToken: token,
      emailVerificationExpires: { $gt: Date.now() }
    });
    
    if (!user) {
      return res.status(400).json({
        error: 'Invalid or expired token',
        message: 'Email verification token is invalid or expired'
      });
    }
    
    user.isEmailVerified = true;
    user.emailVerificationToken = undefined;
    user.emailVerificationExpires = undefined;
    
    await user.save();
    
    logAuthEvent('EMAIL_VERIFIED', user.username, req.ip, req.headers['user-agent'], true);
    
    res.json({
      message: 'Email verified successfully'
    });
    
  } catch (error) {
    res.status(500).json({
      error: 'Email verification failed',
      message: 'An error occurred during email verification'
    });
  }
});

module.exports = router;