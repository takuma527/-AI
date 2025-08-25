/**
 * 🔐 ユーザーモデル
 * セキュリティを重視したユーザー認証・認可システム
 */

const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');

const userSchema = new mongoose.Schema({
  username: {
    type: String,
    required: [true, 'ユーザー名は必須です'],
    unique: true,
    trim: true,
    minlength: [3, 'ユーザー名は3文字以上である必要があります'],
    maxlength: [30, 'ユーザー名は30文字以下である必要があります'],
    match: [/^[a-zA-Z0-9_-]+$/, 'ユーザー名には英数字、アンダースコア、ハイフンのみ使用できます']
  },
  email: {
    type: String,
    required: [true, 'メールアドレスは必須です'],
    unique: true,
    trim: true,
    lowercase: true,
    match: [/^\w+([.-]?\w+)*@\w+([.-]?\w+)*(\.\w{2,3})+$/, '有効なメールアドレスを入力してください']
  },
  password: {
    type: String,
    required: [true, 'パスワードは必須です'],
    minlength: [8, 'パスワードは8文字以上である必要があります']
  },
  firstName: {
    type: String,
    trim: true,
    maxlength: [50, '名前は50文字以下である必要があります']
  },
  lastName: {
    type: String,
    trim: true,
    maxlength: [50, '苗字は50文字以下である必要があります']
  },
  role: {
    type: String,
    enum: ['user', 'premium', 'admin'],
    default: 'user'
  },
  isActive: {
    type: Boolean,
    default: true
  },
  isEmailVerified: {
    type: Boolean,
    default: false
  },
  emailVerificationToken: String,
  emailVerificationExpires: Date,
  passwordResetToken: String,
  passwordResetExpires: Date,
  
  // セキュリティ関連フィールド
  loginAttempts: {
    type: Number,
    default: 0
  },
  lockUntil: Date,
  lastLogin: Date,
  lastLoginIP: String,
  twoFactorSecret: String,
  twoFactorEnabled: {
    type: Boolean,
    default: false
  },
  
  // プロファイル設定
  preferences: {
    language: {
      type: String,
      default: 'ja',
      enum: ['ja', 'en']
    },
    theme: {
      type: String,
      default: 'light',
      enum: ['light', 'dark']
    },
    notifications: {
      email: { type: Boolean, default: true },
      push: { type: Boolean, default: true }
    }
  },
  
  // 使用量追跡
  usage: {
    questionsAsked: { type: Number, default: 0 },
    dailyLimit: { type: Number, default: 100 },
    monthlyLimit: { type: Number, default: 3000 },
    lastResetDate: { type: Date, default: Date.now }
  },
  
  // チャット履歴の参照
  conversations: [{
    type: mongoose.Schema.Types.ObjectId,
    ref: 'Conversation'
  }]
}, {
  timestamps: true
});

// アカウントロック状態の仮想プロパティ
userSchema.virtual('isLocked').get(function() {
  return !!(this.lockUntil && this.lockUntil > Date.now());
});

// パスワードハッシュ化ミドルウェア
userSchema.pre('save', async function(next) {
  if (!this.isModified('password')) return next();
  
  try {
    const saltRounds = parseInt(process.env.BCRYPT_ROUNDS) || 12;
    this.password = await bcrypt.hash(this.password, saltRounds);
    next();
  } catch (error) {
    next(error);
  }
});

// パスワード検証メソッド
userSchema.methods.comparePassword = async function(candidatePassword) {
  return await bcrypt.compare(candidatePassword, this.password);
};

// JWTトークン生成メソッド
userSchema.methods.generateAccessToken = function() {
  return jwt.sign(
    { 
      id: this._id,
      username: this.username,
      role: this.role 
    },
    process.env.JWT_SECRET,
    { 
      expiresIn: process.env.JWT_EXPIRES_IN || '15m',
      issuer: 'excel-chatbot',
      audience: 'excel-chatbot-users'
    }
  );
};

// リフレッシュトークン生成メソッド
userSchema.methods.generateRefreshToken = function() {
  return jwt.sign(
    { 
      id: this._id,
      type: 'refresh'
    },
    process.env.JWT_REFRESH_SECRET,
    { 
      expiresIn: '7d',
      issuer: 'excel-chatbot',
      audience: 'excel-chatbot-users'
    }
  );
};

// ログイン試行回数の増加
userSchema.methods.incLoginAttempts = async function() {
  // ロックが期限切れの場合、リセット
  if (this.lockUntil && this.lockUntil < Date.now()) {
    return this.updateOne({
      $unset: { lockUntil: 1 },
      $set: { loginAttempts: 1 }
    });
  }
  
  const updates = { $inc: { loginAttempts: 1 } };
  
  // 5回失敗でアカウントロック（2時間）
  if (this.loginAttempts + 1 >= 5 && !this.isLocked) {
    updates.$set = { lockUntil: Date.now() + 2 * 60 * 60 * 1000 }; // 2時間
  }
  
  return this.updateOne(updates);
};

// ログイン成功時のリセット
userSchema.methods.resetLoginAttempts = async function() {
  return this.updateOne({
    $unset: { loginAttempts: 1, lockUntil: 1 },
    $set: { 
      lastLogin: Date.now(),
      lastLoginIP: this.currentIP 
    }
  });
};

// 使用量の追跡
userSchema.methods.incrementUsage = async function() {
  const today = new Date();
  const lastReset = new Date(this.usage.lastResetDate);
  
  // 日付が変わった場合、使用量をリセット
  if (today.toDateString() !== lastReset.toDateString()) {
    this.usage.questionsAsked = 0;
    this.usage.lastResetDate = today;
  }
  
  this.usage.questionsAsked += 1;
  return this.save();
};

// 使用制限チェック
userSchema.methods.canAskQuestion = function() {
  const limit = this.role === 'premium' ? this.usage.dailyLimit * 5 : this.usage.dailyLimit;
  return this.usage.questionsAsked < limit;
};

// メール検証トークン生成
userSchema.methods.generateEmailVerificationToken = function() {
  const token = require('crypto').randomBytes(32).toString('hex');
  this.emailVerificationToken = token;
  this.emailVerificationExpires = Date.now() + 24 * 60 * 60 * 1000; // 24時間
  return token;
};

// パスワードリセットトークン生成
userSchema.methods.generatePasswordResetToken = function() {
  const token = require('crypto').randomBytes(32).toString('hex');
  this.passwordResetToken = token;
  this.passwordResetExpires = Date.now() + 60 * 60 * 1000; // 1時間
  return token;
};

// インデックス設定
userSchema.index({ email: 1 });
userSchema.index({ username: 1 });
userSchema.index({ emailVerificationToken: 1 });
userSchema.index({ passwordResetToken: 1 });

module.exports = mongoose.model('User', userSchema);