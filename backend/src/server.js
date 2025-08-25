/**
 * 🔐 セキュア Excel チャットボット サーバー
 * セキュリティを重視したExpress.jsアプリケーション
 */

require('dotenv').config();
const express = require('express');
const helmet = require('helmet');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const slowDown = require('express-slow-down');
const compression = require('compression');
const cookieParser = require('cookie-parser');
const session = require('express-session');
const MongoStore = require('connect-mongo');
const { createServer } = require('http');
const { Server } = require('socket.io');

// セキュリティミドルウェア
const { securityHeaders } = require('./middleware/security');
const { csrfProtection } = require('./middleware/csrf');
const { auditLogger } = require('./middleware/audit');
const { inputValidation } = require('./middleware/validation');

// ルーター
const authRoutes = require('./routes/auth');
const chatRoutes = require('./routes/chat');
const excelRoutes = require('./routes/excel');

// ユーティリティ
const logger = require('./utils/logger');
const { connectDB } = require('./utils/database');

const app = express();
const server = createServer(app);

// WebSocket設定
const io = new Server(server, {
  cors: {
    origin: process.env.ALLOWED_ORIGINS?.split(',') || ['http://localhost:3000'],
    credentials: true
  }
});

const PORT = process.env.PORT || 5000;

// 🔒 基本セキュリティ設定
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      styleSrc: ["'self'", "'unsafe-inline'"],
      scriptSrc: ["'self'"],
      imgSrc: ["'self'", "data:", "https:"],
      connectSrc: ["'self'", "wss:", "ws:"],
      fontSrc: ["'self'"],
      objectSrc: ["'none'"],
      mediaSrc: ["'self'"],
      frameSrc: ["'none'"],
    },
  },
  crossOriginEmbedderPolicy: false
}));

// CORS設定
app.use(cors({
  origin: function (origin, callback) {
    const allowedOrigins = process.env.ALLOWED_ORIGINS?.split(',') || ['http://localhost:3000'];
    if (!origin || allowedOrigins.includes(origin)) {
      callback(null, true);
    } else {
      callback(new Error('CORS policy violation'));
    }
  },
  credentials: true,
  optionsSuccessStatus: 200
}));

// 圧縮
app.use(compression());

// レート制限
const limiter = rateLimit({
  windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS) || 15 * 60 * 1000, // 15分
  max: parseInt(process.env.RATE_LIMIT_MAX_REQUESTS) || 100,
  message: {
    error: 'Too many requests from this IP, please try again later.',
    code: 'RATE_LIMIT_EXCEEDED'
  },
  standardHeaders: true,
  legacyHeaders: false,
  skip: (req) => {
    // ヘルスチェックはスキップ
    return req.path === '/health';
  }
});

// スローダウン（段階的制限）
const speedLimiter = slowDown({
  windowMs: 15 * 60 * 1000, // 15分
  delayAfter: 50, // 50リクエスト後に遅延開始
  delayMs: 500 // 500ms遅延
});

app.use(limiter);
app.use(speedLimiter);

// Body parser
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(cookieParser());

// セッション設定
app.use(session({
  secret: process.env.SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create({
    mongoUrl: process.env.MONGODB_URI,
    touchAfter: 24 * 3600 // 24時間
  }),
  cookie: {
    secure: process.env.NODE_ENV === 'production', // HTTPSでのみ
    httpOnly: true, // XSS対策
    maxAge: parseInt(process.env.SESSION_MAX_AGE) || 24 * 60 * 60 * 1000, // 24時間
    sameSite: 'strict' // CSRF対策
  },
  name: 'sessionId' // デフォルト名を変更
}));

// カスタムセキュリティミドルウェア
app.use(securityHeaders);
app.use(auditLogger);
app.use(inputValidation);

// CSRF保護（APIルートの前に設定）
app.use('/api', csrfProtection);

// ヘルスチェック
app.get('/health', (req, res) => {
  res.json({ 
    status: 'OK', 
    timestamp: new Date().toISOString(),
    uptime: process.uptime()
  });
});

// API ルート
app.use('/api/auth', authRoutes);
app.use('/api/chat', chatRoutes);
app.use('/api/excel', excelRoutes);

// 静的ファイル（本番環境）
if (process.env.NODE_ENV === 'production') {
  app.use(express.static(path.join(__dirname, '../../frontend/build')));
  
  app.get('*', (req, res) => {
    res.sendFile(path.resolve(__dirname, '../../frontend/build', 'index.html'));
  });
}

// エラーハンドリング
app.use((err, req, res, next) => {
  logger.error('Unhandled error:', err);
  
  // セキュリティ：本番環境では詳細なエラー情報を隠す
  if (process.env.NODE_ENV === 'production') {
    res.status(err.status || 500).json({
      error: 'Internal Server Error',
      message: 'An error occurred processing your request'
    });
  } else {
    res.status(err.status || 500).json({
      error: err.message,
      stack: err.stack
    });
  }
});

// 404ハンドラー
app.use((req, res) => {
  res.status(404).json({
    error: 'Not Found',
    message: 'The requested resource was not found'
  });
});

// WebSocket接続処理
require('./services/socketService')(io);

// サーバー起動
async function startServer() {
  try {
    // データベース接続
    await connectDB();
    
    server.listen(PORT, () => {
      logger.info(`🔐 セキュアExcelチャットボット サーバー起動: http://localhost:${PORT}`);
      logger.info(`🛡️ セキュリティ機能: 認証, レート制限, CSRF保護, XSS防止`);
      logger.info(`📊 Excel専門AI: 準備完了`);
    });
  } catch (error) {
    logger.error('サーバー起動エラー:', error);
    process.exit(1);
  }
}

// Graceful shutdown
process.on('SIGTERM', () => {
  logger.info('SIGTERM received. Shutting down gracefully...');
  server.close(() => {
    logger.info('Process terminated');
    process.exit(0);
  });
});

process.on('SIGINT', () => {
  logger.info('SIGINT received. Shutting down gracefully...');
  server.close(() => {
    logger.info('Process terminated');
    process.exit(0);
  });
});

startServer();

module.exports = app;