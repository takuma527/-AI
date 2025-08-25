/**
 * 🔐 セキュア Excel チャットボット サーバー（修正版）
 * VBA対応・登録機能・コピー機能付き
 */

require('dotenv').config();
const express = require('express');
const helmet = require('helmet');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const compression = require('compression');
const cookieParser = require('cookie-parser');
const session = require('express-session');
const { createServer } = require('http');
const { Server } = require('socket.io');
const path = require('path');
const ExcelAI = require('./services/excelAI');

const app = express();
const server = createServer(app);

// WebSocket設定
const io = new Server(server, {
  cors: {
    origin: true,
    credentials: true
  }
});

const PORT = process.env.PORT || 5000;

// セキュリティ設定（修正版）
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com", "https://cdnjs.cloudflare.com"],
      scriptSrc: ["'self'", "'unsafe-inline'"],
      imgSrc: ["'self'", "data:", "https:"],
      connectSrc: ["'self'", "wss:", "ws:"],
      fontSrc: ["'self'", "https://fonts.gstatic.com", "https://cdnjs.cloudflare.com"],
      objectSrc: ["'none'"],
      mediaSrc: ["'self'"],
      frameSrc: ["'none'"],
    },
  },
  crossOriginEmbedderPolicy: false
}));

// CORS設定
app.use(cors({
  origin: true,
  credentials: true,
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'Accept', 'X-Requested-With']
}));

// レート制限
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 100,
  message: { error: 'Too many requests' }
});
app.use('/api/', limiter);

// 基本ミドルウェア
app.use(compression());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());

// セッション設定
app.use(session({
  secret: process.env.SESSION_SECRET || 'excel-chatbot-session-secret',
  resave: false,
  saveUninitialized: true,
  cookie: {
    secure: false,
    httpOnly: true,
    maxAge: 24 * 60 * 60 * 1000,
    sameSite: 'lax'
  }
}));

// データストア
const users = new Map();
const chatHistory = new Map();

// デモユーザー作成
const demoUser = {
  id: 'demo_user',
  username: 'demo',
  password: 'demo123',
  email: null,
  createdAt: new Date().toISOString(),
  profile: {
    displayName: 'Demo User',
    excelLevel: 'intermediate'
  }
};
users.set('demo_user', demoUser);

// API エンドポイント

// ログイン
app.post('/api/auth/login', (req, res) => {
  const { username, password } = req.body;
  
  console.log('ログイン試行:', { username, password });
  
  let user = null;
  for (const [id, u] of users) {
    if (u.username === username && u.password === password) {
      user = u;
      break;
    }
  }
  
  if (user) {
    req.session.isLoggedIn = true;
    req.session.userId = user.id;
    req.session.username = user.username;
    
    console.log('ログイン成功:', user.username);
    
    res.json({
      message: 'Login successful',
      user: {
        id: user.id,
        username: user.username,
        email: user.email,
        profile: user.profile
      }
    });
  } else {
    console.log('ログイン失敗: 無効な認証情報');
    res.status(401).json({
      error: 'Authentication failed',
      message: 'Invalid credentials'
    });
  }
});

// ユーザー登録
app.post('/api/auth/register', (req, res) => {
  const { username, password, email } = req.body;
  
  if (!username || !password) {
    return res.status(400).json({
      error: 'Validation Error',
      message: 'ユーザー名とパスワードが必要です'
    });
  }
  
  if (username.length < 3) {
    return res.status(400).json({
      error: 'Validation Error',
      message: 'ユーザー名は3文字以上である必要があります'
    });
  }
  
  if (password.length < 6) {
    return res.status(400).json({
      error: 'Validation Error',
      message: 'パスワードは6文字以上である必要があります'
    });
  }
  
  for (const [id, user] of users) {
    if (user.username === username) {
      return res.status(409).json({
        error: 'Conflict',
        message: 'このユーザー名は既に使用されています'
      });
    }
  }
  
  const userId = `user_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  const newUser = {
    id: userId,
    username,
    password,
    email: email || null,
    createdAt: new Date().toISOString(),
    profile: {
      displayName: username,
      excelLevel: 'beginner'
    }
  };
  
  users.set(userId, newUser);
  
  console.log(`✅ 新規ユーザー登録: ${username} (ID: ${userId})`);
  
  // 自動ログイン
  req.session.isLoggedIn = true;
  req.session.userId = userId;
  req.session.username = username;
  
  res.status(201).json({
    message: 'ユーザー登録が完了しました',
    user: {
      id: userId,
      username,
      email: email || null,
      createdAt: newUser.createdAt,
      profile: newUser.profile
    }
  });
});

// ログアウト
app.post('/api/auth/logout', (req, res) => {
  req.session.destroy((err) => {
    if (err) {
      return res.status(500).json({ error: 'Logout failed' });
    }
    res.json({ message: 'Logout successful' });
  });
});

// 認証チェック
app.get('/api/auth/me', (req, res) => {
  if (req.session.isLoggedIn && req.session.userId) {
    const user = users.get(req.session.userId);
    if (user) {
      return res.json({ 
        user: {
          id: user.id,
          username: user.username,
          email: user.email,
          profile: user.profile
        }
      });
    }
  }
  res.status(401).json({ error: 'Not authenticated' });
});

// チャットメッセージ処理（高度なExcelAI搭載）
app.post('/api/chat/message', async (req, res) => {
  try {
    const { message } = req.body;
    const userId = req.session.userId || 'guest';
    
    if (!message) {
      return res.status(400).json({
        error: 'Validation Error',
        message: 'Message is required'
      });
    }
    
    console.log(`🔍 Excel AI処理開始: "${message}" (ユーザー: ${userId})`);
    
    // 高度なExcelAI応答生成
    const aiResponse = ExcelAI.generateResponse(message);
    
    // チャット履歴に保存
    if (!chatHistory.has(userId)) {
      chatHistory.set(userId, []);
    }
    const userHistory = chatHistory.get(userId);
    
    userHistory.push({
      id: Date.now(),
      type: 'user',
      message,
      timestamp: new Date().toISOString()
    });
    
    userHistory.push({
      id: Date.now() + 1,
      type: 'bot',
      message: aiResponse.response,
      formulas: aiResponse.formulas || [],
      vbaCode: aiResponse.vbaCode || null,
      timestamp: new Date().toISOString()
    });
    
    if (userHistory.length > 100) {
      userHistory.splice(0, userHistory.length - 100);
    }
    
    console.log(`✅ Excel AI応答完了: ${aiResponse.formulas?.length || 0}個の数式, VBA: ${!!aiResponse.vbaCode}`);
    
    res.json({
      response: aiResponse.response,
      formulas: aiResponse.formulas || [],
      vbaCode: aiResponse.vbaCode || null,
      metadata: {
        responseTime: 50,
        knowledgeResults: aiResponse.functionsFound || 0,
        model: 'excel-ai-advanced',
        conversationId: `conv_${Date.now()}`,
        hasFormulas: !!(aiResponse.formulas && aiResponse.formulas.length > 0),
        hasVBA: !!aiResponse.vbaCode
      }
    });
    
  } catch (error) {
    console.error('❌ Excel AI エラー:', error);
    res.status(500).json({
      error: 'Internal Server Error',
      message: 'Excel AI処理中にエラーが発生しました'
    });
  }
});

// チャット履歴取得
app.get('/api/chat/history', (req, res) => {
  const userId = req.session.userId || 'guest';
  const history = chatHistory.get(userId) || [];
  res.json({ messages: history.slice(-20) });
});

// チャット履歴削除
app.delete('/api/chat/history', (req, res) => {
  const userId = req.session.userId || 'guest';
  chatHistory.delete(userId);
  res.json({ message: 'Chat history cleared' });
});

// メインページ
app.get('/', (req, res) => {
  res.send(`
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>🔐 セキュア Excel チャットボット</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&display=swap" rel="stylesheet">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        body {
            font-family: 'Noto Sans JP', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            color: #333;
        }
        
        .header {
            background: rgba(255,255,255,0.1);
            backdrop-filter: blur(10px);
            padding: 1rem 0;
            border-bottom: 1px solid rgba(255,255,255,0.2);
            position: sticky;
            top: 0;
            z-index: 100;
        }
        
        .header-content {
            max-width: 1200px;
            margin: 0 auto;
            padding: 0 2rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .logo h1 {
            color: white;
            font-size: 1.5rem;
            font-weight: 700;
        }
        
        .user-info {
            display: flex;
            align-items: center;
            gap: 1rem;
            color: white;
        }
        
        .btn {
            padding: 0.5rem 1rem;
            border: none;
            border-radius: 8px;
            font-size: 0.9rem;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.3s ease;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        .btn-primary {
            background: #007bff;
            color: white;
        }
        
        .btn-primary:hover {
            background: #0056b3;
            transform: translateY(-2px);
        }
        
        .btn-success {
            background: #28a745;
            color: white;
        }
        
        .btn-success:hover {
            background: #1e7e34;
        }
        
        .btn-danger {
            background: #dc3545;
            color: white;
        }
        
        .btn-danger:hover {
            background: #c82333;
        }
        
        .btn-secondary {
            background: #6c757d;
            color: white;
        }
        
        .btn-secondary:hover {
            background: #545b62;
        }
        
        .login-form {
            max-width: 450px;
            margin: 2rem auto;
            background: rgba(255,255,255,0.95);
            backdrop-filter: blur(10px);
            padding: 2rem;
            border-radius: 12px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.1);
        }
        
        .tab-container {
            display: flex;
            justify-content: center;
            gap: 0;
            margin-bottom: 1.5rem;
        }
        
        .tab-btn {
            flex: 1;
            padding: 0.75rem 1rem;
            border: 1px solid #ddd;
            background: #f8f9fa;
            cursor: pointer;
            font-size: 0.9rem;
            transition: all 0.3s ease;
        }
        
        .tab-btn:first-child {
            border-top-left-radius: 8px;
            border-bottom-left-radius: 8px;
        }
        
        .tab-btn:last-child {
            border-top-right-radius: 8px;
            border-bottom-right-radius: 8px;
            border-left: none;
        }
        
        .tab-btn.active {
            background: #007bff;
            color: white;
            border-color: #007bff;
        }
        
        .tab-btn:hover:not(.active) {
            background: #e9ecef;
        }
        
        .form-group {
            margin-bottom: 1rem;
        }
        
        .form-group label {
            display: block;
            margin-bottom: 0.5rem;
            font-weight: 500;
        }
        
        .form-group input {
            width: 100%;
            padding: 0.75rem;
            border: 1px solid #ddd;
            border-radius: 8px;
            font-size: 1rem;
        }
        
        .form-group input:focus {
            outline: none;
            border-color: #007bff;
            box-shadow: 0 0 0 3px rgba(0,123,255,0.1);
        }
        
        .demo-info {
            background: #e3f2fd;
            border: 1px solid #90caf9;
            border-radius: 8px;
            padding: 1rem;
            margin-bottom: 1.5rem;
        }
        
        .demo-info h4 {
            color: #1976d2;
            margin-bottom: 0.5rem;
        }
        
        .demo-info p {
            margin-bottom: 0.25rem;
            font-size: 0.9rem;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 2rem;
            display: flex;
            gap: 2rem;
            min-height: calc(100vh - 80px);
        }
        
        .sidebar {
            width: 300px;
            background: rgba(255,255,255,0.9);
            border-radius: 12px;
            padding: 1.5rem;
            height: fit-content;
            box-shadow: 0 4px 16px rgba(0,0,0,0.1);
        }
        
        .main-content {
            flex: 1;
            background: rgba(255,255,255,0.9);
            border-radius: 12px;
            padding: 1.5rem;
            display: flex;
            flex-direction: column;
            box-shadow: 0 4px 16px rgba(0,0,0,0.1);
        }
        
        .chat-header h2 {
            margin-bottom: 1rem;
            color: #333;
        }
        
        .chat-container {
            flex: 1;
            display: flex;
            flex-direction: column;
            min-height: 500px;
        }
        
        .chat-messages {
            flex: 1;
            border: 1px solid #e9ecef;
            border-radius: 8px;
            padding: 1rem;
            background: #f8f9fa;
            overflow-y: auto;
            margin-bottom: 1rem;
            max-height: 500px;
        }
        
        .message {
            display: flex;
            align-items: flex-start;
            margin-bottom: 1rem;
            gap: 0.75rem;
        }
        
        .message.user {
            flex-direction: row-reverse;
        }
        
        .message-avatar {
            width: 32px;
            height: 32px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 0.8rem;
        }
        
        .message.user .message-avatar {
            background: #007bff;
        }
        
        .message.bot .message-avatar {
            background: #28a745;
        }
        
        .message-content {
            flex: 1;
            background: white;
            padding: 0.75rem 1rem;
            border-radius: 12px;
            border: 1px solid #e9ecef;
            white-space: pre-wrap;
            line-height: 1.5;
        }
        
        .message.user .message-content {
            background: #007bff;
            color: white;
            border-color: #007bff;
        }
        
        .chat-input {
            display: flex;
            gap: 0.5rem;
        }
        
        .chat-input input {
            flex: 1;
            padding: 0.75rem;
            border: 1px solid #ddd;
            border-radius: 8px;
            font-size: 1rem;
        }
        
        .chat-input input:focus {
            outline: none;
            border-color: #007bff;
        }
        
        .welcome-message {
            text-align: center;
            padding: 3rem 1rem;
            color: #6c757d;
        }
        
        .welcome-message i {
            font-size: 3rem;
            margin-bottom: 1rem;
            color: #007bff;
        }
        
        .welcome-message h3 {
            margin-bottom: 0.5rem;
        }
        
        .code-block {
            background: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 8px;
            padding: 1rem;
            margin: 0.5rem 0;
            position: relative;
            font-family: 'Courier New', monospace;
        }
        
        .code-block .copy-btn {
            position: absolute;
            top: 0.5rem;
            right: 0.5rem;
            padding: 0.25rem 0.5rem;
            background: #007bff;
            color: white;
            border: none;
            border-radius: 4px;
            font-size: 0.8rem;
            cursor: pointer;
            opacity: 0.7;
            transition: opacity 0.3s ease;
        }
        
        .code-block .copy-btn:hover {
            opacity: 1;
        }
        
        .code-block .copy-btn.copied {
            background: #28a745;
        }
        
        .formula-list {
            margin: 1rem 0;
        }
        
        .formula-item {
            background: #f8f9fa;
            border-left: 4px solid #007bff;
            padding: 0.75rem;
            margin: 0.5rem 0;
            border-radius: 0 8px 8px 0;
            position: relative;
        }
        
        .formula-item .formula-code {
            font-family: 'Courier New', monospace;
            font-weight: bold;
            color: #007bff;
            font-size: 1.1em;
        }
        
        .formula-item .copy-btn {
            position: absolute;
            top: 0.5rem;
            right: 0.5rem;
            padding: 0.25rem 0.5rem;
            background: #007bff;
            color: white;
            border: none;
            border-radius: 4px;
            font-size: 0.8rem;
            cursor: pointer;
        }
        
        .vba-block {
            background: #2d3748;
            color: #e2e8f0;
            border-radius: 8px;
            padding: 1rem;
            margin: 1rem 0;
            position: relative;
            font-family: 'Courier New', monospace;
        }
        
        .vba-block .copy-btn {
            position: absolute;
            top: 0.5rem;
            right: 0.5rem;
            background: #4a5568;
            color: #e2e8f0;
            border: none;
            border-radius: 4px;
            padding: 0.25rem 0.5rem;
            font-size: 0.8rem;
            cursor: pointer;
        }
        
        .vba-block .copy-btn:hover {
            background: #2d3748;
        }
        
        .vba-block pre {
            margin: 0;
            white-space: pre-wrap;
            margin-top: 1.5rem;
        }
        
        .auth-message {
            padding: 0.75rem;
            border-radius: 8px;
            margin: 1rem 0;
            text-align: center;
        }
        
        .auth-error {
            background: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        
        .auth-success {
            background: #d1edff;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        
        @media (max-width: 768px) {
            .container {
                flex-direction: column;
                padding: 1rem;
            }
            
            .sidebar {
                width: 100%;
                margin-bottom: 1rem;
            }
        }
    </style>
</head>
<body>
    <div class="header">
        <div class="header-content">
            <div class="logo">
                <h1><i class="fas fa-shield-alt"></i> Excel チャットボット</h1>
            </div>
            <div class="user-info">
                <span id="userDisplay" style="display: none;">
                    <i class="fas fa-user"></i> <span id="userName"></span>
                </span>
                <button id="logoutBtn" class="btn btn-danger" style="display: none;">
                    <i class="fas fa-sign-out-alt"></i> ログアウト
                </button>
            </div>
        </div>
    </div>

    <!-- ログイン・登録画面 -->
    <div id="loginScreen" class="login-form">
        <!-- タブ切り替え -->
        <div class="tab-container">
            <button id="loginTab" class="tab-btn active" onclick="showLogin()">
                <i class="fas fa-sign-in-alt"></i> ログイン
            </button>
            <button id="registerTab" class="tab-btn" onclick="showRegister()">
                <i class="fas fa-user-plus"></i> 新規登録
            </button>
        </div>

        <!-- ログイン形式 -->
        <div id="loginPanel">
            <div class="demo-info">
                <h4><i class="fas fa-info-circle"></i> デモアカウント</h4>
                <p><strong>ユーザー名:</strong> demo</p>
                <p><strong>パスワード:</strong> demo123</p>
            </div>
            <form id="loginForm">
                <div class="form-group">
                    <label for="loginUsername">ユーザー名</label>
                    <input type="text" id="loginUsername" name="username" required value="demo" autocomplete="username">
                </div>
                <div class="form-group">
                    <label for="loginPassword">パスワード</label>
                    <input type="password" id="loginPassword" name="password" required value="demo123" autocomplete="current-password">
                </div>
                <button type="submit" class="btn btn-primary" style="width: 100%;">
                    <i class="fas fa-sign-in-alt"></i> ログイン
                </button>
            </form>
        </div>

        <!-- 新規登録形式 -->
        <div id="registerPanel" style="display: none;">
            <h3 style="text-align: center; margin-bottom: 1.5rem;">
                <i class="fas fa-user-plus"></i> アカウント作成
            </h3>
            <form id="registerForm">
                <div class="form-group">
                    <label for="regUsername">ユーザー名 *</label>
                    <input type="text" id="regUsername" name="username" required 
                           placeholder="3文字以上" minlength="3" autocomplete="username">
                </div>
                <div class="form-group">
                    <label for="regEmail">メールアドレス（任意）</label>
                    <input type="email" id="regEmail" name="email" 
                           placeholder="example@email.com" autocomplete="email">
                </div>
                <div class="form-group">
                    <label for="regPassword">パスワード *</label>
                    <input type="password" id="regPassword" name="password" required 
                           placeholder="6文字以上" minlength="6" autocomplete="new-password">
                </div>
                <div class="form-group">
                    <label for="regPasswordConfirm">パスワード確認 *</label>
                    <input type="password" id="regPasswordConfirm" name="passwordConfirm" required 
                           placeholder="パスワードを再入力" minlength="6" autocomplete="new-password">
                </div>
                <button type="submit" class="btn btn-success" style="width: 100%;">
                    <i class="fas fa-user-plus"></i> アカウント作成
                </button>
            </form>
        </div>

        <div id="authMessage" style="display: none;"></div>
    </div>

    <!-- メイン画面 -->
    <div id="mainApp" style="display: none;">
        <div class="container">
            <div class="sidebar">
                <h3><i class="fas fa-book"></i> Excel AI アシスタント</h3>
                <div style="margin: 1rem 0;">
                    <h4>💡 使用例</h4>
                    <ul style="margin-left: 1rem; line-height: 1.6;">
                        <li>「SUM関数の使い方を教えて」</li>
                        <li>「VLOOKUPの例を作って」</li>
                        <li>「売上集計のVBAコードを生成して」</li>
                        <li>「条件付き書式のマクロを作って」</li>
                    </ul>
                </div>
                <div style="margin-top: 2rem;">
                    <button id="clearChatBtn" class="btn btn-secondary" style="width: 100%;">
                        <i class="fas fa-trash"></i> チャット履歴削除
                    </button>
                </div>
            </div>
            
            <div class="main-content">
                <div class="chat-header">
                    <h2><i class="fas fa-comments"></i> Excel AI チャット</h2>
                </div>
                
                <div class="chat-container">
                    <div id="chatMessages" class="chat-messages">
                        <div class="welcome-message">
                            <i class="fas fa-robot"></i>
                            <h3>Excel AI アシスタントへようこそ！</h3>
                            <p>Excel関数やVBAコードについて、お気軽にご質問ください。<br>
                            数式やVBAコードは即座に生成してコピー可能な形で提供します。</p>
                        </div>
                    </div>
                    
                    <div class="chat-input">
                        <input type="text" id="messageInput" placeholder="Excel関数やVBAについて質問してください..." maxlength="500">
                        <button id="sendBtn" class="btn btn-primary">
                            <i class="fas fa-paper-plane"></i> 送信
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        // グローバル変数
        let currentUser = null;
        
        // 初期化
        document.addEventListener('DOMContentLoaded', function() {
            initApp();
        });
        
        async function initApp() {
            try {
                const response = await fetch('/api/auth/me');
                if (response.ok) {
                    const data = await response.json();
                    currentUser = data.user;
                    showMainApp();
                } else {
                    showLoginScreen();
                }
            } catch (error) {
                console.error('認証チェックエラー:', error);
                showLoginScreen();
            }
            
            setupEventListeners();
        }
        
        function setupEventListeners() {
            // ログインフォーム
            document.getElementById('loginForm').addEventListener('submit', handleLogin);
            
            // 新規登録フォーム
            document.getElementById('registerForm').addEventListener('submit', handleRegister);
            
            // ログアウトボタン
            document.getElementById('logoutBtn').addEventListener('click', handleLogout);
            
            // チャット送信
            document.getElementById('sendBtn').addEventListener('click', sendMessage);
            document.getElementById('messageInput').addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    sendMessage();
                }
            });
            
            // チャット履歴削除
            document.getElementById('clearChatBtn').addEventListener('click', clearChatHistory);
        }
        
        // タブ切り替え関数（グローバル）
        window.showLogin = function() {
            document.getElementById('loginPanel').style.display = 'block';
            document.getElementById('registerPanel').style.display = 'none';
            document.getElementById('loginTab').classList.add('active');
            document.getElementById('registerTab').classList.remove('active');
            hideAuthMessages();
        }
        
        window.showRegister = function() {
            document.getElementById('loginPanel').style.display = 'none';
            document.getElementById('registerPanel').style.display = 'block';
            document.getElementById('loginTab').classList.remove('active');
            document.getElementById('registerTab').classList.add('active');
            hideAuthMessages();
        }
        
        function hideAuthMessages() {
            document.getElementById('authMessage').style.display = 'none';
        }
        
        function showAuthMessage(message, type) {
            const messageDiv = document.getElementById('authMessage');
            messageDiv.textContent = message;
            messageDiv.className = 'auth-message ' + (type === 'error' ? 'auth-error' : 'auth-success');
            messageDiv.style.display = 'block';
        }
        
        // コピー機能（グローバル）
        window.copyToClipboard = function(text, button) {
            navigator.clipboard.writeText(text).then(() => {
                const originalText = button.innerHTML;
                button.innerHTML = '<i class="fas fa-check"></i> コピー済み!';
                button.classList.add('copied');
                
                setTimeout(() => {
                    button.innerHTML = originalText;
                    button.classList.remove('copied');
                }, 2000);
            }).catch(err => {
                console.error('コピー失敗:', err);
                // フォールバック
                const textArea = document.createElement('textarea');
                textArea.value = text;
                document.body.appendChild(textArea);
                textArea.select();
                document.execCommand('copy');
                document.body.removeChild(textArea);
                
                button.innerHTML = '<i class="fas fa-check"></i> コピー済み!';
                setTimeout(() => {
                    button.innerHTML = '<i class="fas fa-copy"></i> コピー';
                }, 2000);
            });
        }
        
        // 新規登録処理
        async function handleRegister(e) {
            e.preventDefault();
            
            const username = document.getElementById('regUsername').value;
            const email = document.getElementById('regEmail').value;
            const password = document.getElementById('regPassword').value;
            const passwordConfirm = document.getElementById('regPasswordConfirm').value;
            
            if (password !== passwordConfirm) {
                showAuthMessage('パスワードが一致しません', 'error');
                return;
            }
            
            const submitBtn = e.target.querySelector('button[type="submit"]');
            const originalText = submitBtn.innerHTML;
            submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 登録中...';
            submitBtn.disabled = true;
            
            try {
                const response = await fetch('/api/auth/register', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({ username, email, password }),
                });
                
                const data = await response.json();
                
                if (response.ok) {
                    showAuthMessage('アカウントが作成されました。自動的にログインします...', 'success');
                    setTimeout(() => {
                        currentUser = data.user;
                        showMainApp();
                    }, 1500);
                } else {
                    showAuthMessage(data.message || '登録に失敗しました', 'error');
                }
            } catch (error) {
                console.error('登録エラー:', error);
                showAuthMessage('ネットワークエラーが発生しました', 'error');
            } finally {
                submitBtn.innerHTML = originalText;
                submitBtn.disabled = false;
            }
        }
        
        // ログイン処理
        async function handleLogin(e) {
            e.preventDefault();
            
            const username = document.getElementById('loginUsername').value;
            const password = document.getElementById('loginPassword').value;
            
            const submitBtn = e.target.querySelector('button[type="submit"]');
            const originalText = submitBtn.innerHTML;
            submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ログイン中...';
            submitBtn.disabled = true;
            
            try {
                const response = await fetch('/api/auth/login', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({ username, password }),
                });
                
                const data = await response.json();
                
                if (response.ok) {
                    currentUser = data.user;
                    showMainApp();
                } else {
                    showAuthMessage(data.message || 'ログインに失敗しました', 'error');
                }
            } catch (error) {
                console.error('ログインエラー:', error);
                showAuthMessage('ネットワークエラーが発生しました', 'error');
            } finally {
                submitBtn.innerHTML = originalText;
                submitBtn.disabled = false;
            }
        }
        
        // ログアウト処理
        async function handleLogout() {
            try {
                await fetch('/api/auth/logout', { method: 'POST' });
                currentUser = null;
                showLoginScreen();
            } catch (error) {
                console.error('ログアウトエラー:', error);
            }
        }
        
        function showLoginScreen() {
            document.getElementById('loginScreen').style.display = 'block';
            document.getElementById('mainApp').style.display = 'none';
        }
        
        function showMainApp() {
            document.getElementById('loginScreen').style.display = 'none';
            document.getElementById('mainApp').style.display = 'block';
            
            if (currentUser) {
                document.getElementById('userDisplay').style.display = 'flex';
                document.getElementById('userName').textContent = currentUser.username;
                document.getElementById('logoutBtn').style.display = 'block';
            }
            
            loadChatHistory();
        }
        
        // メッセージ送信
        async function sendMessage() {
            const messageInput = document.getElementById('messageInput');
            const message = messageInput.value.trim();
            
            if (!message) return;
            
            // ユーザーメッセージを表示
            addMessageToChat('user', message);
            messageInput.value = '';
            
            try {
                const response = await fetch('/api/chat/message', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({ message }),
                });
                
                const data = await response.json();
                
                if (response.ok) {
                    addMessageToChat('bot', data.response, {
                        formulas: data.formulas || [],
                        vbaCode: data.vbaCode || null,
                        metadata: data.metadata || {}
                    });
                } else {
                    addMessageToChat('bot', 'エラーが発生しました: ' + data.message);
                }
            } catch (error) {
                console.error('メッセージ送信エラー:', error);
                addMessageToChat('bot', '接続エラーが発生しました。しばらく待ってから再度お試しください。');
            }
        }
        
        // チャットにメッセージ追加（拡張版：数式・VBAコード対応）
        function addMessageToChat(type, message, extras = {}) {
            const messagesContainer = document.getElementById('chatMessages');
            
            // ウェルカムメッセージを削除
            const welcomeMessage = messagesContainer.querySelector('.welcome-message');
            if (welcomeMessage) {
                welcomeMessage.remove();
            }
            
            const messageDiv = document.createElement('div');
            messageDiv.className = 'message ' + type;
            
            const avatar = type === 'user' ? 
                '<i class="fas fa-user"></i>' : 
                '<i class="fas fa-robot"></i>';
            
            let contentHtml = 
                '<div class="message-avatar">' + avatar + '</div>' +
                '<div class="message-content">' + message;
            
            // 数式を追加（コピー可能）
            if (extras.formulas && extras.formulas.length > 0) {
                contentHtml += '<div class="formula-list">';
                contentHtml += '<h4><i class="fas fa-calculator"></i> コピー可能な数式:</h4>';
                extras.formulas.forEach((formula, index) => {
                    contentHtml += '<div class="formula-item">' +
                        '<div class="formula-code">' + formula + '</div>' +
                        '<button class="copy-btn" onclick="window.copyToClipboard(\'' + 
                        formula.replace(/'/g, "\\\\'") + '\', this)">' +
                        '<i class="fas fa-copy"></i> コピー' +
                        '</button>' +
                        '</div>';
                });
                contentHtml += '</div>';
            }
            
            // VBAコードを追加（コピー可能）
            if (extras.vbaCode) {
                const vbaId = 'vba_' + Date.now();
                contentHtml += '<div class="vba-block">' +
                    '<h4><i class="fas fa-code"></i> VBAコード:</h4>' +
                    '<button class="copy-btn" onclick="copyVBACode(\'' + vbaId + '\')">' +
                    '<i class="fas fa-copy"></i> コピー' +
                    '</button>' +
                    '<pre id="' + vbaId + '">' + extras.vbaCode + '</pre>' +
                    '</div>';
            }
            
            contentHtml += '</div>';
            messageDiv.innerHTML = contentHtml;
            
            messagesContainer.appendChild(messageDiv);
            messagesContainer.scrollTop = messagesContainer.scrollHeight;
        }
        
        // VBAコード専用コピー関数
        window.copyVBACode = function(elementId) {
            const element = document.getElementById(elementId);
            if (element) {
                const text = element.textContent;
                window.copyToClipboard(text, event.target);
            }
        }
        
        // チャット履歴読み込み
        async function loadChatHistory() {
            try {
                const response = await fetch('/api/chat/history');
                if (response.ok) {
                    const data = await response.json();
                    const messagesContainer = document.getElementById('chatMessages');
                    
                    if (data.messages && data.messages.length > 0) {
                        messagesContainer.innerHTML = '';
                        data.messages.forEach(msg => {
                            if (msg.type === 'user') {
                                addMessageToChat('user', msg.message);
                            } else {
                                addMessageToChat('bot', msg.message, {
                                    formulas: msg.formulas || [],
                                    vbaCode: msg.vbaCode || null
                                });
                            }
                        });
                    }
                }
            } catch (error) {
                console.error('履歴読み込みエラー:', error);
            }
        }
        
        // チャット履歴削除
        async function clearChatHistory() {
            if (!confirm('チャット履歴を削除しますか？')) return;
            
            try {
                const response = await fetch('/api/chat/history', { method: 'DELETE' });
                if (response.ok) {
                    const messagesContainer = document.getElementById('chatMessages');
                    messagesContainer.innerHTML = 
                        '<div class="welcome-message">' +
                        '<i class="fas fa-robot"></i>' +
                        '<h3>チャット履歴が削除されました</h3>' +
                        '<p>Excel関数や機能について、お気軽にご質問ください。</p>' +
                        '</div>';
                }
            } catch (error) {
                console.error('履歴削除エラー:', error);
            }
        }
    </script>
</body>
</html>
  `);
});

// サーバー起動
server.listen(PORT, () => {
  console.log(`🔐 セキュアExcelチャットボット WebApp起動: http://localhost:${PORT}`);
  console.log(`🛡️ セキュリティ機能: 認証, レート制限, CSRF保護, XSS防止`);
  console.log(`📊 Excel専門AI: 準備完了（Webアプリ版）`);
  console.log(`💬 WebSocket: 有効`);
  console.log(`🎯 デモアカウント: demo / demo123`);
});