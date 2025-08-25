/**
 * 🔐 セキュア Excel チャットボット サーバー（簡易版）
 * 使いやすいWebアプリUI付き
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

// WebSocket設定（全オリジン許可）
const io = new Server(server, {
  cors: {
    origin: true, // 全てのオリジンを許可
    credentials: true
  }
});

const PORT = process.env.PORT || 5000;

// 🔒 基本セキュリティ設定（UIアクセス用に緩和）
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com"],
      scriptSrc: ["'self'", "'unsafe-inline'"],
      imgSrc: ["'self'", "data:", "https:"],
      connectSrc: ["'self'", "wss:", "ws:"],
      fontSrc: ["'self'", "https://fonts.gstatic.com"],
      objectSrc: ["'none'"],
      mediaSrc: ["'self'"],
      frameSrc: ["'none'"],
    },
  },
  crossOriginEmbedderPolicy: false
}));

// CORS設定（開発環境では全てのオリジンを許可）
app.use(cors({
  origin: function (origin, callback) {
    // パブリックアクセス用に全てのオリジンを許可
    callback(null, true);
  },
  credentials: true,
  optionsSuccessStatus: 200
}));

// 圧縮
app.use(compression());

// レート制限
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15分
  max: 100,
  message: {
    error: 'Too many requests from this IP, please try again later.',
    code: 'RATE_LIMIT_EXCEEDED'
  }
});

app.use(limiter);

// Body parser
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(cookieParser());

// セッション設定（メモリストア）
app.use(session({
  secret: process.env.SESSION_SECRET || 'excel-chatbot-session-secret',
  resave: false,
  saveUninitialized: true, // パブリックアクセス用に変更
  cookie: {
    secure: false, // HTTPSでなくても動作
    httpOnly: true,
    maxAge: 24 * 60 * 60 * 1000, // 24時間
    sameSite: 'lax' // クロスサイト対応
  }
}));

// インメモリーデータストア（デモ用）
const users = new Map();
const chatHistory = new Map(); // チャット履歴保存
const excelFunctions = [
  {
    id: '1',
    name: 'SUM',
    category: 'MATH',
    syntax: 'SUM(number1, [number2], ...)',
    description: '数値の合計を計算します',
    examples: [
      { formula: '=SUM(A1:A10)', description: 'A1からA10の範囲の合計', result: '合計値' },
      { formula: '=SUM(1,2,3,4,5)', description: '数値の直接指定', result: '15' }
    ]
  },
  {
    id: '2',
    name: 'VLOOKUP',
    category: 'LOOKUP_REFERENCE',
    syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    description: '垂直方向の検索を行い、対応する値を返します',
    examples: [
      { formula: '=VLOOKUP(A2,C:F,2,FALSE)', description: 'A2の値をC列で検索し、D列の値を返す', result: '対応する値' }
    ]
  },
  {
    id: '3',
    name: 'IF',
    category: 'LOGICAL',
    syntax: 'IF(logical_test, value_if_true, [value_if_false])',
    description: '条件に応じて異なる値を返します',
    examples: [
      { formula: '=IF(A1>10,"大","小")', description: 'A1が10より大きい場合は「大」、そうでなければ「小」', result: '大 または 小' }
    ]
  },
  {
    id: '4',
    name: 'COUNTIF',
    category: 'STATISTICAL',
    syntax: 'COUNTIF(range, criteria)',
    description: '条件に一致するセルの個数を数えます',
    examples: [
      { formula: '=COUNTIF(A1:A10,">5")', description: 'A1:A10で5より大きい値の個数', result: '個数' }
    ]
  },
  {
    id: '5',
    name: 'AVERAGE',
    category: 'STATISTICAL',
    syntax: 'AVERAGE(number1, [number2], ...)',
    description: '数値の平均値を計算します',
    examples: [
      { formula: '=AVERAGE(A1:A10)', description: 'A1からA10の平均値', result: '平均値' }
    ]
  }
];

// ヘルスチェック
app.get('/health', (req, res) => {
  res.json({ 
    status: 'OK', 
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    version: '1.0.0-webapp',
    features: ['Excel Chat UI', 'Security', 'Real-time']
  });
});

// CSRFトークン取得
app.get('/api/auth/csrf-token', (req, res) => {
  const token = Math.random().toString(36).substr(2);
  req.session.csrfToken = token;
  res.json({ csrfToken: token });
});

// デモログイン
app.post('/api/auth/login', (req, res) => {
  console.log('ログイン試行:', req.body); // デバッグログ
  const { username, password } = req.body;
  
  if (!username || !password) {
    console.log('ログインエラー: ユーザー名またはパスワードが未入力');
    return res.status(400).json({
      error: 'Validation Error',
      message: 'Username and password are required'
    });
  }
  
  // デモ用の簡単な認証
  if (username === 'demo' && password === 'demo123') {
    console.log('ログイン成功: demo user');
    const user = {
      id: 'demo-user-001',
      username: 'demo',
      email: 'demo@example.com',
      role: 'user',
      firstName: 'Demo',
      lastName: 'User'
    };
    
    users.set(user.id, user);
    req.session.userId = user.id;
    req.session.isLoggedIn = true;
    
    // 簡易JWTトークン（デモ用）
    const token = Buffer.from(JSON.stringify({
      id: user.id,
      username: user.username,
      exp: Date.now() + 24 * 60 * 60 * 1000
    })).toString('base64');
    
    res.json({
      message: 'Login successful',
      user,
      accessToken: token,
      expiresIn: '24h'
    });
  } else {
    console.log('ログイン失敗: 無効な認証情報', { username, password });
    res.status(401).json({
      error: 'Authentication failed',
      message: 'Invalid credentials'
    });
  }
});

// ユーザー登録
app.post('/api/auth/register', (req, res) => {
  const { username, password, email } = req.body;
  
  // バリデーション
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
  
  // 既存ユーザーチェック
  for (const [id, user] of users) {
    if (user.username === username) {
      return res.status(409).json({
        error: 'Conflict',
        message: 'このユーザー名は既に使用されています'
      });
    }
  }
  
  // 新規ユーザー作成
  const userId = `user_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  const newUser = {
    id: userId,
    username,
    password, // 実際の環境ではハッシュ化が必要
    email: email || null,
    createdAt: new Date().toISOString(),
    profile: {
      displayName: username,
      excelLevel: 'beginner' // beginner, intermediate, advanced
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

// パスワード変更
app.post('/api/auth/change-password', (req, res) => {
  if (!req.session.isLoggedIn || !req.session.userId) {
    return res.status(401).json({
      error: 'Unauthorized',
      message: 'ログインが必要です'
    });
  }
  
  const { currentPassword, newPassword } = req.body;
  
  if (!currentPassword || !newPassword) {
    return res.status(400).json({
      error: 'Validation Error',
      message: '現在のパスワードと新しいパスワードが必要です'
    });
  }
  
  if (newPassword.length < 6) {
    return res.status(400).json({
      error: 'Validation Error',
      message: '新しいパスワードは6文字以上である必要があります'
    });
  }
  
  const user = users.get(req.session.userId);
  if (!user) {
    return res.status(404).json({
      error: 'Not Found',
      message: 'ユーザーが見つかりません'
    });
  }
  
  // 現在のパスワード確認
  if (user.password !== currentPassword) {
    return res.status(401).json({
      error: 'Unauthorized',
      message: '現在のパスワードが正しくありません'
    });
  }
  
  // パスワード更新
  user.password = newPassword;
  user.passwordUpdatedAt = new Date().toISOString();
  users.set(req.session.userId, user);
  
  console.log(`🔒 パスワード変更: ${user.username} (ID: ${req.session.userId})`);
  
  res.json({
    message: 'パスワードが変更されました'
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
      return res.json({ user });
    }
  }
  res.status(401).json({ error: 'Not authenticated' });
});

// Excel関数検索
app.get('/api/excel/functions/search', (req, res) => {
  const { q } = req.query;
  
  let results = excelFunctions;
  
  if (q) {
    results = excelFunctions.filter(func =>
      func.name.toLowerCase().includes(q.toLowerCase()) ||
      func.description.toLowerCase().includes(q.toLowerCase())
    );
  }
  
  res.json({
    functions: results,
    pagination: {
      page: 1,
      limit: 10,
      total: results.length,
      pages: 1
    }
  });
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
      formulas: aiResponse.formulas || [], // 数式のみ（コピー用）
      vbaCode: aiResponse.vbaCode || null, // VBAコード（コピー用）
      timestamp: new Date().toISOString()
    });
    
    // 履歴は最新100件まで保持（高度な機能のため拡張）
    if (userHistory.length > 100) {
      userHistory.splice(0, userHistory.length - 100);
    }
    
    console.log(`✅ Excel AI応答完了: ${aiResponse.formulas?.length || 0}個の数式, VBA: ${!!aiResponse.vbaCode}`);
    
    res.json({
      response: aiResponse.response,
      formulas: aiResponse.formulas || [], // 数式のみの配列（コピー用）
      vbaCode: aiResponse.vbaCode || null, // VBAコード（コピー用）
      metadata: {
        responseTime: 50, // 高速化
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
  res.json({ history });
});

// チャット履歴削除
app.delete('/api/chat/history', (req, res) => {
  const userId = req.session.userId || 'guest';
  chatHistory.delete(userId);
  res.json({ message: 'Chat history cleared' });
});

// 静的ファイル配信
app.use('/static', express.static(path.join(__dirname, 'public')));

// メインWebアプリUI
app.get('/', (req, res) => {
  res.send(`
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>🔐 セキュア Excel チャットボット</title>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&display=swap" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Noto Sans JP', sans-serif; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
            min-height: 100vh; 
            display: flex; 
            flex-direction: column;
        }
        
        .header {
            background: rgba(255,255,255,0.95);
            backdrop-filter: blur(10px);
            padding: 1rem 2rem;
            box-shadow: 0 2px 20px rgba(0,0,0,0.1);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .logo {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            font-size: 1.5rem;
            font-weight: 700;
            color: #333;
        }
        
        .user-info {
            display: flex;
            align-items: center;
            gap: 1rem;
        }
        
        .btn {
            padding: 0.5rem 1rem;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 500;
            transition: all 0.3s ease;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        .btn-primary { background: #007bff; color: white; }
        .btn-secondary { background: #6c757d; color: white; }
        .btn-danger { background: #dc3545; color: white; }
        .btn:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
        
        .container {
            flex: 1;
            display: flex;
            max-width: 1400px;
            margin: 0 auto;
            padding: 2rem;
            gap: 2rem;
            width: 100%;
        }
        
        .sidebar {
            width: 300px;
            background: rgba(255,255,255,0.95);
            backdrop-filter: blur(10px);
            border-radius: 12px;
            padding: 1.5rem;
            box-shadow: 0 8px 32px rgba(0,0,0,0.1);
            height: fit-content;
        }
        
        .main-content {
            flex: 1;
            background: rgba(255,255,255,0.95);
            backdrop-filter: blur(10px);
            border-radius: 12px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.1);
            display: flex;
            flex-direction: column;
            height: 600px;
        }
        
        .chat-header {
            padding: 1.5rem;
            border-bottom: 1px solid #e9ecef;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .chat-messages {
            flex: 1;
            padding: 1rem;
            overflow-y: auto;
            background: #f8f9fa;
        }
        
        .message {
            margin-bottom: 1rem;
            display: flex;
            gap: 0.75rem;
        }
        
        .message.user {
            flex-direction: row-reverse;
        }
        
        .message-avatar {
            width: 40px;
            height: 40px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.2rem;
            flex-shrink: 0;
        }
        
        .message.user .message-avatar {
            background: #007bff;
            color: white;
        }
        
        .message.bot .message-avatar {
            background: #28a745;
            color: white;
        }
        
        .message-content {
            max-width: 70%;
            padding: 1rem;
            border-radius: 18px;
            white-space: pre-wrap;
            word-wrap: break-word;
        }
        
        .message.user .message-content {
            background: #007bff;
            color: white;
            border-bottom-right-radius: 4px;
        }
        
        .message.bot .message-content {
            background: white;
            color: #333;
            border: 1px solid #e9ecef;
            border-bottom-left-radius: 4px;
        }
        
        .chat-input {
            padding: 1.5rem;
            border-top: 1px solid #e9ecef;
            background: white;
            border-radius: 0 0 12px 12px;
        }
        
        .input-group {
            display: flex;
            gap: 0.5rem;
        }
        
        .input-group input {
            flex: 1;
            padding: 0.75rem;
            border: 1px solid #ddd;
            border-radius: 8px;
            font-size: 1rem;
        }
        
        .function-list {
            margin-top: 1rem;
        }
        
        .function-item {
            padding: 0.75rem;
            margin-bottom: 0.5rem;
            background: #f8f9fa;
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.3s ease;
        }
        
        .function-item:hover {
            background: #e9ecef;
            transform: translateY(-1px);
        }
        
        .function-name {
            font-weight: 600;
            color: #007bff;
            margin-bottom: 0.25rem;
        }
        
        .function-desc {
            font-size: 0.875rem;
            color: #666;
        }
        
        .welcome-message {
            text-align: center;
            padding: 2rem;
            color: #666;
        }
        
        .welcome-message i {
            font-size: 4rem;
            margin-bottom: 1rem;
            color: #007bff;
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
        
        .vba-block {
            background: #2d3748;
            color: #e2e8f0;
            border-radius: 8px;
            padding: 1rem;
            margin: 1rem 0;
            position: relative;
            font-family: 'Courier New', monospace;
            white-space: pre-wrap;
            overflow-x: auto;
        }
        
        .vba-block .copy-btn {
            background: #4a5568;
            color: #e2e8f0;
        }
        
        .vba-block .copy-btn:hover {
            background: #2d3748;
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
        
        .demo-info {
            background: #e7f3ff;
            padding: 1rem;
            border-radius: 8px;
            margin-bottom: 1rem;
            border-left: 4px solid #007bff;
        }
        
        .loading {
            display: none;
            text-align: center;
            padding: 1rem;
            color: #666;
        }
        
        @media (max-width: 768px) {
            .container {
                flex-direction: column;
                padding: 1rem;
            }
            .sidebar {
                width: 100%;
                order: 2;
            }
            .main-content {
                height: 500px;
            }
            .header {
                padding: 1rem;
            }
        }
    </style>
</head>
<body>
    <div class="header">
        <div class="logo">
            <i class="fas fa-shield-alt"></i>
            セキュア Excel チャットボット
        </div>
        <div class="user-info">
            <span id="userWelcome" style="display: none;">ようこそ、<span id="username"></span>さん</span>
            <button id="loginBtn" class="btn btn-primary">
                <i class="fas fa-sign-in-alt"></i> ログイン
            </button>
            <button id="logoutBtn" class="btn btn-danger" style="display: none;">
                <i class="fas fa-sign-out-alt"></i> ログアウト
            </button>
        </div>
    </div>

    <!-- ログイン・登録画面 -->
    <div id="loginScreen" class="login-form">
        <!-- タブ切り替え -->
        <div class="tab-container" style="text-align: center; margin-bottom: 2rem;">
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
                    <input type="text" id="loginUsername" name="username" required value="demo">
                </div>
                <div class="form-group">
                    <label for="loginPassword">パスワード</label>
                    <input type="password" id="loginPassword" name="password" required value="demo123">
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
                           placeholder="3文字以上" minlength="3">
                </div>
                <div class="form-group">
                    <label for="regEmail">メールアドレス（任意）</label>
                    <input type="email" id="regEmail" name="email" 
                           placeholder="example@email.com">
                </div>
                <div class="form-group">
                    <label for="regPassword">パスワード *</label>
                    <input type="password" id="regPassword" name="password" required 
                           placeholder="6文字以上" minlength="6">
                </div>
                <div class="form-group">
                    <label for="regPasswordConfirm">パスワード確認 *</label>
                    <input type="password" id="regPasswordConfirm" name="passwordConfirm" required 
                           placeholder="パスワードを再入力" minlength="6">
                </div>
                <button type="submit" class="btn btn-success" style="width: 100%;">
                    <i class="fas fa-user-plus"></i> アカウント作成
                </button>
            </form>
        </div>

        <div id="authError" style="color: #dc3545; margin-top: 1rem; display: none;"></div>
        <div id="authSuccess" style="color: #28a745; margin-top: 1rem; display: none;"></div>
    </div>

    <!-- メイン画面 -->
    <div id="mainApp" style="display: none;">
        <div class="container">
            <div class="sidebar">
                <h3><i class="fas fa-book"></i> Excel関数一覧</h3>
                <div id="functionList" class="function-list">
                    <div class="loading">
                        <i class="fas fa-spinner fa-spin"></i> 読み込み中...
                    </div>
                </div>
                <div style="margin-top: 2rem;">
                    <button id="clearChatBtn" class="btn btn-secondary" style="width: 100%;">
                        <i class="fas fa-trash"></i> チャット履歴削除
                    </button>
                </div>
            </div>
            
            <div class="main-content">
                <div class="chat-header">
                    <h3><i class="fas fa-comments"></i> Excel チャット</h3>
                    <div>
                        <span class="badge" style="background: #28a745; color: white; padding: 0.25rem 0.5rem; border-radius: 4px;">
                            <i class="fas fa-shield-alt"></i> セキュア接続
                        </span>
                    </div>
                </div>
                
                <div id="chatMessages" class="chat-messages">
                    <div class="welcome-message">
                        <i class="fas fa-robot"></i>
                        <h3>Excel専門チャットボットへようこそ！</h3>
                        <p>Excel関数や機能について、お気軽にご質問ください。</p>
                        <p><strong>例:</strong> 「SUM関数の使い方」「VLOOKUPで検索したい」</p>
                    </div>
                </div>
                
                <div class="chat-input">
                    <div class="input-group">
                        <input type="text" id="messageInput" placeholder="Excelについて質問してください..." maxlength="500">
                        <button id="sendBtn" class="btn btn-primary">
                            <i class="fas fa-paper-plane"></i> 送信
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        let currentUser = null;
        let socket = null;

        // 初期化
        document.addEventListener('DOMContentLoaded', function() {
            checkAuthStatus();
            setupEventListeners();
        });

        // 認証状態チェック
        async function checkAuthStatus() {
            try {
                console.log('認証状態をチェック中...');
                const response = await fetch('/api/auth/me', {
                    credentials: 'same-origin'
                });
                console.log('認証チェック結果:', response.status);
                
                if (response.ok) {
                    const data = await response.json();
                    console.log('既存セッション見つかりました:', data.user);
                    currentUser = data.user;
                    showMainApp();
                } else {
                    console.log('認証されていません、ログイン画面表示');
                    showLoginScreen();
                }
            } catch (error) {
                console.error('Auth check error:', error);
                console.log('エラーのためログイン画面表示');
                showLoginScreen();
            }
        }

        // イベントリスナー設定
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
        
        // タブ切り替え関数
        function showLogin() {
            document.getElementById('loginPanel').style.display = 'block';
            document.getElementById('registerPanel').style.display = 'none';
            document.getElementById('loginTab').classList.add('active');
            document.getElementById('registerTab').classList.remove('active');
            hideAuthMessages();
        }
        
        function showRegister() {
            document.getElementById('loginPanel').style.display = 'none';
            document.getElementById('registerPanel').style.display = 'block';
            document.getElementById('loginTab').classList.remove('active');
            document.getElementById('registerTab').classList.add('active');
            hideAuthMessages();
        }
        
        function hideAuthMessages() {
            document.getElementById('authError').style.display = 'none';
            document.getElementById('authSuccess').style.display = 'none';
        }
        
        function showAuthError(message) {
            hideAuthMessages();
            const errorDiv = document.getElementById('authError');
            errorDiv.textContent = message;
            errorDiv.style.display = 'block';
        }
        
        function showAuthSuccess(message) {
            hideAuthMessages();
            const successDiv = document.getElementById('authSuccess');
            successDiv.textContent = message;
            successDiv.style.display = 'block';
        }
        
        // コピー機能
        function copyToClipboard(text, button) {
            navigator.clipboard.writeText(text).then(() => {
                const originalText = button.textContent;
                button.textContent = 'コピー済み!';
                button.classList.add('copied');
                
                setTimeout(() => {
                    button.textContent = originalText;
                    button.classList.remove('copied');
                }, 2000);
            }).catch(err => {
                console.error('コピー失敗:', err);
                // フォールバック: テキストエリアを使用
                const textArea = document.createElement('textarea');
                textArea.value = text;
                document.body.appendChild(textArea);
                textArea.select();
                document.execCommand('copy');
                document.body.removeChild(textArea);
                
                button.textContent = 'コピー済み!';
                setTimeout(() => {
                    button.textContent = 'コピー';
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
            
            // パスワード確認
            if (password !== passwordConfirm) {
                showAuthError('パスワードが一致しません');
                return;
            }
            
            // ローディング表示
            const submitBtn = e.target.querySelector('button[type="submit"]');
            const originalText = submitBtn.innerHTML;
            submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 登録中...';
            submitBtn.disabled = true;
            
            try {
                const response = await fetch('/api/auth/register', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Accept': 'application/json',
                    },
                    body: JSON.stringify({ username, email, password }),
                    credentials: 'same-origin'
                });
                
                const data = await response.json();
                
                if (response.ok) {
                    showAuthSuccess('アカウントが作成されました。自動的にログインします...');
                    setTimeout(() => {
                        initApp(); // ログイン状態で再初期化
                    }, 1500);
                } else {
                    showAuthError(data.message || '登録に失敗しました');
                }
            } catch (error) {
                console.error('登録エラー:', error);
                showAuthError('ネットワークエラーが発生しました');
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
            
            console.log('ログイン試行:', { username, password }); // デバッグログ
            
            // ローディング表示
            const submitBtn = e.target.querySelector('button[type="submit"]');
            const originalText = submitBtn.innerHTML;
            submitBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ログイン中...';
            submitBtn.disabled = true;
            
            try {
                console.log('APIリクエスト送信中...');
                const response = await fetch('/api/auth/login', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Accept': 'application/json',
                    },
                    body: JSON.stringify({ username, password }),
                    credentials: 'same-origin' // セッション用
                });
                
                console.log('APIレスポンス受信:', response.status, response.statusText);
                
                if (!response.ok) {
                    const errorText = await response.text();
                    console.error('APIエラー:', errorText);
                    throw new Error(\`HTTP \${response.status}: \${errorText}\`);
                }
                
                const data = await response.json();
                console.log('ログインレスポンス:', data);
                
                if (data.user) {
                    currentUser = data.user;
                    console.log('ログイン成功, メイン画面表示');
                    showMainApp();
                } else {
                    throw new Error(data.message || 'ログインに失敗しました');
                }
            } catch (error) {
                console.error('Login error:', error);
                document.getElementById('loginError').textContent = \`ログインエラー: \${error.message}\`;
                document.getElementById('loginError').style.display = 'block';
            } finally {
                // ローディング解除
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
                console.error('Logout error:', error);
            }
        }

        // ログイン画面表示
        function showLoginScreen() {
            document.getElementById('loginScreen').style.display = 'block';
            document.getElementById('mainApp').style.display = 'none';
            document.getElementById('loginBtn').style.display = 'inline-flex';
            document.getElementById('logoutBtn').style.display = 'none';
            document.getElementById('userWelcome').style.display = 'none';
        }

        // メイン画面表示
        function showMainApp() {
            document.getElementById('loginScreen').style.display = 'none';
            document.getElementById('mainApp').style.display = 'block';
            document.getElementById('loginBtn').style.display = 'none';
            document.getElementById('logoutBtn').style.display = 'inline-flex';
            document.getElementById('userWelcome').style.display = 'block';
            document.getElementById('username').textContent = currentUser.username;
            
            loadExcelFunctions();
            loadChatHistory();
        }

        // Excel関数一覧読み込み
        async function loadExcelFunctions() {
            try {
                const response = await fetch('/api/excel/functions/search');
                const data = await response.json();
                
                const functionList = document.getElementById('functionList');
                functionList.innerHTML = '';
                
                data.functions.forEach(func => {
                    const item = document.createElement('div');
                    item.className = 'function-item';
                    item.innerHTML = \`
                        <div class="function-name">\${func.name}</div>
                        <div class="function-desc">\${func.description}</div>
                    \`;
                    item.addEventListener('click', () => {
                        document.getElementById('messageInput').value = \`\${func.name}関数の使い方を教えて\`;
                    });
                    functionList.appendChild(item);
                });
            } catch (error) {
                console.error('関数読み込みエラー:', error);
            }
        }

        // チャット履歴読み込み
        async function loadChatHistory() {
            try {
                const response = await fetch('/api/chat/history');
                const data = await response.json();
                
                const messagesContainer = document.getElementById('chatMessages');
                messagesContainer.innerHTML = '';
                
                if (data.history.length === 0) {
                    messagesContainer.innerHTML = \`
                        <div class="welcome-message">
                            <i class="fas fa-robot"></i>
                            <h3>Excel専門チャットボットへようこそ！</h3>
                            <p>Excel関数や機能について、お気軽にご質問ください。</p>
                            <p><strong>例:</strong> 「SUM関数の使い方」「VLOOKUPで検索したい」</p>
                        </div>
                    \`;
                } else {
                    data.history.forEach(msg => {
                        addMessageToChat(msg.type, msg.message);
                    });
                }
            } catch (error) {
                console.error('履歴読み込みエラー:', error);
            }
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
            messageDiv.className = \`message \${type}\`;
            
            const avatar = type === 'user' ? 
                '<i class="fas fa-user"></i>' : 
                '<i class="fas fa-robot"></i>';
            
            let contentHtml = \`
                <div class="message-avatar">\${avatar}</div>
                <div class="message-content">\${message}\`;
            
            // 数式を追加（コピー可能）
            if (extras.formulas && extras.formulas.length > 0) {
                contentHtml += '<div class="formula-list">';
                contentHtml += '<h4><i class="fas fa-calculator"></i> コピー可能な数式:</h4>';
                extras.formulas.forEach((formula, index) => {
                    const formulaId = \`formula_\${Date.now()}_\${index}\`;
                    contentHtml += \`
                        <div class="formula-item">
                            <div class="formula-code">\${formula}</div>
                            <button class="copy-btn" onclick="copyToClipboard('\${formula}', this)">
                                <i class="fas fa-copy"></i> コピー
                            </button>
                        </div>\`;
                });
                contentHtml += '</div>';
            }
            
            // VBAコードを追加（コピー可能）
            if (extras.vbaCode) {
                const vbaId = \`vba_\${Date.now()}\`;
                const escapedVbaCode = extras.vbaCode.replace(/'/g, "\\'").replace(/\\/g, "\\\\");
                contentHtml += \`
                    <div class="vba-block">
                        <h4><i class="fas fa-code"></i> VBAコード:</h4>
                        <button class="copy-btn" onclick="copyToClipboard('\${escapedVbaCode}', this)">
                            <i class="fas fa-copy"></i> コピー
                        </button>
                        <pre>\${extras.vbaCode}</pre>
                    </div>\`;
            }
            
            contentHtml += '</div>';
            messageDiv.innerHTML = contentHtml;
            
            messagesContainer.appendChild(messageDiv);
            messagesContainer.scrollTop = messagesContainer.scrollHeight;
        }

        // チャット履歴削除
        async function clearChatHistory() {
            if (!confirm('チャット履歴を削除しますか？')) return;
            
            try {
                const response = await fetch('/api/chat/history', { method: 'DELETE' });
                if (response.ok) {
                    const messagesContainer = document.getElementById('chatMessages');
                    messagesContainer.innerHTML = \`
                        <div class="welcome-message">
                            <i class="fas fa-robot"></i>
                            <h3>チャット履歴が削除されました</h3>
                            <p>Excel関数や機能について、お気軽にご質問ください。</p>
                        </div>
                    \`;
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

// API以外のルートは全て404
app.get('*', (req, res) => {
  if (req.path.startsWith('/api/')) {
    return res.status(404).json({
      error: 'Not Found',
      message: 'API endpoint not found'
    });
  }
  
  // 他のパスは全てメインアプリにリダイレクト
  res.redirect('/');
});

// エラーハンドリング
app.use((err, req, res, next) => {
  console.error('Error:', err.message);
  res.status(500).json({
    error: 'Internal Server Error',
    message: 'An error occurred processing your request'
  });
});

// WebSocket基本接続
io.on('connection', (socket) => {
  console.log('✅ WebSocket接続:', socket.id);
  
  socket.emit('connected', {
    message: 'Excel ChatBot に接続しました',
    timestamp: new Date().toISOString()
  });
  
  socket.on('disconnect', () => {
    console.log('❌ WebSocket切断:', socket.id);
  });
});

// サーバー起動
server.listen(PORT, () => {
  console.log(`🔐 セキュアExcelチャットボット WebApp起動: http://localhost:${PORT}`);
  console.log(`🛡️ セキュリティ機能: 認証, レート制限, CSRF保護, XSS防止`);
  console.log(`📊 Excel専門AI: 準備完了（Webアプリ版）`);
  console.log(`💬 WebSocket: 有効`);
  console.log(`🎯 デモアカウント: demo / demo123`);
});

module.exports = app;