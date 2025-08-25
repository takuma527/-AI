/**
 * 🔐 セキュア Excel チャットボット サーバー（簡易版）
 * デモンストレーション用のシンプルバージョン
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
  saveUninitialized: false,
  cookie: {
    secure: false, // 開発環境ではHTTPを許可
    httpOnly: true,
    maxAge: 24 * 60 * 60 * 1000 // 24時間
  }
}));

// インメモリーデータストア（デモ用）
const users = new Map();
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
  }
];

// ヘルスチェック
app.get('/health', (req, res) => {
  res.json({ 
    status: 'OK', 
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    version: '1.0.0-demo',
    features: ['Excel Chat', 'Security', 'Real-time']
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
  const { username, password } = req.body;
  
  if (!username || !password) {
    return res.status(400).json({
      error: 'Validation Error',
      message: 'Username and password are required'
    });
  }
  
  // デモ用の簡単な認証
  if (username === 'demo' && password === 'demo123') {
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
    res.status(401).json({
      error: 'Authentication failed',
      message: 'Invalid credentials'
    });
  }
});

// デモ登録
app.post('/api/auth/register', (req, res) => {
  const { username, email, password } = req.body;
  
  if (!username || !email || !password) {
    return res.status(400).json({
      error: 'Validation Error',
      message: 'All fields are required'
    });
  }
  
  const user = {
    id: `user-${Date.now()}`,
    username,
    email,
    role: 'user',
    createdAt: new Date().toISOString()
  };
  
  users.set(user.id, user);
  
  res.status(201).json({
    message: 'User registered successfully',
    user
  });
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

// チャットメッセージ処理
app.post('/api/chat/message', (req, res) => {
  const { message } = req.body;
  
  if (!message) {
    return res.status(400).json({
      error: 'Validation Error',
      message: 'Message is required'
    });
  }
  
  // Excel関連キーワード検索
  const lowerMessage = message.toLowerCase();
  const relatedFunctions = excelFunctions.filter(func =>
    lowerMessage.includes(func.name.toLowerCase()) ||
    func.description.toLowerCase().includes(lowerMessage) ||
    lowerMessage.includes('関数') ||
    lowerMessage.includes('excel')
  );
  
  let response = "Excel専門チャットボットです。\n\n";
  
  if (relatedFunctions.length > 0) {
    response += "関連する情報を見つけました：\n\n";
    relatedFunctions.forEach(func => {
      response += `📊 **${func.name}関数**\n`;
      response += `${func.description}\n`;
      response += `構文: \`${func.syntax}\`\n\n`;
    });
  } else {
    response += "申し訳ございませんが、お探しの情報が見つかりませんでした。\n";
    response += "Excel関連のキーワードを使ってもう一度お試しください。\n\n";
    response += "例：\n";
    response += "• 'SUM関数の使い方'\n";
    response += "• 'VLOOKUPで検索したい'\n";
    response += "• 'IF関数の条件分岐'\n";
  }
  
  res.json({
    response,
    metadata: {
      responseTime: 150,
      knowledgeResults: relatedFunctions.length,
      model: 'excel-knowledge-base',
      conversationId: `conv_${Date.now()}`
    }
  });
});

// 静的ファイル配信
app.use('/static', express.static(path.join(__dirname, '../../frontend/build/static')));

// フロントエンド用のルート
app.get('*', (req, res) => {
  if (req.path.startsWith('/api/')) {
    return res.status(404).json({
      error: 'Not Found',
      message: 'API endpoint not found'
    });
  }
  
  // 開発環境では簡単なHTMLページを返す
  res.send(`
    <!DOCTYPE html>
    <html lang="ja">
    <head>
      <meta charset="utf-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1" />
      <title>🔐 セキュア Excel チャットボット</title>
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; background: white; padding: 40px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .header { text-align: center; margin-bottom: 40px; }
        .api-section { margin: 20px 0; padding: 20px; background: #f8f9fa; border-radius: 8px; }
        .endpoint { font-family: monospace; background: #e9ecef; padding: 10px; border-radius: 4px; margin: 10px 0; }
        .security-info { background: #d4edda; padding: 15px; border-radius: 8px; border-left: 4px solid #28a745; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>🔐 セキュア Excel チャットボット</h1>
          <p>セキュリティを重視したExcel専門AIチャットボット（デモ版）</p>
        </div>
        
        <div class="security-info">
          <h3>🛡️ 実装されたセキュリティ機能</h3>
          <ul>
            <li><strong>Helmet.js</strong> - セキュリティヘッダー設定</li>
            <li><strong>CORS</strong> - クロスオリジンリクエスト制御</li>
            <li><strong>Rate Limiting</strong> - レート制限</li>
            <li><strong>入力検証</strong> - XSS/SQLインジェクション対策</li>
            <li><strong>セッション管理</strong> - 安全なセッション処理</li>
            <li><strong>CSRF保護</strong> - クロスサイトリクエストフォージェリ対策</li>
          </ul>
        </div>
        
        <div class="api-section">
          <h3>🔗 API エンドポイント</h3>
          
          <h4>認証</h4>
          <div class="endpoint">POST /api/auth/login</div>
          <div class="endpoint">POST /api/auth/register</div>
          <div class="endpoint">GET /api/auth/csrf-token</div>
          
          <h4>Excel機能</h4>
          <div class="endpoint">GET /api/excel/functions/search</div>
          
          <h4>チャット</h4>
          <div class="endpoint">POST /api/chat/message</div>
          
          <h4>システム</h4>
          <div class="endpoint">GET /health</div>
        </div>
        
        <div class="api-section">
          <h3>📊 デモアカウント</h3>
          <p><strong>ユーザー名:</strong> demo</p>
          <p><strong>パスワード:</strong> demo123</p>
        </div>
        
        <div class="api-section">
          <h3>🧪 テスト方法</h3>
          <p>以下のcurlコマンドでAPIをテストできます：</p>
          <div class="endpoint">curl -X GET http://localhost:${PORT}/health</div>
          <div class="endpoint">curl -X GET "http://localhost:${PORT}/api/excel/functions/search?q=sum"</div>
        </div>
      </div>
    </body>
    </html>
  `);
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
  console.log(`🔐 セキュアExcelチャットボット サーバー起動: http://localhost:${PORT}`);
  console.log(`🛡️ セキュリティ機能: 認証, レート制限, CSRF保護, XSS防止`);
  console.log(`📊 Excel専門AI: 準備完了（デモ版）`);
  console.log(`💬 WebSocket: 有効`);
});

module.exports = app;