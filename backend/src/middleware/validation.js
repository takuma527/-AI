/**
 * 🔐 入力検証ミドルウェア
 * 厳格な入力サニタイゼーションとバリデーション
 */

const DOMPurify = require('dompurify');
const { JSDOM } = require('jsdom');
const { body, validationResult } = require('express-validator');
const logger = require('../utils/logger');

// DOMPurifyの初期化
const window = new JSDOM('').window;
const purify = DOMPurify(window);

/**
 * 入力サニタイゼーション設定
 */
const sanitizeConfig = {
  ALLOWED_TAGS: [], // HTMLタグを全て除去
  ALLOWED_ATTR: [], // 属性を全て除去
  KEEP_CONTENT: true, // テキストコンテンツは保持
  ALLOW_DATA_ATTR: false
};

/**
 * 汎用入力サニタイゼーション
 */
const sanitizeInput = (input) => {
  if (typeof input === 'string') {
    // HTMLタグの除去
    let sanitized = purify.sanitize(input, sanitizeConfig);
    
    // 危険な文字のエスケープ
    sanitized = sanitized
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#x27;')
      .replace(/\//g, '&#x2F;');
    
    // 制御文字の除去
    sanitized = sanitized.replace(/[\x00-\x1F\x7F]/g, '');
    
    return sanitized.trim();
  }
  
  if (Array.isArray(input)) {
    return input.map(sanitizeInput);
  }
  
  if (typeof input === 'object' && input !== null) {
    const sanitized = {};
    for (const [key, value] of Object.entries(input)) {
      sanitized[sanitizeInput(key)] = sanitizeInput(value);
    }
    return sanitized;
  }
  
  return input;
};

/**
 * 入力サニタイゼーションミドルウェア
 */
const inputValidation = (req, res, next) => {
  try {
    // ボディの サニタイゼーション
    if (req.body && typeof req.body === 'object') {
      req.body = sanitizeInput(req.body);
    }
    
    // クエリパラメータのサニタイゼーション
    if (req.query && typeof req.query === 'object') {
      req.query = sanitizeInput(req.query);
    }
    
    // パラメータのサニタイゼーション
    if (req.params && typeof req.params === 'object') {
      req.params = sanitizeInput(req.params);
    }
    
    next();
  } catch (error) {
    logger.error('Input sanitization error:', error);
    res.status(400).json({
      error: 'Bad Request',
      message: 'Invalid input format'
    });
  }
};

/**
 * チャットメッセージのバリデーション
 */
const chatMessageValidation = [
  body('message')
    .isLength({ min: 1, max: 5000 })
    .withMessage('Message must be between 1 and 5000 characters')
    .custom((value) => {
      // Excelに関連するコンテンツかチェック
      const excelKeywords = [
        'excel', 'spreadsheet', 'formula', 'function', 'cell', 'row', 'column',
        'pivot', 'chart', 'macro', 'vba', 'worksheet', 'workbook', 'sum',
        'vlookup', 'hlookup', 'if', 'countif', 'sumif', 'index', 'match'
      ];
      
      const lowerMessage = value.toLowerCase();
      const hasExcelContent = excelKeywords.some(keyword => 
        lowerMessage.includes(keyword)
      );
      
      if (!hasExcelContent) {
        logger.warn(`Non-Excel related message detected: ${value}`);
        // 警告するが、完全にブロックはしない
      }
      
      return true;
    }),
  
  body('conversationId')
    .optional()
    .isUUID()
    .withMessage('Invalid conversation ID format'),
  
  body('attachments')
    .optional()
    .isArray({ max: 5 })
    .withMessage('Too many attachments (max 5)')
];

/**
 * ユーザー登録のバリデーション
 */
const userRegistrationValidation = [
  body('username')
    .isLength({ min: 3, max: 30 })
    .withMessage('Username must be between 3 and 30 characters')
    .matches(/^[a-zA-Z0-9_-]+$/)
    .withMessage('Username can only contain letters, numbers, underscores, and hyphens'),
  
  body('email')
    .isEmail()
    .withMessage('Invalid email format')
    .normalizeEmail(),
  
  body('password')
    .isLength({ min: 8, max: 128 })
    .withMessage('Password must be between 8 and 128 characters')
    .matches(/^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[@$!%*?&])[A-Za-z\d@$!%*?&]/)
    .withMessage('Password must contain at least one uppercase letter, lowercase letter, number, and special character'),
  
  body('firstName')
    .optional()
    .isLength({ min: 1, max: 50 })
    .withMessage('First name must be between 1 and 50 characters')
    .matches(/^[a-zA-Z\s]+$/)
    .withMessage('First name can only contain letters and spaces'),
  
  body('lastName')
    .optional()
    .isLength({ min: 1, max: 50 })
    .withMessage('Last name must be between 1 and 50 characters')
    .matches(/^[a-zA-Z\s]+$/)
    .withMessage('Last name can only contain letters and spaces')
];

/**
 * ログインのバリデーション
 */
const loginValidation = [
  body('username')
    .isLength({ min: 1, max: 100 })
    .withMessage('Username is required'),
  
  body('password')
    .isLength({ min: 1, max: 128 })
    .withMessage('Password is required')
];

/**
 * ファイルアップロードのバリデーション
 */
const fileUploadValidation = (req, res, next) => {
  const allowedMimeTypes = [
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'text/csv',
    'image/jpeg',
    'image/png',
    'image/gif'
  ];
  
  const maxFileSize = 10 * 1024 * 1024; // 10MB
  
  if (req.files) {
    for (const file of req.files) {
      if (!allowedMimeTypes.includes(file.mimetype)) {
        return res.status(400).json({
          error: 'Invalid file type',
          message: 'Only Excel, CSV, and image files are allowed'
        });
      }
      
      if (file.size > maxFileSize) {
        return res.status(400).json({
          error: 'File too large',
          message: 'File size must be less than 10MB'
        });
      }
    }
  }
  
  next();
};

/**
 * バリデーションエラーハンドラー
 */
const handleValidationErrors = (req, res, next) => {
  const errors = validationResult(req);
  
  if (!errors.isEmpty()) {
    logger.warn(`Validation errors from IP ${req.ip}:`, errors.array());
    
    return res.status(400).json({
      error: 'Validation Error',
      message: 'Invalid input data',
      details: errors.array()
    });
  }
  
  next();
};

/**
 * SQL インジェクション検出
 */
const sqlInjectionDetection = (req, res, next) => {
  const sqlPatterns = [
    /(\b(SELECT|INSERT|UPDATE|DELETE|DROP|CREATE|ALTER|EXEC|EXECUTE)\b)/gi,
    /(\bUNION\b.*\bSELECT\b)/gi,
    /(\bOR\b.*=.*)/gi,
    /('.*OR.*'.*=.*')/gi,
    /(;.*--)/g,
    /(\bxp_cmdshell\b)/gi
  ];
  
  const checkInput = (input) => {
    if (typeof input === 'string') {
      return sqlPatterns.some(pattern => pattern.test(input));
    }
    if (Array.isArray(input)) {
      return input.some(checkInput);
    }
    if (typeof input === 'object' && input !== null) {
      return Object.values(input).some(checkInput);
    }
    return false;
  };
  
  const inputs = [req.body, req.query, req.params];
  const hasSQLInjection = inputs.some(checkInput);
  
  if (hasSQLInjection) {
    logger.error(`SQL injection attempt detected from IP ${req.ip}`);
    return res.status(400).json({
      error: 'Bad Request',
      message: 'Invalid characters detected'
    });
  }
  
  next();
};

module.exports = {
  inputValidation,
  sanitizeInput,
  chatMessageValidation,
  userRegistrationValidation,
  loginValidation,
  fileUploadValidation,
  handleValidationErrors,
  sqlInjectionDetection
};