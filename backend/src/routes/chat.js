/**
 * 💬 チャット関連ルート
 * Excel専門AIチャットボットのメイン機能
 */

const express = require('express');
const rateLimit = require('express-rate-limit');
const { 
  chatMessageValidation, 
  handleValidationErrors 
} = require('../middleware/validation');
const { authMiddleware } = require('../middleware/auth');
const { criticalActionAudit } = require('../middleware/audit');
const { ExcelFunction, ExcelFeature, BestPractice, ExcelFAQ } = require('../models/ExcelKnowledge');
const { logChatInteraction, logUserAction } = require('../utils/logger');

const router = express.Router();

// チャット専用レート制限
const chatLimiter = rateLimit({
  windowMs: 60 * 1000, // 1分
  max: 10, // 1分間に10メッセージ
  message: {
    error: 'Too many messages, please slow down',
    code: 'CHAT_RATE_LIMIT_EXCEEDED'
  },
  keyGenerator: (req) => {
    return `chat:${req.user?.id || req.ip}`;
  }
});

/**
 * Excel知識検索
 */
const searchExcelKnowledge = async (query, category = null) => {
  const searchOptions = {
    $text: { $search: query },
    ...(category && { category })
  };
  
  const [functions, features, practices, faqs] = await Promise.all([
    ExcelFunction.find(searchOptions).limit(5).lean(),
    ExcelFeature.find(searchOptions).limit(3).lean(),
    BestPractice.find(searchOptions).limit(3).lean(),
    ExcelFAQ.find(searchOptions).limit(5).lean()
  ]);
  
  return {
    functions,
    features,
    practices,
    faqs,
    totalResults: functions.length + features.length + practices.length + faqs.length
  };
};

/**
 * AIレスポンス生成（模擬）
 */
const generateAIResponse = async (message, knowledgeBase, user) => {
  const startTime = Date.now();
  
  try {
    // 実際のAI APIを呼び出す場合のプレースホルダー
    // const response = await callOpenAI(message, knowledgeBase);
    
    // デモンストレーション用の模擬レスポンス
    let response = "申し訳ございませんが、現在AI機能は開発中です。\n\n";
    
    if (knowledgeBase.totalResults > 0) {
      response += "関連する情報を見つけました：\n\n";
      
      if (knowledgeBase.functions.length > 0) {
        response += "📊 **関連するExcel関数:**\n";
        knowledgeBase.functions.forEach(func => {
          response += `• **${func.name}**: ${func.description}\n`;
          response += `  構文: \`${func.syntax}\`\n\n`;
        });
      }
      
      if (knowledgeBase.features.length > 0) {
        response += "🔧 **関連するExcel機能:**\n";
        knowledgeBase.features.forEach(feature => {
          response += `• **${feature.name}**: ${feature.description}\n\n`;
        });
      }
      
      if (knowledgeBase.practices.length > 0) {
        response += "💡 **ベストプラクティス:**\n";
        knowledgeBase.practices.forEach(practice => {
          response += `• **${practice.title}**: ${practice.description}\n\n`;
        });
      }
    } else {
      response += "申し訳ございませんが、お探しの情報が見つかりませんでした。\n";
      response += "より具体的なExcel関連のキーワードを使ってもう一度お試しください。\n\n";
      response += "例：\n";
      response += "• 「VLOOKUP関数の使い方」\n";
      response += "• 「ピボットテーブルの作成方法」\n";
      response += "• 「条件付き書式の設定」";
    }
    
    const responseTime = Date.now() - startTime;
    
    // 使用量追跡
    if (user) {
      await user.incrementUsage();
    }
    
    return {
      message: response,
      responseTime,
      knowledgeResults: knowledgeBase.totalResults,
      model: 'excel-knowledge-base'
    };
    
  } catch (error) {
    throw new Error(`AI response generation failed: ${error.message}`);
  }
};

/**
 * チャットメッセージ送信
 */
router.post('/message',
  authMiddleware,
  chatLimiter,
  chatMessageValidation,
  handleValidationErrors,
  criticalActionAudit('CHAT_MESSAGE'),
  async (req, res) => {
    const startTime = Date.now();
    
    try {
      const { message, conversationId } = req.body;
      const user = req.user;
      
      // 使用制限チェック
      if (!user.canAskQuestion()) {
        return res.status(429).json({
          error: 'Daily limit exceeded',
          message: 'You have reached your daily question limit',
          dailyLimit: user.usage.dailyLimit,
          questionsAsked: user.usage.questionsAsked
        });
      }
      
      // Excel関連キーワードの検出
      const excelKeywords = [
        'excel', 'spreadsheet', 'formula', 'function', 'cell', 'row', 'column',
        'pivot', 'chart', 'macro', 'vba', 'worksheet', 'workbook', 'sum',
        'vlookup', 'hlookup', 'if', 'countif', 'sumif', 'index', 'match',
        '関数', 'エクセル', '数式', 'ピボット', 'グラフ', 'マクロ'
      ];
      
      const hasExcelKeywords = excelKeywords.some(keyword =>
        message.toLowerCase().includes(keyword)
      );
      
      if (!hasExcelKeywords) {
        return res.status(400).json({
          error: 'Non-Excel question',
          message: 'This chatbot specializes in Excel questions only. Please ask questions related to Excel functions, features, or best practices.',
          suggestion: 'Try asking about Excel functions like VLOOKUP, pivot tables, or chart creation.'
        });
      }
      
      // 知識ベース検索
      const knowledgeBase = await searchExcelKnowledge(message);
      
      // AIレスポンス生成
      const aiResponse = await generateAIResponse(message, knowledgeBase, user);
      
      const responseTime = Date.now() - startTime;
      
      // ログ記録
      logChatInteraction(
        user.id,
        conversationId || 'new',
        message.length,
        responseTime,
        aiResponse.model
      );
      
      logUserAction(
        user.id,
        'CHAT_MESSAGE_PROCESSED',
        '/api/chat/message',
        req.ip,
        req.headers['user-agent'],
        {
          messageLength: message.length,
          responseTime,
          knowledgeResults: aiResponse.knowledgeResults,
          hasExcelKeywords
        }
      );
      
      res.json({
        response: aiResponse.message,
        metadata: {
          responseTime: aiResponse.responseTime,
          knowledgeResults: aiResponse.knowledgeResults,
          model: aiResponse.model,
          conversationId: conversationId || `conv_${Date.now()}`,
          remainingQuestions: user.usage.dailyLimit - user.usage.questionsAsked
        }
      });
      
    } catch (error) {
      const responseTime = Date.now() - startTime;
      
      logUserAction(
        req.user?.id || 'unknown',
        'CHAT_MESSAGE_ERROR',
        '/api/chat/message',
        req.ip,
        req.headers['user-agent'],
        {
          error: error.message,
          responseTime
        }
      );
      
      res.status(500).json({
        error: 'Chat processing failed',
        message: 'An error occurred while processing your message',
        conversationId: req.body.conversationId
      });
    }
  }
);

/**
 * Excel関数一覧取得
 */
router.get('/excel-functions',
  authMiddleware,
  async (req, res) => {
    try {
      const { category, search, page = 1, limit = 20 } = req.query;
      
      const filter = {};
      if (category) filter.category = category;
      if (search) filter.$text = { $search: search };
      
      const functions = await ExcelFunction.find(filter)
        .select('name category description syntax tags')
        .limit(limit * 1)
        .skip((page - 1) * limit)
        .sort({ name: 1 })
        .lean();
      
      const total = await ExcelFunction.countDocuments(filter);
      
      res.json({
        functions,
        pagination: {
          page: parseInt(page),
          limit: parseInt(limit),
          total,
          pages: Math.ceil(total / limit)
        }
      });
      
    } catch (error) {
      res.status(500).json({
        error: 'Failed to fetch Excel functions',
        message: 'An error occurred while retrieving Excel functions'
      });
    }
  }
);

/**
 * 特定のExcel関数詳細取得
 */
router.get('/excel-functions/:name',
  authMiddleware,
  async (req, res) => {
    try {
      const { name } = req.params;
      
      const func = await ExcelFunction.findOne({ 
        name: name.toUpperCase() 
      }).lean();
      
      if (!func) {
        return res.status(404).json({
          error: 'Function not found',
          message: `Excel function '${name}' not found in knowledge base`
        });
      }
      
      res.json(func);
      
    } catch (error) {
      res.status(500).json({
        error: 'Failed to fetch function details',
        message: 'An error occurred while retrieving function details'
      });
    }
  }
);

/**
 * Excel機能一覧取得
 */
router.get('/excel-features',
  authMiddleware,
  async (req, res) => {
    try {
      const { category, difficulty, page = 1, limit = 10 } = req.query;
      
      const filter = {};
      if (category) filter.category = category;
      if (difficulty) filter.difficulty = difficulty;
      
      const features = await ExcelFeature.find(filter)
        .select('name category description difficulty')
        .limit(limit * 1)
        .skip((page - 1) * limit)
        .sort({ name: 1 })
        .lean();
      
      const total = await ExcelFeature.countDocuments(filter);
      
      res.json({
        features,
        pagination: {
          page: parseInt(page),
          limit: parseInt(limit),
          total,
          pages: Math.ceil(total / limit)
        }
      });
      
    } catch (error) {
      res.status(500).json({
        error: 'Failed to fetch Excel features',
        message: 'An error occurred while retrieving Excel features'
      });
    }
  }
);

/**
 * チャット履歴取得
 */
router.get('/history',
  authMiddleware,
  async (req, res) => {
    try {
      const { page = 1, limit = 10 } = req.query;
      const userId = req.user.id;
      
      // 実際の実装では会話履歴モデルから取得
      // const conversations = await Conversation.find({ userId })...
      
      // デモ用の空レスポンス
      res.json({
        conversations: [],
        pagination: {
          page: parseInt(page),
          limit: parseInt(limit),
          total: 0,
          pages: 0
        }
      });
      
    } catch (error) {
      res.status(500).json({
        error: 'Failed to fetch chat history',
        message: 'An error occurred while retrieving chat history'
      });
    }
  }
);

module.exports = router;