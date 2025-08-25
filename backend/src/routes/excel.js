/**
 * 📊 Excel知識ベース関連ルート
 * Excel関数、機能、ベストプラクティス管理
 */

const express = require('express');
const { authMiddleware, requireAdmin } = require('../middleware/auth');
const { dataChangeAudit } = require('../middleware/audit');
const { 
  ExcelFunction, 
  ExcelFeature, 
  BestPractice, 
  ExcelFAQ,
  Troubleshooting 
} = require('../models/ExcelKnowledge');

const router = express.Router();

/**
 * Excel関数カテゴリ一覧取得
 */
router.get('/functions/categories', authMiddleware, async (req, res) => {
  try {
    const categories = await ExcelFunction.distinct('category');
    res.json({ categories });
  } catch (error) {
    res.status(500).json({
      error: 'Failed to fetch categories',
      message: 'An error occurred while retrieving function categories'
    });
  }
});

/**
 * Excel関数検索
 */
router.get('/functions/search', authMiddleware, async (req, res) => {
  try {
    const { q, category, page = 1, limit = 10 } = req.query;
    
    const filter = {};
    if (q) {
      filter.$text = { $search: q };
    }
    if (category) {
      filter.category = category;
    }
    
    const functions = await ExcelFunction.find(filter)
      .select('name category description syntax tags')
      .limit(limit * 1)
      .skip((page - 1) * limit)
      .sort(q ? { score: { $meta: 'textScore' } } : { name: 1 })
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
      error: 'Search failed',
      message: 'An error occurred during function search'
    });
  }
});

/**
 * Excel機能検索
 */
router.get('/features/search', authMiddleware, async (req, res) => {
  try {
    const { q, category, difficulty, page = 1, limit = 10 } = req.query;
    
    const filter = {};
    if (q) {
      filter.$text = { $search: q };
    }
    if (category) {
      filter.category = category;
    }
    if (difficulty) {
      filter.difficulty = difficulty;
    }
    
    const features = await ExcelFeature.find(filter)
      .select('name category description difficulty')
      .limit(limit * 1)
      .skip((page - 1) * limit)
      .sort(q ? { score: { $meta: 'textScore' } } : { name: 1 })
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
      error: 'Search failed',
      message: 'An error occurred during feature search'
    });
  }
});

/**
 * FAQ検索
 */
router.get('/faq/search', authMiddleware, async (req, res) => {
  try {
    const { q, category, page = 1, limit = 10 } = req.query;
    
    const filter = {};
    if (q) {
      filter.$text = { $search: q };
    }
    if (category) {
      filter.category = category;
    }
    
    const faqs = await ExcelFAQ.find(filter)
      .select('question answer category difficulty votes')
      .limit(limit * 1)
      .skip((page - 1) * limit)
      .sort(q ? { score: { $meta: 'textScore' } } : { 'votes.helpful': -1 })
      .lean();
    
    const total = await ExcelFAQ.countDocuments(filter);
    
    res.json({
      faqs,
      pagination: {
        page: parseInt(page),
        limit: parseInt(limit),
        total,
        pages: Math.ceil(total / limit)
      }
    });
    
  } catch (error) {
    res.status(500).json({
      error: 'FAQ search failed',
      message: 'An error occurred during FAQ search'
    });
  }
});

/**
 * 統合検索（全ての知識ベースから検索）
 */
router.get('/search', authMiddleware, async (req, res) => {
  try {
    const { q, page = 1, limit = 20 } = req.query;
    
    if (!q) {
      return res.status(400).json({
        error: 'Search query required',
        message: 'Please provide a search query'
      });
    }
    
    const searchFilter = { $text: { $search: q } };
    
    const [functions, features, practices, faqs] = await Promise.all([
      ExcelFunction.find(searchFilter)
        .select('name category description syntax')
        .limit(5)
        .sort({ score: { $meta: 'textScore' } })
        .lean(),
      
      ExcelFeature.find(searchFilter)
        .select('name category description difficulty')
        .limit(3)
        .sort({ score: { $meta: 'textScore' } })
        .lean(),
      
      BestPractice.find(searchFilter)
        .select('title category description importance')
        .limit(3)
        .sort({ score: { $meta: 'textScore' } })
        .lean(),
      
      ExcelFAQ.find(searchFilter)
        .select('question answer category')
        .limit(5)
        .sort({ score: { $meta: 'textScore' } })
        .lean()
    ]);
    
    const results = {
      functions: functions.map(f => ({ ...f, type: 'function' })),
      features: features.map(f => ({ ...f, type: 'feature' })),
      practices: practices.map(p => ({ ...p, type: 'practice' })),
      faqs: faqs.map(f => ({ ...f, type: 'faq' })),
      totalResults: functions.length + features.length + practices.length + faqs.length
    };
    
    res.json(results);
    
  } catch (error) {
    res.status(500).json({
      error: 'Search failed',
      message: 'An error occurred during search'
    });
  }
});

/**
 * FAQ投票
 */
router.post('/faq/:id/vote', authMiddleware, async (req, res) => {
  try {
    const { id } = req.params;
    const { helpful } = req.body;
    
    const updateField = helpful ? 'votes.helpful' : 'votes.notHelpful';
    
    const faq = await ExcelFAQ.findByIdAndUpdate(
      id,
      { $inc: { [updateField]: 1 } },
      { new: true }
    ).select('votes');
    
    if (!faq) {
      return res.status(404).json({
        error: 'FAQ not found',
        message: 'The requested FAQ was not found'
      });
    }
    
    res.json({
      message: 'Vote recorded successfully',
      votes: faq.votes
    });
    
  } catch (error) {
    res.status(500).json({
      error: 'Vote failed',
      message: 'An error occurred while recording vote'
    });
  }
});

// 管理者専用ルート

/**
 * Excel関数追加
 */
router.post('/functions', 
  requireAdmin,
  dataChangeAudit('ADD_EXCEL_FUNCTION'),
  async (req, res) => {
    try {
      const functionData = req.body;
      const func = new ExcelFunction(functionData);
      await func.save();
      
      res.status(201).json({
        message: 'Excel function added successfully',
        function: func
      });
      
    } catch (error) {
      if (error.code === 11000) {
        return res.status(409).json({
          error: 'Function already exists',
          message: 'A function with this name already exists'
        });
      }
      
      res.status(500).json({
        error: 'Failed to add function',
        message: 'An error occurred while adding the function'
      });
    }
  }
);

/**
 * Excel機能追加
 */
router.post('/features',
  requireAdmin,
  dataChangeAudit('ADD_EXCEL_FEATURE'),
  async (req, res) => {
    try {
      const featureData = req.body;
      const feature = new ExcelFeature(featureData);
      await feature.save();
      
      res.status(201).json({
        message: 'Excel feature added successfully',
        feature
      });
      
    } catch (error) {
      res.status(500).json({
        error: 'Failed to add feature',
        message: 'An error occurred while adding the feature'
      });
    }
  }
);

/**
 * FAQ追加
 */
router.post('/faq',
  requireAdmin,
  dataChangeAudit('ADD_FAQ'),
  async (req, res) => {
    try {
      const faqData = req.body;
      const faq = new ExcelFAQ(faqData);
      await faq.save();
      
      res.status(201).json({
        message: 'FAQ added successfully',
        faq
      });
      
    } catch (error) {
      res.status(500).json({
        error: 'Failed to add FAQ',
        message: 'An error occurred while adding the FAQ'
      });
    }
  }
);

/**
 * 統計情報取得
 */
router.get('/stats', requireAdmin, async (req, res) => {
  try {
    const [
      functionCount,
      featureCount,
      practiceCount,
      faqCount,
      troubleshootingCount
    ] = await Promise.all([
      ExcelFunction.countDocuments(),
      ExcelFeature.countDocuments(),
      BestPractice.countDocuments(),
      ExcelFAQ.countDocuments(),
      Troubleshooting.countDocuments()
    ]);
    
    const stats = {
      knowledgeBase: {
        functions: functionCount,
        features: featureCount,
        practices: practiceCount,
        faqs: faqCount,
        troubleshooting: troubleshootingCount,
        total: functionCount + featureCount + practiceCount + faqCount + troubleshootingCount
      },
      lastUpdated: new Date().toISOString()
    };
    
    res.json(stats);
    
  } catch (error) {
    res.status(500).json({
      error: 'Failed to fetch statistics',
      message: 'An error occurred while retrieving statistics'
    });
  }
});

module.exports = router;