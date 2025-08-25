/**
 * 📊 Excel知識データベースモデル
 * Excel関数、機能、ベストプラクティスのデータ構造
 */

const mongoose = require('mongoose');

// Excel関数スキーマ
const excelFunctionSchema = new mongoose.Schema({
  name: {
    type: String,
    required: true,
    unique: true,
    uppercase: true
  },
  category: {
    type: String,
    required: true,
    enum: [
      'MATH', 'STATISTICAL', 'LOGICAL', 'TEXT', 'DATE_TIME', 
      'LOOKUP_REFERENCE', 'DATABASE', 'FINANCIAL', 'INFORMATION', 
      'ENGINEERING', 'CUBE', 'WEB', 'COMPATIBILITY'
    ]
  },
  syntax: {
    type: String,
    required: true
  },
  description: {
    type: String,
    required: true
  },
  parameters: [{
    name: String,
    required: Boolean,
    description: String,
    type: String // number, text, reference, array, etc.
  }],
  examples: [{
    formula: String,
    description: String,
    result: String
  }],
  relatedFunctions: [String],
  version: {
    type: String,
    default: 'All'
  },
  tips: [String],
  commonErrors: [String],
  tags: [String]
}, {
  timestamps: true
});

// Excel機能スキーマ
const excelFeatureSchema = new mongoose.Schema({
  name: {
    type: String,
    required: true,
    unique: true
  },
  category: {
    type: String,
    required: true,
    enum: [
      'CHARTS', 'PIVOT_TABLES', 'CONDITIONAL_FORMATTING', 'DATA_VALIDATION',
      'MACROS_VBA', 'POWER_QUERY', 'POWER_PIVOT', 'SOLVER', 'ANALYSIS_TOOLPAK',
      'FORMS', 'PROTECTION', 'COLLABORATION', 'TEMPLATES'
    ]
  },
  description: {
    type: String,
    required: true
  },
  howTo: [{
    step: Number,
    instruction: String,
    screenshot: String // URL to screenshot
  }],
  prerequisites: [String],
  tips: [String],
  troubleshooting: [{
    problem: String,
    solution: String
  }],
  relatedFeatures: [String],
  version: String,
  difficulty: {
    type: String,
    enum: ['BEGINNER', 'INTERMEDIATE', 'ADVANCED'],
    default: 'BEGINNER'
  }
}, {
  timestamps: true
});

// Excel VBAスキーマ
const excelVBASchema = new mongoose.Schema({
  name: {
    type: String,
    required: true
  },
  category: {
    type: String,
    required: true,
    enum: [
      'BASIC_SYNTAX', 'VARIABLES', 'FUNCTIONS', 'PROCEDURES', 'OBJECTS',
      'EVENTS', 'ERROR_HANDLING', 'FILE_OPERATIONS', 'USER_FORMS',
      'AUTOMATION', 'ADVANCED'
    ]
  },
  description: {
    type: String,
    required: true
  },
  code: {
    type: String,
    required: true
  },
  explanation: [String],
  parameters: [{
    name: String,
    type: String,
    description: String,
    optional: Boolean
  }],
  returnValue: {
    type: String,
    description: String
  },
  examples: [{
    code: String,
    description: String,
    output: String
  }],
  relatedConcepts: [String],
  difficulty: {
    type: String,
    enum: ['BEGINNER', 'INTERMEDIATE', 'ADVANCED'],
    default: 'BEGINNER'
  }
}, {
  timestamps: true
});

// ベストプラクティススキーマ
const bestPracticeSchema = new mongoose.Schema({
  title: {
    type: String,
    required: true
  },
  category: {
    type: String,
    required: true,
    enum: [
      'DESIGN', 'PERFORMANCE', 'SECURITY', 'COLLABORATION', 'MAINTENANCE',
      'DATA_MANAGEMENT', 'FORMULAS', 'VISUALIZATION', 'AUTOMATION'
    ]
  },
  description: {
    type: String,
    required: true
  },
  dos: [String],
  donts: [String],
  examples: [{
    good: String,
    bad: String,
    explanation: String
  }],
  applicableScenarios: [String],
  relatedTopics: [String],
  importance: {
    type: String,
    enum: ['LOW', 'MEDIUM', 'HIGH', 'CRITICAL'],
    default: 'MEDIUM'
  }
}, {
  timestamps: true
});

// FAQ スキーマ
const excelFAQSchema = new mongoose.Schema({
  question: {
    type: String,
    required: true,
    unique: true
  },
  answer: {
    type: String,
    required: true
  },
  category: {
    type: String,
    required: true,
    enum: [
      'BASICS', 'FORMULAS', 'CHARTS', 'PIVOT_TABLES', 'MACROS', 'TROUBLESHOOTING',
      'PERFORMANCE', 'COMPATIBILITY', 'SECURITY', 'COLLABORATION'
    ]
  },
  keywords: [String],
  relatedQuestions: [String],
  difficulty: {
    type: String,
    enum: ['BEGINNER', 'INTERMEDIATE', 'ADVANCED'],
    default: 'BEGINNER'
  },
  votes: {
    helpful: { type: Number, default: 0 },
    notHelpful: { type: Number, default: 0 }
  },
  lastUpdated: {
    type: Date,
    default: Date.now
  }
}, {
  timestamps: true
});

// トラブルシューティングスキーマ
const troubleshootingSchema = new mongoose.Schema({
  problem: {
    type: String,
    required: true
  },
  symptoms: [String],
  causes: [String],
  solutions: [{
    description: String,
    steps: [String],
    success_rate: Number // 0-100
  }],
  prevention: [String],
  category: {
    type: String,
    required: true,
    enum: [
      'FORMULA_ERRORS', 'PERFORMANCE', 'COMPATIBILITY', 'FILE_CORRUPTION',
      'PRINTING', 'CHARTS', 'MACROS', 'SECURITY', 'FORMATTING'
    ]
  },
  severity: {
    type: String,
    enum: ['LOW', 'MEDIUM', 'HIGH', 'CRITICAL'],
    default: 'MEDIUM'
  },
  affectedVersions: [String]
}, {
  timestamps: true
});

// インデックス設定
excelFunctionSchema.index({ name: 'text', description: 'text', tags: 'text' });
excelFeatureSchema.index({ name: 'text', description: 'text' });
excelVBASchema.index({ name: 'text', description: 'text', code: 'text' });
bestPracticeSchema.index({ title: 'text', description: 'text' });
excelFAQSchema.index({ question: 'text', answer: 'text', keywords: 'text' });
troubleshootingSchema.index({ problem: 'text', symptoms: 'text', solutions: 'text' });

// モデル作成
const ExcelFunction = mongoose.model('ExcelFunction', excelFunctionSchema);
const ExcelFeature = mongoose.model('ExcelFeature', excelFeatureSchema);
const ExcelVBA = mongoose.model('ExcelVBA', excelVBASchema);
const BestPractice = mongoose.model('BestPractice', bestPracticeSchema);
const ExcelFAQ = mongoose.model('ExcelFAQ', excelFAQSchema);
const Troubleshooting = mongoose.model('Troubleshooting', troubleshootingSchema);

module.exports = {
  ExcelFunction,
  ExcelFeature,
  ExcelVBA,
  BestPractice,
  ExcelFAQ,
  Troubleshooting
};