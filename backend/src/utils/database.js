/**
 * 🔐 データベース接続ユーティリティ
 * MongoDB接続とセキュリティ設定
 */

const mongoose = require('mongoose');
const { logger } = require('./logger');

/**
 * MongoDB接続設定
 */
const connectDB = async () => {
  try {
    const mongoUri = process.env.MONGODB_URI || 'mongodb://localhost:27017/excel-chatbot';
    
    const options = {
      // 接続プールの設定
      maxPoolSize: 10,
      minPoolSize: 2,
      
      // タイムアウト設定
      serverSelectionTimeoutMS: 5000,
      socketTimeoutMS: 45000,
      connectTimeoutMS: 10000,
      
      // セキュリティ設定
      authSource: 'admin',
      
      // パフォーマンス設定
      bufferMaxEntries: 0,
      
      // デバッグ設定
      family: 4 // IPv4を使用
    };
    
    await mongoose.connect(mongoUri, options);
    
    logger.info('✅ MongoDB接続成功');
    
    // 接続イベント監視
    mongoose.connection.on('disconnected', () => {
      logger.warn('📡 MongoDB接続切断');
    });
    
    mongoose.connection.on('error', (err) => {
      logger.error('❌ MongoDB接続エラー:', err);
    });
    
    mongoose.connection.on('reconnected', () => {
      logger.info('🔄 MongoDB再接続完了');
    });
    
    // Graceful shutdown
    process.on('SIGINT', async () => {
      try {
        await mongoose.connection.close();
        logger.info('🔒 MongoDB接続を正常にクローズしました');
        process.exit(0);
      } catch (err) {
        logger.error('❌ MongoDB接続クローズエラー:', err);
        process.exit(1);
      }
    });
    
  } catch (error) {
    logger.error('❌ MongoDB接続失敗:', error);
    process.exit(1);
  }
};

/**
 * データベース初期化
 */
const initializeDatabase = async () => {
  try {
    // インデックス作成
    await createIndexes();
    
    // 初期データ投入
    await seedInitialData();
    
    logger.info('✅ データベース初期化完了');
  } catch (error) {
    logger.error('❌ データベース初期化エラー:', error);
    throw error;
  }
};

/**
 * インデックス作成
 */
const createIndexes = async () => {
  try {
    // User コレクションのインデックス
    await mongoose.connection.db.collection('users').createIndex(
      { email: 1 }, 
      { unique: true, background: true }
    );
    
    await mongoose.connection.db.collection('users').createIndex(
      { username: 1 }, 
      { unique: true, background: true }
    );
    
    // セキュリティ関連インデックス
    await mongoose.connection.db.collection('users').createIndex(
      { emailVerificationToken: 1 }, 
      { background: true, sparse: true }
    );
    
    await mongoose.connection.db.collection('users').createIndex(
      { passwordResetToken: 1 }, 
      { background: true, sparse: true }
    );
    
    // Excel知識ベースのインデックス
    await mongoose.connection.db.collection('excelfunctions').createIndex(
      { name: 'text', description: 'text', tags: 'text' },
      { background: true }
    );
    
    await mongoose.connection.db.collection('excelfeatures').createIndex(
      { name: 'text', description: 'text' },
      { background: true }
    );
    
    logger.info('✅ インデックス作成完了');
  } catch (error) {
    logger.error('❌ インデックス作成エラー:', error);
    throw error;
  }
};

/**
 * 初期データ投入
 */
const seedInitialData = async () => {
  try {
    const { ExcelFunction, ExcelFeature, BestPractice } = require('../models/ExcelKnowledge');
    const User = require('../models/User');
    
    // 管理者アカウント作成（存在しない場合）
    const adminExists = await User.findOne({ role: 'admin' });
    if (!adminExists) {
      const admin = new User({
        username: 'admin',
        email: 'admin@excel-chatbot.com',
        password: process.env.ADMIN_PASSWORD || 'SecureAdmin123!',
        role: 'admin',
        isEmailVerified: true,
        firstName: 'System',
        lastName: 'Administrator'
      });
      
      await admin.save();
      logger.info('✅ 管理者アカウント作成完了');
    }
    
    // 基本Excel関数データの投入
    const functionCount = await ExcelFunction.countDocuments();
    if (functionCount === 0) {
      await seedExcelFunctions();
      logger.info('✅ Excel関数データ投入完了');
    }
    
    // Excel機能データの投入
    const featureCount = await ExcelFeature.countDocuments();
    if (featureCount === 0) {
      await seedExcelFeatures();
      logger.info('✅ Excel機能データ投入完了');
    }
    
    // ベストプラクティスデータの投入
    const practiceCount = await BestPractice.countDocuments();
    if (practiceCount === 0) {
      await seedBestPractices();
      logger.info('✅ ベストプラクティスデータ投入完了');
    }
    
  } catch (error) {
    logger.error('❌ 初期データ投入エラー:', error);
    throw error;
  }
};

/**
 * Excel関数の初期データ
 */
const seedExcelFunctions = async () => {
  const { ExcelFunction } = require('../models/ExcelKnowledge');
  
  const functions = [
    {
      name: 'SUM',
      category: 'MATH',
      syntax: 'SUM(number1, [number2], ...)',
      description: '数値の合計を計算します',
      parameters: [
        { name: 'number1', required: true, description: '合計する最初の数値または範囲', type: 'number' },
        { name: 'number2', required: false, description: '合計する追加の数値または範囲', type: 'number' }
      ],
      examples: [
        { formula: '=SUM(A1:A10)', description: 'A1からA10の範囲の合計', result: '合計値' },
        { formula: '=SUM(1,2,3,4,5)', description: '数値の直接指定', result: '15' }
      ],
      relatedFunctions: ['SUMIF', 'SUMIFS', 'AVERAGE'],
      tips: ['範囲指定時は絶対参照($)を使うと便利', '非数値は自動的に無視される'],
      commonErrors: ['#VALUE! - テキストが含まれている場合'],
      tags: ['基本', '数学', '集計']
    },
    {
      name: 'VLOOKUP',
      category: 'LOOKUP_REFERENCE',
      syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
      description: '垂直方向の検索を行い、対応する値を返します',
      parameters: [
        { name: 'lookup_value', required: true, description: '検索する値', type: 'text' },
        { name: 'table_array', required: true, description: '検索テーブル', type: 'reference' },
        { name: 'col_index_num', required: true, description: '戻り値の列番号', type: 'number' },
        { name: 'range_lookup', required: false, description: 'TRUE:近似一致, FALSE:完全一致', type: 'boolean' }
      ],
      examples: [
        { formula: '=VLOOKUP(A2,C:F,2,FALSE)', description: 'A2の値をC列で検索し、D列の値を返す', result: '対応する値' }
      ],
      relatedFunctions: ['HLOOKUP', 'INDEX', 'MATCH', 'XLOOKUP'],
      tips: ['完全一致にはFALSEを使用', '検索列は表の一番左に配置'],
      commonErrors: ['#N/A - 値が見つからない', '#REF! - 列番号が範囲外'],
      tags: ['検索', '参照', '中級']
    }
  ];
  
  await ExcelFunction.insertMany(functions);
};

/**
 * Excel機能の初期データ
 */
const seedExcelFeatures = async () => {
  const { ExcelFeature } = require('../models/ExcelKnowledge');
  
  const features = [
    {
      name: 'ピボットテーブル',
      category: 'PIVOT_TABLES',
      description: 'データの集計・分析を行う強力な機能です',
      howTo: [
        { step: 1, instruction: 'データ範囲を選択する' },
        { step: 2, instruction: '挿入タブ → ピボットテーブルをクリック' },
        { step: 3, instruction: 'フィールドを適切な領域にドラッグ' }
      ],
      prerequisites: ['データが表形式で整理されている', 'ヘッダー行がある'],
      tips: ['データは連続した範囲にする', '空白行・列は避ける'],
      troubleshooting: [
        { problem: 'データが表示されない', solution: 'データ範囲を確認し、更新ボタンを押す' }
      ],
      relatedFeatures: ['条件付き書式', 'グラフ'],
      difficulty: 'INTERMEDIATE'
    }
  ];
  
  await ExcelFeature.insertMany(features);
};

/**
 * ベストプラクティスの初期データ
 */
const seedBestPractices = async () => {
  const { BestPractice } = require('../models/ExcelKnowledge');
  
  const practices = [
    {
      title: 'セル参照の適切な使い方',
      category: 'FORMULAS',
      description: '相対参照と絶対参照を正しく使い分けることで、効率的な数式を作成できます',
      dos: [
        '固定値には絶対参照($A$1)を使用する',
        '数式をコピーする前に参照方式を確認する',
        '範囲名を活用して可読性を向上させる'
      ],
      donts: [
        '不要な絶対参照を多用しない',
        '循環参照を作らない',
        '他のシートへの参照を過度に使わない'
      ],
      examples: [
        {
          good: '=$A$1*B1',
          bad: '=A1*B1',
          explanation: '税率など固定値は絶対参照にする'
        }
      ],
      applicableScenarios: ['数式のコピー', '税率計算', 'テンプレート作成'],
      importance: 'HIGH'
    }
  ];
  
  await BestPractice.insertMany(practices);
};

module.exports = {
  connectDB,
  initializeDatabase
};