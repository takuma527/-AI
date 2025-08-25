/**
 * 🔐 PM2 設定ファイル（簡易版）
 * デモンストレーション用の設定
 */

module.exports = {
  apps: [
    {
      name: 'excel-chatbot-simple',
      script: 'src/server-simple.js',
      cwd: '/home/user/webapp/backend',
      instances: 1, // シンプル版は1インスタンス
      exec_mode: 'fork', // フォークモード
      
      // 環境設定
      env: {
        NODE_ENV: 'development',
        PORT: 5000
      },
      
      // ログ設定
      log_file: '/home/user/webapp/logs/combined.log',
      out_file: '/home/user/webapp/logs/out.log',
      error_file: '/home/user/webapp/logs/error.log',
      log_date_format: 'YYYY-MM-DD HH:mm:ss Z',
      
      // プロセス管理
      max_memory_restart: '500M',
      restart_delay: 1000,
      max_restarts: 3,
      min_uptime: '5s',
      
      // 監視設定
      watch: false,
      
      // その他設定
      merge_logs: true,
      time: true,
      autorestart: true
    }
  ]
};