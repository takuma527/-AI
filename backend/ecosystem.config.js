/**
 * 🔐 PM2 設定ファイル
 * 本番環境用のプロセス管理設定
 */

module.exports = {
  apps: [
    {
      name: 'excel-chatbot-backend',
      script: 'src/server.js',
      cwd: '/home/user/webapp/backend',
      instances: 'max', // CPUコア数に応じて自動調整
      exec_mode: 'cluster', // クラスターモード
      
      // 環境設定
      env: {
        NODE_ENV: 'development',
        PORT: 5000
      },
      env_production: {
        NODE_ENV: 'production',
        PORT: 5000
      },
      
      // ログ設定
      log_file: '/home/user/webapp/logs/combined.log',
      out_file: '/home/user/webapp/logs/out.log',
      error_file: '/home/user/webapp/logs/error.log',
      log_date_format: 'YYYY-MM-DD HH:mm:ss Z',
      
      // プロセス管理
      max_memory_restart: '1G', // メモリ1GB超過で再起動
      restart_delay: 4000, // 再起動間隔
      max_restarts: 5, // 最大再起動回数
      min_uptime: '10s', // 最小稼働時間
      
      // 監視設定
      watch: false, // 本番では無効
      ignore_watch: ['node_modules', 'logs', '.git'],
      
      // その他設定
      merge_logs: true,
      time: true,
      autorestart: true,
      
      // クラスター設定
      listen_timeout: 3000,
      kill_timeout: 5000
    }
  ]
};