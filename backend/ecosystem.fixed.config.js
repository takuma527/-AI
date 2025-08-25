/**
 * PM2 設定ファイル（修正版サーバー用）
 */

module.exports = {
  apps: [{
    name: 'excel-chatbot-fixed',
    script: './src/server-fixed.js',
    cwd: '/home/user/webapp/backend',
    instances: 1,
    exec_mode: 'fork',
    
    // 環境変数
    env: {
      NODE_ENV: 'development',
      PORT: 5000,
      SESSION_SECRET: 'excel-chatbot-fixed-session-secret-2025'
    },
    
    // ログ設定
    log_file: '/home/user/webapp/logs/combined.log',
    out_file: '/home/user/webapp/logs/out.log',
    error_file: '/home/user/webapp/logs/error.log',
    log_date_format: 'YYYY-MM-DDTHH:mm:ss',
    
    // 再起動設定
    autorestart: true,
    watch: false,
    max_memory_restart: '1G',
    restart_delay: 4000,
    
    // プロセス管理
    kill_timeout: 5000,
    listen_timeout: 8000,
    shutdown_with_message: true,
    
    // 詳細設定
    node_args: ['--max-old-space-size=1024'],
    max_restarts: 10,
    min_uptime: '10s'
  }]
};