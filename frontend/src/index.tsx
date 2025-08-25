/**
 * 🔐 セキュア Excel チャットボット - エントリーポイント
 * React アプリケーションのメインエントリー
 */

import React from 'react';
import ReactDOM from 'react-dom/client';
import { BrowserRouter } from 'react-router-dom';
import { HelmetProvider } from 'react-helmet-async';
import { ThemeProvider, createTheme } from '@mui/material/styles';
import { CssBaseline } from '@mui/material';
import { SnackbarProvider } from 'notistack';

import App from './App';
import { AuthProvider } from './contexts/AuthContext';
import { ChatProvider } from './contexts/ChatContext';
import { SecurityProvider } from './contexts/SecurityContext';
import reportWebVitals from './reportWebVitals';

// セキュアなテーマ設定
const theme = createTheme({
  palette: {
    primary: {
      main: '#1976d2',
      dark: '#1565c0',
      light: '#42a5f5',
    },
    secondary: {
      main: '#dc004e',
    },
    background: {
      default: '#f5f5f5',
      paper: '#ffffff',
    },
    text: {
      primary: '#333333',
      secondary: '#666666',
    },
  },
  typography: {
    fontFamily: '"Noto Sans JP", "Roboto", "Helvetica", "Arial", sans-serif',
    h1: {
      fontSize: '2.125rem',
      fontWeight: 500,
    },
    h2: {
      fontSize: '1.75rem',
      fontWeight: 500,
    },
    h3: {
      fontSize: '1.5rem',
      fontWeight: 500,
    },
    body1: {
      fontSize: '1rem',
      lineHeight: 1.6,
    },
    body2: {
      fontSize: '0.875rem',
      lineHeight: 1.5,
    },
  },
  components: {
    MuiButton: {
      styleOverrides: {
        root: {
          textTransform: 'none',
          borderRadius: '8px',
        },
      },
    },
    MuiTextField: {
      styleOverrides: {
        root: {
          '& .MuiOutlinedInput-root': {
            borderRadius: '8px',
          },
        },
      },
    },
    MuiPaper: {
      styleOverrides: {
        root: {
          borderRadius: '12px',
        },
      },
    },
  },
});

// セキュリティ監視
const securityMonitoring = () => {
  // コンソール操作の検知
  const originalConsole = console.log;
  console.log = (...args) => {
    if (process.env.NODE_ENV === 'production') {
      // 本番環境では制限された情報のみログ出力
      return;
    }
    originalConsole.apply(console, args);
  };
  
  // 開発者ツールの検知（簡易版）
  let devtools = { open: false, orientation: null };
  
  setInterval(() => {
    if (window.outerHeight - window.innerHeight > 160 && !devtools.open) {
      devtools.open = true;
      devtools.orientation = 'horizontal';
      if (process.env.NODE_ENV === 'production') {
        console.warn('Developer tools detected');
      }
    } else if (window.outerWidth - window.innerWidth > 160 && !devtools.open) {
      devtools.open = true;
      devtools.orientation = 'vertical';
      if (process.env.NODE_ENV === 'production') {
        console.warn('Developer tools detected');
      }
    } else if (
      window.outerHeight - window.innerHeight <= 160 &&
      window.outerWidth - window.innerWidth <= 160 &&
      devtools.open
    ) {
      devtools.open = false;
      devtools.orientation = null;
    }
  }, 500);
};

// セキュリティ監視の開始
if (process.env.NODE_ENV === 'production') {
  securityMonitoring();
}

// React 18 の厳格モード対応
const root = ReactDOM.createRoot(
  document.getElementById('root') as HTMLElement
);

root.render(
  <React.StrictMode>
    <HelmetProvider>
      <BrowserRouter>
        <ThemeProvider theme={theme}>
          <CssBaseline />
          <SnackbarProvider 
            maxSnack={3}
            anchorOrigin={{
              vertical: 'top',
              horizontal: 'right',
            }}
            autoHideDuration={5000}
          >
            <SecurityProvider>
              <AuthProvider>
                <ChatProvider>
                  <App />
                </ChatProvider>
              </AuthProvider>
            </SecurityProvider>
          </SnackbarProvider>
        </ThemeProvider>
      </BrowserRouter>
    </HelmetProvider>
  </React.StrictMode>
);

// パフォーマンス測定
reportWebVitals((metric) => {
  if (process.env.NODE_ENV === 'production') {
    // 本番環境では分析サービスに送信
    console.log(metric);
  }
});