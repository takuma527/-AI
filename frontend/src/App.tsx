/**
 * 🔐 メイン App コンポーネント
 * アプリケーションのルーティングと基本レイアウト
 */

import React, { Suspense } from 'react';
import { Routes, Route, Navigate } from 'react-router-dom';
import { Box, CircularProgress, Container } from '@mui/material';
import { Helmet } from 'react-helmet-async';

import { useAuth } from './contexts/AuthContext';
import { useSecurity } from './contexts/SecurityContext';
import Navbar from './components/Common/Navbar';
import Footer from './components/Common/Footer';
import SecurityBanner from './components/Common/SecurityBanner';
import LoadingScreen from './components/Common/LoadingScreen';
import ErrorBoundary from './components/Common/ErrorBoundary';

// Lazy loading でコンポーネントを読み込み
const Login = React.lazy(() => import('./components/Auth/Login'));
const Register = React.lazy(() => import('./components/Auth/Register'));
const Dashboard = React.lazy(() => import('./components/Dashboard/Dashboard'));
const ChatInterface = React.lazy(() => import('./components/Chat/ChatInterface'));
const ExcelKnowledge = React.lazy(() => import('./components/ExcelKnowledge/ExcelKnowledge'));
const Profile = React.lazy(() => import('./components/Profile/Profile'));
const Admin = React.lazy(() => import('./components/Admin/Admin'));
const NotFound = React.lazy(() => import('./components/Common/NotFound'));

/**
 * 認証が必要なルートのラッパー
 */
const ProtectedRoute: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { isAuthenticated, loading } = useAuth();

  if (loading) {
    return <LoadingScreen />;
  }

  if (!isAuthenticated) {
    return <Navigate to="/login" replace />;
  }

  return <>{children}</>;
};

/**
 * 管理者専用ルートのラッパー
 */
const AdminRoute: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { user, isAuthenticated } = useAuth();

  if (!isAuthenticated || user?.role !== 'admin') {
    return <Navigate to="/dashboard" replace />;
  }

  return <>{children}</>;
};

/**
 * ゲスト専用ルート（ログイン済みなら Dashboard へリダイレクト）
 */
const GuestRoute: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { isAuthenticated, loading } = useAuth();

  if (loading) {
    return <LoadingScreen />;
  }

  if (isAuthenticated) {
    return <Navigate to="/dashboard" replace />;
  }

  return <>{children}</>;
};

/**
 * サスペンス用のローディングコンポーネント
 */
const SuspenseLoader: React.FC = () => (
  <Box
    display="flex"
    justifyContent="center"
    alignItems="center"
    minHeight="60vh"
  >
    <CircularProgress size={60} />
  </Box>
);

/**
 * メイン App コンポーネント
 */
const App: React.FC = () => {
  const { securityLevel, threats } = useSecurity();
  const { isAuthenticated } = useAuth();

  return (
    <ErrorBoundary>
      <Helmet>
        <title>🔐 セキュア Excel チャットボット</title>
        <meta 
          name="description" 
          content="セキュリティを重視したExcel専門AIチャットボット。Excel関数、機能、ベストプラクティスについて専門的なアドバイスを提供します。"
        />
        <meta name="keywords" content="Excel, チャットボット, AI, セキュリティ, 関数, 自動化" />
        <meta name="author" content="Excel Chatbot Team" />
        <link rel="canonical" href={window.location.href} />
      </Helmet>

      <Box
        sx={{
          display: 'flex',
          flexDirection: 'column',
          minHeight: '100vh',
          bgcolor: 'background.default',
        }}
      >
        {/* セキュリティ警告バナー */}
        {(securityLevel === 'HIGH' || threats.length > 0) && (
          <SecurityBanner level={securityLevel} threats={threats} />
        )}

        {/* ナビゲーションバー */}
        <Navbar />

        {/* メインコンテンツ */}
        <Container
          component="main"
          maxWidth="lg"
          sx={{
            flex: 1,
            py: 4,
            px: { xs: 2, md: 3 },
          }}
        >
          <Suspense fallback={<SuspenseLoader />}>
            <Routes>
              {/* パブリックルート */}
              <Route
                path="/login"
                element={
                  <GuestRoute>
                    <Login />
                  </GuestRoute>
                }
              />
              <Route
                path="/register"
                element={
                  <GuestRoute>
                    <Register />
                  </GuestRoute>
                }
              />

              {/* 保護されたルート */}
              <Route
                path="/dashboard"
                element={
                  <ProtectedRoute>
                    <Dashboard />
                  </ProtectedRoute>
                }
              />
              <Route
                path="/chat"
                element={
                  <ProtectedRoute>
                    <ChatInterface />
                  </ProtectedRoute>
                }
              />
              <Route
                path="/knowledge"
                element={
                  <ProtectedRoute>
                    <ExcelKnowledge />
                  </ProtectedRoute>
                }
              />
              <Route
                path="/profile"
                element={
                  <ProtectedRoute>
                    <Profile />
                  </ProtectedRoute>
                }
              />

              {/* 管理者専用ルート */}
              <Route
                path="/admin/*"
                element={
                  <AdminRoute>
                    <Admin />
                  </AdminRoute>
                }
              />

              {/* ルートパス */}
              <Route
                path="/"
                element={
                  isAuthenticated ? (
                    <Navigate to="/dashboard" replace />
                  ) : (
                    <Navigate to="/login" replace />
                  )
                }
              />

              {/* 404 Not Found */}
              <Route path="*" element={<NotFound />} />
            </Routes>
          </Suspense>
        </Container>

        {/* フッター */}
        <Footer />
      </Box>
    </ErrorBoundary>
  );
};

export default App;