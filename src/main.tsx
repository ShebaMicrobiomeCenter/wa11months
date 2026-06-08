import React, {StrictMode, ReactNode} from 'react';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import './index.css';

interface Props {
  children: ReactNode;
}

interface State {
  hasError: boolean;
  error: Error | null;
}

class ErrorBoundary extends React.Component<Props, State> {
  constructor(props: Props) {
    super(props);
    this.state = {
      hasError: false,
      error: null
    };
  }

  static getDerivedStateFromError(error: Error): State {
    return {hasError: true, error};
  }

  componentDidCatch(error: Error, errorInfo: any) {
    console.error('ErrorBoundary caught an error:', error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div style={{padding: '20px', color: '#721c24', backgroundColor: '#f8d7da', border: '1px solid #f5c6cb', margin: '20px', borderRadius: '5px', fontFamily: 'sans-serif'}}>
          <h2>משהו השתבש ברמת התצוגה</h2>
          <p>{this.state.error?.message}</p>
          <pre style={{whiteSpace: 'pre-wrap', fontSize: '12px'}}>{this.state.error?.stack}</pre>
          <button onClick={() => window.location.reload()} style={{padding: '8px 16px', background: '#721c24', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer'}}>טען מחדש</button>
        </div>
      );
    }

    return this.props.children;
  }
}

console.log('main.tsx: starting render');
const rootElement = document.getElementById('root');
if (!rootElement) {
  console.error('Failed to find the root element');
} else {
  createRoot(rootElement).render(
    <StrictMode>
      <ErrorBoundary>
        <App />
      </ErrorBoundary>
    </StrictMode>,
  );
}
