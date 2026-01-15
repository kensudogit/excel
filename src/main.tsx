import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.tsx'
import './index.css'

// ブラウザ拡張機能のエラーを無視（React DevToolsなど）
const originalError = console.error
console.error = (...args) => {
  const errorMessage = args[0]?.toString() || ''
  // ブラウザ拡張機能のエラーを無視
  if (
    errorMessage.includes('message channel closed') ||
    errorMessage.includes('asynchronous response') ||
    errorMessage.includes('Extension context invalidated')
  ) {
    return
  }
  originalError.apply(console, args)
}

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
