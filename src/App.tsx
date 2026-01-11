import React, { useState, useEffect } from 'react'
import './App.css'
import SearchForm from './components/SearchForm'
import ResultsTable from './components/ResultsTable'
import CellDetails from './components/CellDetails'
import { SearchResult, CellDetail } from './types'

function App() {
  const [searchResults, setSearchResults] = useState<SearchResult[]>([])
  const [selectedCell, setSelectedCell] = useState<CellDetail | null>(null)
  const [isLoading, setIsLoading] = useState(false)
  const [outputFile, setOutputFile] = useState<string>('')
  const [errorMessage, setErrorMessage] = useState<string>('')
  const [errorSuggestion, setErrorSuggestion] = useState<string>('')

  // ローディング中にカーソルを変更
  useEffect(() => {
    if (isLoading) {
      document.body.style.cursor = 'wait'
      document.body.classList.add('loading')
    } else {
      document.body.style.cursor = 'default'
      document.body.classList.remove('loading')
    }
    
    // クリーンアップ
    return () => {
      document.body.style.cursor = 'default'
      document.body.classList.remove('loading')
    }
  }, [isLoading])

  const handleSearch = async (folderPath: string, keywords: string[]) => {
    setIsLoading(true)
    setSearchResults([])
    setSelectedCell(null)
    setOutputFile('')
    setErrorMessage('')
    setErrorSuggestion('')

    try {
      const response = await fetch('/api/search', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          folder_path: folderPath,
          keywords: keywords,
        }),
      })

      // レスポンスのクローンを作成して、複数回読み込めるようにする
      const responseClone = response.clone()
      
      if (!response.ok) {
        // レスポンスがエラーの場合
        let errorMessage = `HTTP error! status: ${response.status}`
        let errorSuggestion = ''
        let errorDetails: any = {}
        try {
          // レスポンスのテキストを取得してからJSON解析
          const text = await responseClone.text()
          if (text && text.trim() !== '') {
            try {
              const errorData = JSON.parse(text)
              errorMessage = errorData.error || errorMessage
              errorSuggestion = errorData.suggestion || ''
              errorDetails = {
                original_path: errorData.original_path,
                normalized_path: errorData.normalized_path,
                folder_path: errorData.folder_path,
                files_in_folder: errorData.files_in_folder
              }
            } catch (parseError) {
              // JSON解析に失敗した場合は、テキストをそのまま使用
              errorMessage = text || errorMessage
            }
          }
        } catch (e) {
          // テキスト読み込みも失敗した場合は、ステータスコードのみを使用
          errorMessage = `HTTP error! status: ${response.status}`
        }
        
        // エラーメッセージを構築
        let fullErrorMessage = errorMessage
        if (errorSuggestion) {
          fullErrorMessage += '\n\n' + errorSuggestion
        }
        if (errorDetails.original_path && errorDetails.normalized_path) {
          fullErrorMessage += `\n\n入力されたパス: ${errorDetails.original_path}`
          fullErrorMessage += `\n正規化後のパス: ${errorDetails.normalized_path}`
        }
        if (errorDetails.files_in_folder && errorDetails.files_in_folder.length > 0) {
          fullErrorMessage += `\n\nフォルダ内のファイル: ${errorDetails.files_in_folder.join(', ')}`
        }
        
        // エラーオブジェクトに追加情報を付与
        const error = new Error(fullErrorMessage) as any
        error.suggestion = errorSuggestion
        error.details = errorDetails
        throw error
      }

      // レスポンスのテキストを取得してからJSON解析
      const text = await response.text()
      if (!text || text.trim() === '') {
        throw new Error('Empty response from server')
      }

      let data
      try {
        data = JSON.parse(text)
      } catch (parseError) {
        console.error('Failed to parse JSON response:', text)
        throw new Error('Invalid JSON response from server')
      }

      if (data.success) {
        setSearchResults(data.results || [])
        setOutputFile(data.output_file || '')
        setErrorMessage('')
        setErrorSuggestion('')
      } else {
        const errorMsg = data.error || '検索に失敗しました'
        const suggestion = data.suggestion || ''
        console.error('Search error:', errorMsg)
        if (suggestion) {
          console.error('Suggestion:', suggestion)
        }
        setErrorMessage(errorMsg)
        setErrorSuggestion(suggestion)
      }
    } catch (error) {
      console.error('Search error:', error)
      let errorMsg = '不明なエラー'
      let suggestion = ''
      
      if (error instanceof SyntaxError) {
        errorMsg = 'サーバーからの応答を解析できませんでした。サーバーが正常に動作しているか確認してください。'
      } else if (error instanceof Error) {
        errorMsg = error.message
        // エラーオブジェクトにsuggestionが含まれている場合
        if ((error as any).suggestion) {
          suggestion = (error as any).suggestion
        }
      }
      
      console.error(`検索中にエラーが発生しました: ${errorMsg}`)
      if (suggestion) {
        console.error('提案:', suggestion)
      }
      
      setErrorMessage(errorMsg)
      setErrorSuggestion(suggestion)
    } finally {
      setIsLoading(false)
    }
  }

  const handleSearchWithFiles = async (files: File[], keywords: string[]) => {
    setIsLoading(true)
    setSearchResults([])
    setSelectedCell(null)
    setOutputFile('')
    setErrorMessage('')
    setErrorSuggestion('')

    try {
      // ファイルが空でないか確認
      if (!files || files.length === 0) {
        console.warn('Excelファイルが選択されていません')
        setIsLoading(false)
        return
      }

      const formData = new FormData()
      
      // キーワードをJSON文字列として追加
      formData.append('keywords', JSON.stringify(keywords))
      
      // 各Excelファイルを追加
      files.forEach((file) => {
        formData.append('files', file)
      })

      const response = await fetch('/api/search-files', {
        method: 'POST',
        body: formData,
      })

      // レスポンスのクローンを作成して、複数回読み込めるようにする
      const responseClone = response.clone()

      if (!response.ok) {
        let errorMessage = `HTTP error! status: ${response.status}`
        try {
          // レスポンスのテキストを取得してからJSON解析
          const text = await responseClone.text()
          if (text && text.trim() !== '') {
            try {
              const errorData = JSON.parse(text)
              errorMessage = errorData.error || errorMessage
            } catch (parseError) {
              // JSON解析に失敗した場合は、テキストをそのまま使用
              errorMessage = text || errorMessage
            }
          }
        } catch (e) {
          // テキスト読み込みも失敗した場合は、ステータスコードのみを使用
          errorMessage = `HTTP error! status: ${response.status}`
        }
        throw new Error(errorMessage)
      }

      // レスポンスのテキストを取得してからJSON解析
      const text = await response.text()
      if (!text || text.trim() === '') {
        throw new Error('Empty response from server')
      }

      let data
      try {
        data = JSON.parse(text)
      } catch (parseError) {
        console.error('Failed to parse JSON response:', text)
        throw new Error('Invalid JSON response from server')
      }

      if (data.success) {
        setSearchResults(data.results || [])
        setOutputFile(data.output_file || '')
        setErrorMessage('')
        setErrorSuggestion('')
      } else {
        const errorMsg = data.error || '検索に失敗しました'
        const suggestion = data.suggestion || ''
        console.error('Search error:', errorMsg)
        if (suggestion) {
          console.error('Suggestion:', suggestion)
        }
        setErrorMessage(errorMsg)
        setErrorSuggestion(suggestion)
      }
    } catch (error) {
      console.error('Search with files error:', error)
      let errorMsg = '不明なエラー'
      let suggestion = ''
      
      if (error instanceof TypeError && error.message.includes('Failed to fetch')) {
        errorMsg = 'サーバーに接続できませんでした。バックエンドサーバー（ポート5001）が起動しているか確認してください。'
      } else if (error instanceof SyntaxError) {
        errorMsg = 'サーバーからの応答を解析できませんでした。サーバーが正常に動作しているか確認してください。'
      } else if (error instanceof Error) {
        errorMsg = error.message
        if ((error as any).suggestion) {
          suggestion = (error as any).suggestion
        }
      }
      
      console.error(`検索中にエラーが発生しました: ${errorMsg}`)
      if (suggestion) {
        console.error('提案:', suggestion)
      }
      
      setErrorMessage(errorMsg)
      setErrorSuggestion(suggestion)
    } finally {
      setIsLoading(false)
    }
  }

  const handleCellClick = async (result: SearchResult) => {
    try {
      const response = await fetch('/api/get-cell-details', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          file_path: result.file,
          sheet_name: result.sheet,
          row: result.row,
          col: result.col,
          keyword: result.keyword,
          context_rows: 5,
        }),
      })

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`)
      }

      // レスポンスのテキストを取得してからJSON解析
      const text = await response.text()
      if (!text || text.trim() === '') {
        throw new Error('Empty response from server')
      }

      let data
      try {
        data = JSON.parse(text)
      } catch (parseError) {
        console.error('Failed to parse JSON response:', text)
        throw new Error('Invalid JSON response from server')
      }

      if (data.success) {
        setSelectedCell(data)
      } else {
        console.error('Cell details error:', data.error)
      }
    } catch (error) {
      console.error('Cell details error:', error)
    }
  }

  const handleOpenExcelFile = async (result: SearchResult) => {
    try {
      const response = await fetch('/api/open-excel-file', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          file_path: result.file,
          sheet_name: result.sheet,
          row: result.row,
          col: result.col,
        }),
      })

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`)
      }

      // レスポンスのテキストを取得してからJSON解析
      const text = await response.text()
      if (!text || text.trim() === '') {
        throw new Error('Empty response from server')
      }

      let data
      try {
        data = JSON.parse(text)
      } catch (parseError) {
        console.error('Failed to parse JSON response:', text)
        throw new Error('Invalid JSON response from server')
      }

      if (data.success) {
        // 成功メッセージは表示しない（ファイルが開かれるだけ）
      } else {
        console.error('Open Excel file error:', data.error)
      }
    } catch (error) {
      console.error('Open Excel file error:', error)
    }
  }

  const handleDownloadResults = () => {
    if (outputFile) {
      window.open(`/api/download-results?file_path=${encodeURIComponent(outputFile)}`, '_blank')
    }
  }

  return (
    <div className="app">
      {isLoading && (
        <div className="loading-overlay">
          <div className="loading-spinner"></div>
        </div>
      )}
      <header className="app-header">
        <h1>📊 Excel キーワード検索</h1>
        <p>指定したフォルダ内のExcelファイルから複数のキーワードを検索します</p>
      </header>

      <main className="app-main">
        <SearchForm onSearch={handleSearch} onSearchWithFiles={handleSearchWithFiles} isLoading={isLoading} />

        {errorMessage && (
          <div className="error-message-container" style={{
            margin: '1rem 0',
            padding: '1rem',
            backgroundColor: '#fee',
            border: '1px solid #fcc',
            borderRadius: '6px',
            color: '#c33'
          }}>
            <div style={{ fontWeight: 'bold', marginBottom: '0.5rem', fontSize: '1.1rem' }}>
              ⚠️ エラーが発生しました
            </div>
            <div style={{ whiteSpace: 'pre-wrap', marginBottom: '0.5rem' }}>
              {errorMessage}
            </div>
            {errorSuggestion && (
              <div style={{
                marginTop: '0.75rem',
                padding: '0.75rem',
                backgroundColor: '#fff9e6',
                border: '1px solid #ffd700',
                borderRadius: '4px',
                whiteSpace: 'pre-wrap',
                fontSize: '0.95rem',
                color: '#856404'
              }}>
                <strong>💡 解決方法:</strong>
                <div style={{ marginTop: '0.5rem' }}>
                  {errorSuggestion}
                </div>
              </div>
            )}
            <button
              onClick={() => {
                setErrorMessage('')
                setErrorSuggestion('')
              }}
              style={{
                marginTop: '0.75rem',
                padding: '0.5rem 1rem',
                backgroundColor: '#c33',
                color: 'white',
                border: 'none',
                borderRadius: '4px',
                cursor: 'pointer',
                fontSize: '0.9rem'
              }}
            >
              閉じる
            </button>
          </div>
        )}

        {searchResults.length > 0 && (
          <div className="results-section">
            <div className="results-header">
              <h2>検索結果 ({searchResults.length}件)</h2>
              {outputFile && (
                <button onClick={handleDownloadResults} className="download-btn">
                  📥 結果をExcelでダウンロード
                </button>
              )}
            </div>
            <ResultsTable
              results={searchResults}
              onCellClick={handleCellClick}
              onOpenExcel={handleOpenExcelFile}
            />
          </div>
        )}

        {selectedCell && (
          <CellDetails
            cellDetail={selectedCell}
            onClose={() => setSelectedCell(null)}
          />
        )}
      </main>
    </div>
  )
}

export default App
