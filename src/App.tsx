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
        try {
          const errorData = await response.json()
          errorMessage = errorData.error || errorMessage
        } catch (e) {
          // JSON解析に失敗した場合は、ステータスコードのみを使用
          errorMessage = `HTTP error! status: ${response.status}`
        }
        throw new Error(errorMessage)
      }

      const data = await response.json()

      if (data.success) {
        setSearchResults(data.results || [])
        setOutputFile(data.output_file || '')
      } else {
        console.error('Search error:', data.error || '検索に失敗しました')
      }
    } catch (error) {
      console.error('Search error:', error)
      if (error instanceof SyntaxError) {
        console.error('サーバーからの応答を解析できませんでした。サーバーが正常に動作しているか確認してください。')
      } else {
        console.error(`検索中にエラーが発生しました: ${error instanceof Error ? error.message : '不明なエラー'}`)
      }
    } finally {
      setIsLoading(false)
    }
  }

  const handleSearchWithFiles = async (files: File[], keywords: string[]) => {
    setIsLoading(true)
    setSearchResults([])
    setSelectedCell(null)
    setOutputFile('')

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
          const errorData = await response.json()
          errorMessage = errorData.error || errorMessage
        } catch (e) {
          // JSON解析に失敗した場合は、テキストとして読み込む
          try {
            const errorText = await responseClone.text()
            errorMessage = errorText || errorMessage
          } catch (textError) {
            // テキスト読み込みも失敗した場合は、ステータスコードのみを使用
            errorMessage = `HTTP error! status: ${response.status}`
          }
        }
        throw new Error(errorMessage)
      }

      const data = await response.json()

      if (data.success) {
        setSearchResults(data.results || [])
        setOutputFile(data.output_file || '')
      } else {
        console.error('Search error:', data.error || '検索に失敗しました')
      }
    } catch (error) {
      console.error('Search with files error:', error)
      if (error instanceof TypeError && error.message.includes('Failed to fetch')) {
        console.error('サーバーに接続できませんでした。バックエンドサーバー（ポート5001）が起動しているか確認してください。')
      } else if (error instanceof SyntaxError) {
        console.error('サーバーからの応答を解析できませんでした。サーバーが正常に動作しているか確認してください。')
      } else {
        console.error(`検索中にエラーが発生しました: ${error instanceof Error ? error.message : '不明なエラー'}`)
      }
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

      const data = await response.json()

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

      const data = await response.json()

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
