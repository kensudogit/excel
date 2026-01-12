import React, { useState, useEffect } from 'react'
import './SearchReplace.css'

interface SearchReplaceResult {
  file_path: string
  file_name: string
  matches: Array<{
    line: number
    start: number
    end: number
    match_text: string
    line_content: string
    context_before: string
    context_after: string
  }>
  total_matches: number
  replaced: boolean
  backup_path?: string
  error?: string
}

interface SearchReplaceProps {
  onClose?: () => void
}

const SearchReplace: React.FC<SearchReplaceProps> = ({ onClose }) => {
  const [folderPath, setFolderPath] = useState('')
  const [searchPattern, setSearchPattern] = useState('')
  const [replacePattern, setReplacePattern] = useState('')
  const [useRegex, setUseRegex] = useState(false)
  const [fileExtensions, setFileExtensions] = useState<string[]>(['.txt', '.csv', '.html', '.js', '.ts', '.tsx', '.jsx', '.py', '.json', '.xml', '.css', '.xlsx', '.xls'])
  const [customExtension, setCustomExtension] = useState('')
  const [isLoading, setIsLoading] = useState(false)
  const [results, setResults] = useState<SearchReplaceResult[]>([])
  const [_previewMode, setPreviewMode] = useState(true)
  const [isDragging, setIsDragging] = useState(false)
  const fileInputRef = React.useRef<HTMLInputElement>(null)
  const [totalStats, setTotalStats] = useState<{
    total_files: number
    files_with_matches: number
    total_replacements: number
  } | null>(null)

  const commonExtensions = ['.txt', '.csv', '.html', '.js', '.ts', '.tsx', '.jsx', '.py', '.json', '.xml', '.css', '.md', '.yml', '.yaml', '.sql', '.sh', '.bat', '.ps1', '.xlsx', '.xls']

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

  const handleAddExtension = () => {
    if (customExtension && !fileExtensions.includes(customExtension)) {
      setFileExtensions([...fileExtensions, customExtension])
      setCustomExtension('')
    }
  }

  const handleRemoveExtension = (ext: string) => {
    setFileExtensions(fileExtensions.filter(e => e !== ext))
  }

  const handleToggleExtension = (ext: string) => {
    if (fileExtensions.includes(ext)) {
      handleRemoveExtension(ext)
    } else {
      setFileExtensions([...fileExtensions, ext])
    }
  }

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(true)
  }

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)
  }

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)

    const files = e.dataTransfer.files
    if (!files || files.length === 0) return

    // 最初のファイルを使用
    const file = files[0]
    const fileName = file.name

    // ファイルの絶対パスを取得を試みる
    // 方法1: Fileオブジェクトのpathプロパティ（Electron環境などで利用可能）
    const filePath = (file as any).path
    
    if (filePath) {
      // 絶対パスが取得できた場合
      const parentDir = filePath.substring(0, filePath.lastIndexOf('\\') || filePath.lastIndexOf('/'))
      setFolderPath(parentDir)
      return
    }

    // 方法2: webkitRelativePathを使用（相対パス）
    const relativePath = (file as any).webkitRelativePath
    if (relativePath) {
      // 相対パスから親ディレクトリを取得
      const parentDir = relativePath.substring(0, relativePath.lastIndexOf('/'))
      setFolderPath(parentDir)
      return
    }

    // 方法3: バックエンドのフォルダ選択ダイアログを使用して完全パスを取得
    // ブラウザのセキュリティ制限により、元のファイルパスは取得できません
    try {
      const response = await fetch('/api/browse-folder', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
      })

      const data = await response.json()

      if (data.success && data.folder_path) {
        // バックエンドから取得した完全パスを使用
        setFolderPath(data.folder_path)
      } else {
        // バックエンドのダイアログが利用できない場合は、ファイル名から推測
        console.warn(`ファイル "${fileName}" がドロップされました。このファイルが保存されているフォルダのパスを入力欄に入力してください。`)
      }
    } catch (error) {
      console.error('Browse folder error:', error)
      // すべての方法が失敗した場合
      console.warn(`ファイル "${fileName}" がドロップされました。このファイルが保存されているフォルダのパスを入力欄に入力してください。`)
    }
  }

  const handleBrowseFolder = () => {
    // HTML5のフォルダ選択ダイアログを開く
    if (fileInputRef.current) {
      fileInputRef.current.click()
    }
  }

  const handleFolderSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files
    if (!files || files.length === 0) return

    // 最初のファイルからフォルダパスを推測
    const firstFile = files[0]
    
    // webkitRelativePathからフォルダパスを取得
    const relativePath = (firstFile as any).webkitRelativePath
    if (relativePath) {
      // 相対パスからフォルダ名を取得
      const folderName = relativePath.split('/')[0]
      
      // フォルダ名を入力欄に設定
      setFolderPath(folderName)
      
      // フォルダ内のファイル数を確認
      const fileCount = files.length
      console.log(`フォルダ "${folderName}" が選択されました。${fileCount}個のファイルが見つかりました。`)
    } else {
      // webkitRelativePathが利用できない場合
      console.warn('フォルダが選択されましたが、パスを取得できませんでした。フォルダパスを手動で入力してください。')
    }
    
    // 入力要素をリセット（同じフォルダを再度選択できるように）
    if (fileInputRef.current) {
      fileInputRef.current.value = ''
    }
  }

  const handleSearch = async (executeReplace: boolean = false) => {
    if (!folderPath.trim() || !searchPattern.trim()) {
      console.warn('フォルダパスと検索パターンを入力してください')
      return
    }

    setIsLoading(true)
    setResults([])
    setTotalStats(null)

    try {
      const response = await fetch('/api/search-replace', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          folder_path: folderPath.trim(),
          search_pattern: searchPattern,
          replace_pattern: replacePattern,
          use_regex: useRegex,
          file_extensions: fileExtensions,
          preview_only: !executeReplace,
        }),
      })

      const data = await response.json()

      if (data.success) {
        setResults(data.results)
        setTotalStats({
          total_files: data.total_files,
          files_with_matches: data.files_with_matches,
          total_replacements: data.total_replacements,
        })
        setPreviewMode(!executeReplace)
        
        if (executeReplace) {
          console.log(`置換が完了しました。${data.total_replacements}箇所を置換しました。`)
        }
      } else {
        console.error('Search/Replace error:', data.error)
      }
    } catch (error) {
      console.error('Search/Replace error:', error)
    } finally {
      setIsLoading(false)
    }
  }

  return (
    <div className="search-replace-container">
      <div className="search-replace-header">
        <h2>🔍 一括検索・置換</h2>
        {onClose && (
          <button onClick={onClose} className="close-btn">✕</button>
        )}
      </div>

      <div className="search-replace-form">
        <div className="form-group">
          <label htmlFor="folderPath">対象フォルダ</label>
          <div
            className={`drop-zone ${isDragging ? 'dragging' : ''}`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            <div className="drop-zone-content">
              <div className="drop-zone-icon">📁</div>
              <div className="drop-zone-text">
                {isDragging ? 'ここにドロップしてください' : 'フォルダまたはファイルをここにドラッグ&ドロップ'}
              </div>
            </div>
          </div>
          <div className="folder-input-group" style={{ marginTop: '1rem' }}>
            <input
              type="text"
              id="folderPath"
              value={folderPath}
              onChange={(e) => setFolderPath(e.target.value)}
              placeholder="例: C:\Users\Documents\Project"
              className="form-input folder-input"
              disabled={isLoading}
            />
            <button
              type="button"
              onClick={handleBrowseFolder}
              className="browse-folder-btn"
              disabled={isLoading}
              title="フォルダを選択"
            >
              📁 フォルダ選択
            </button>
            <input
              ref={fileInputRef}
              type="file"
              {...({ webkitdirectory: '', directory: '' } as any)}
              multiple
              style={{ display: 'none' }}
              onChange={handleFolderSelect}
            />
          </div>
          <small className="form-hint">検索・置換を実行するフォルダのパスを入力するか、フォルダ/ファイルをドラッグ&ドロップしてください</small>
        </div>

        <div className="form-group">
          <label htmlFor="searchPattern">検索パターン</label>
          <textarea
            id="searchPattern"
            value={searchPattern}
            onChange={(e) => setSearchPattern(e.target.value)}
            placeholder="検索する文字列または正規表現"
            className="form-input"
            rows={2}
            disabled={isLoading}
          />
          <div className="checkbox-group">
            <label>
              <input
                type="checkbox"
                checked={useRegex}
                onChange={(e) => setUseRegex(e.target.checked)}
                disabled={isLoading}
              />
              正規表現を使用
            </label>
          </div>
        </div>

        <div className="form-group">
          <label htmlFor="replacePattern">置換パターン</label>
          <textarea
            id="replacePattern"
            value={replacePattern}
            onChange={(e) => setReplacePattern(e.target.value)}
            placeholder="置換後の文字列（正規表現使用時は$1, $2などが使用可能）"
            className="form-input"
            rows={2}
            disabled={isLoading}
          />
        </div>

        <div className="form-group">
          <label>対象ファイル拡張子</label>
          <div className="extension-selector">
            <div className="common-extensions">
              {commonExtensions.map(ext => (
                <label key={ext} className="extension-checkbox">
                  <input
                    type="checkbox"
                    checked={fileExtensions.includes(ext)}
                    onChange={() => handleToggleExtension(ext)}
                    disabled={isLoading}
                  />
                  {ext}
                </label>
              ))}
            </div>
            <div className="custom-extension">
              <input
                type="text"
                value={customExtension}
                onChange={(e) => setCustomExtension(e.target.value)}
                placeholder="カスタム拡張子（例: .log）"
                className="form-input"
                style={{ width: '200px', marginRight: '0.5rem' }}
                disabled={isLoading}
                onKeyPress={(e) => {
                  if (e.key === 'Enter') {
                    handleAddExtension()
                  }
                }}
              />
              <button
                onClick={handleAddExtension}
                className="add-extension-btn"
                disabled={isLoading}
              >
                + 追加
              </button>
            </div>
            <div className="selected-extensions">
              <strong>選択中:</strong> {fileExtensions.join(', ')}
            </div>
          </div>
        </div>

        <div className="button-group">
          <button
            onClick={() => handleSearch(false)}
            className="search-btn"
            disabled={isLoading}
          >
            {isLoading ? '検索中...' : '🔍 プレビュー（検索のみ）'}
          </button>
          <button
            onClick={() => {
              if (window.confirm('本当に置換を実行しますか？バックアップファイル（.bak）が作成されます。')) {
                handleSearch(true)
              }
            }}
            className="replace-btn"
            disabled={isLoading || !replacePattern.trim()}
          >
            {isLoading ? '置換中...' : '🔄 置換実行'}
          </button>
        </div>
      </div>

      {totalStats && (
        <div className="stats-section">
          <h3>検索結果サマリー</h3>
          <div className="stats-grid">
            <div className="stat-item">
              <span className="stat-label">対象ファイル数:</span>
              <span className="stat-value">{totalStats.total_files}</span>
            </div>
            <div className="stat-item">
              <span className="stat-label">マッチしたファイル数:</span>
              <span className="stat-value">{totalStats.files_with_matches}</span>
            </div>
            <div className="stat-item">
              <span className="stat-label">総置換数:</span>
              <span className="stat-value">{totalStats.total_replacements}</span>
            </div>
          </div>
        </div>
      )}

      {results.length > 0 && (
        <div className="results-section">
          <h3>検索結果詳細</h3>
          <div className="results-list">
            {results.map((result, index) => (
              <div key={index} className="result-item">
                <div className="result-header">
                  <span className="file-name">{result.file_name}</span>
                  <span className="match-count">{result.total_matches}件</span>
                  {result.replaced && (
                    <span className="replaced-badge">✓ 置換済み</span>
                  )}
                  {result.error && (
                    <span className="error-badge">✗ エラー</span>
                  )}
                </div>
                {result.error ? (
                  <div className="error-message">{result.error}</div>
                ) : (
                  <div className="matches-list">
                    {result.matches.slice(0, 10).map((match, matchIndex) => (
                      <div key={matchIndex} className="match-item">
                        <div className="match-line">
                          <span className="line-number">行 {match.line}:</span>
                          <span className="line-content">{match.line_content}</span>
                        </div>
                        <div className="match-details">
                          <span className="match-text">「{match.match_text}」</span>
                        </div>
                      </div>
                    ))}
                    {result.matches.length > 10 && (
                      <div className="more-matches">
                        ... 他 {result.matches.length - 10} 件
                      </div>
                    )}
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  )
}

export default SearchReplace
