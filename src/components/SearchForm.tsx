import React, { useState, useRef } from 'react'
import './SearchForm.css'

interface SearchFormProps {
  onSearch: (folderPath: string, keywords: string[]) => void
  onSearchWithFiles?: (files: File[], keywords: string[]) => void
  isLoading: boolean
}

interface FolderFile {
  name: string
  path: string
  size: number
  type: string
}

const SearchForm: React.FC<SearchFormProps> = ({ onSearch, onSearchWithFiles, isLoading }) => {
  const [folderPath, setFolderPath] = useState('')
  const [keywords, setKeywords] = useState<string[]>([''])
  const [error, setError] = useState('')
  const [isDragging, setIsDragging] = useState(false)
  const [folderFiles, setFolderFiles] = useState<FolderFile[]>([])
  const [showFolderContents, setShowFolderContents] = useState(false)
  const [selectedExcelFiles, setSelectedExcelFiles] = useState<File[]>([])
  const fileInputRef = useRef<HTMLInputElement>(null)

  const handleBrowseFolder = async () => {
    // バックエンドのフォルダ選択ダイアログを使用
    try {
      const response = await fetch('/api/browse-folder', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
      })

      // レスポンスが空でないか確認
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

      if (data.success && data.folder_path) {
        setFolderPath(data.folder_path)
        setFolderFiles([])
        setSelectedExcelFiles([])
        setShowFolderContents(false)
        
        // フォルダ内のExcelファイルを検索
        // 注意: ブラウザからは直接ファイルシステムにアクセスできないため、
        // バックエンドでフォルダ内のファイルを取得する必要がある
      } else {
        // バックエンドのダイアログが利用できない場合は、HTML5のフォルダ選択を使用
        if (fileInputRef.current) {
          fileInputRef.current.click()
        }
      }
    } catch (error) {
      console.error('Browse folder error:', error)
      // エラーが発生した場合は、HTML5のフォルダ選択を使用
      if (fileInputRef.current) {
        fileInputRef.current.click()
      }
    }
  }

  const handleFolderSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files
    if (!files || files.length === 0) return

    // 最初のファイルからフォルダパスを推測
    const firstFile = files[0]
    
    // webkitRelativePathからフォルダパスを取得
    const relativePath = (firstFile as any).webkitRelativePath
    if (relativePath) {
      // 相対パスからフォルダ名を取得
      const folderName = relativePath.split('/')[0]
      
      // ブラウザのセキュリティ制限により、HTML5フォルダ選択では完全パスを取得できません
      // バックエンドのフォルダ選択ダイアログを使用して完全パスを取得することを推奨
      // バックエンドのフォルダ選択ダイアログを自動的に開く
      try {
        const response = await fetch('/api/browse-folder', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
        })

        // レスポンスが空でないか確認
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

        if (data.success && data.folder_path) {
          // バックエンドから取得した完全パスを使用
          setFolderPath(data.folder_path)
          setFolderFiles([])
          setSelectedExcelFiles([])
          setShowFolderContents(false)
          return
        }
      } catch (error) {
        console.error('Browse folder error:', error)
      }
      
      // フォルダ名のみを設定（完全パスは手動入力が必要）
      setFolderPath(folderName)
      
      // フォルダ内の全ファイル情報を取得
      const fileList: FolderFile[] = Array.from(files).map(file => ({
        name: file.name,
        path: (file as any).webkitRelativePath || file.name,
        size: file.size,
        type: file.type || (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') ? 'Excel' : 'その他')
      }))
      
      setFolderFiles(fileList)
      setShowFolderContents(true)
      
      // Excelファイルを抽出して保持
      const excelFiles = Array.from(files).filter(file => 
        file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
      )
      setSelectedExcelFiles(excelFiles)
      
      if (excelFiles.length === 0) {
        console.warn(`フォルダ "${folderName}" が選択されました。Excelファイルが見つかりませんでした。`)
      }
    } else {
      // webkitRelativePathが利用できない場合
      console.warn('フォルダが選択されましたが、パスを取得できませんでした。フォルダパスを手動で入力してください。')
      setFolderFiles([])
      setSelectedExcelFiles([])
      setShowFolderContents(false)
    }
    
    // 入力要素をリセット（同じフォルダを再度選択できるように）
    // 注意: リセットするとファイルが失われるため、リセットしない
    // if (fileInputRef.current) {
    //   fileInputRef.current.value = ''
    // }
  }

  const handleClearFolderContents = () => {
    setFolderFiles([])
    setSelectedExcelFiles([])
    setShowFolderContents(false)
    setFolderPath('')
    if (fileInputRef.current) {
      fileInputRef.current.value = ''
    }
  }
  
  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 B'
    const k = 1024
    const sizes = ['B', 'KB', 'MB', 'GB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i]
  }

  const handleAddKeyword = () => {
    setKeywords([...keywords, ''])
  }

  const handleRemoveKeyword = (index: number) => {
    if (keywords.length > 1) {
      setKeywords(keywords.filter((_, i) => i !== index))
    }
  }

  const handleKeywordChange = (index: number, value: string) => {
    const newKeywords = [...keywords]
    newKeywords[index] = value
    setKeywords(newKeywords)
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

    const items = e.dataTransfer.items
    const files = e.dataTransfer.files
    
    if (!items && (!files || files.length === 0)) return

    // フォルダかファイルかを判定
    let isFolder = false
    let folderName = ''
    
    if (items && items.length > 0) {
      // DataTransferItemListを使用してフォルダかどうかを判定
      const item = items[0]
      if (item.webkitGetAsEntry) {
        const entry = item.webkitGetAsEntry()
        if (entry && entry.isDirectory) {
          isFolder = true
          folderName = entry.name
        }
      }
    }

    // フォルダがドロップされた場合
    if (isFolder && files && files.length > 0) {
      // フォルダ内のファイルからフォルダ名を取得
      const firstFile = files[0]
      const relativePath = (firstFile as any).webkitRelativePath
      if (relativePath) {
        folderName = relativePath.split('/')[0]
      }
      
      // バックエンドのフォルダ選択ダイアログを開いて完全パスを取得
      try {
        const response = await fetch('/api/browse-folder', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
        })

        // レスポンスが空でないか確認
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

        if (data.success && data.folder_path) {
          // バックエンドから取得した完全パスを使用
          setFolderPath(data.folder_path)
          
          // フォルダ内のExcelファイルを抽出
          const excelFiles = Array.from(files).filter(file => 
            file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
          )
          
          if (excelFiles.length > 0) {
            setSelectedExcelFiles(excelFiles)
            
            // フォルダ内の全ファイル情報を取得
            const fileList: FolderFile[] = Array.from(files).map(file => ({
              name: file.name,
              path: (file as any).webkitRelativePath || file.name,
              size: file.size,
              type: file.type || (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') ? 'Excel' : 'その他')
            }))
            
            setFolderFiles(fileList)
            setShowFolderContents(true)
          }
        } else {
          console.warn(`フォルダ "${folderName}" がドロップされました。バックエンドのフォルダ選択ダイアログが利用できません。`)
        }
      } catch (error) {
        console.error('Browse folder error:', error)
        console.warn(`フォルダ "${folderName}" がドロップされました。フォルダパスを手動で入力してください。`)
      }
      return
    }

    // ファイルがドロップされた場合
    if (files && files.length > 0) {
      const file = files[0]
      const fileName = file.name

      // Excelファイルかチェック
      if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
        console.warn('Excelファイル（.xlsx または .xls）をドロップしてください')
        return
      }

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
      try {
        const response = await fetch('/api/browse-folder', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
        })

        // レスポンスが空でないか確認
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

        if (data.success && data.folder_path) {
          // バックエンドから取得した完全パスを使用
          setFolderPath(data.folder_path)
        } else {
          console.warn(`ファイル "${fileName}" がドロップされました。このファイルが保存されているフォルダのパスを入力欄に入力してください。`)
        }
      } catch (error) {
        console.error('Browse folder error:', error)
        console.warn(`ファイル "${fileName}" がドロップされました。このファイルが保存されているフォルダのパスを入力欄に入力してください。`)
      }
    }
  }

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault()
    setError('')

    const validKeywords = keywords.filter(k => k.trim() !== '')
    if (validKeywords.length === 0) {
      setError('少なくとも1つのキーワードを入力してください')
      return
    }

    // フォルダ選択で取得したExcelファイルがある場合は、それを使用
    if (selectedExcelFiles.length > 0 && onSearchWithFiles) {
      onSearchWithFiles(selectedExcelFiles, validKeywords)
      return
    }

    // フォルダパスが入力されている場合は、通常の検索を実行
    if (!folderPath.trim()) {
      setError('フォルダパスを入力するか、フォルダを選択してください')
      return
    }

    onSearch(folderPath.trim(), validKeywords)
  }

  return (
    <div className="search-form-container">
      <form onSubmit={handleSubmit} className="search-form">
        <div className="form-group">
          <label htmlFor="folderPath">検索対象フォルダ</label>
          <div
            className={`drop-zone ${isDragging ? 'dragging' : ''}`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            <div className="drop-zone-content">
              <div className="drop-zone-icon">📁</div>
              <div className="drop-zone-text">
                {isDragging ? 'ここにドロップしてください' : 'フォルダまたはExcelファイルをここにドラッグ&ドロップ'}
              </div>
            </div>
          </div>
          <div className="folder-input-group" style={{ marginTop: '1rem' }}>
            <input
              type="text"
              id="folderPath"
              value={folderPath}
              onChange={(e) => setFolderPath(e.target.value)}
              placeholder="例: C:\Users\Documents\ExcelFiles"
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
              accept=".xlsx,.xls"
            />
          </div>
          <small className="form-hint">検索したいExcelファイルが保存されているフォルダのパスを入力するか、フォルダ/Excelファイルをドラッグ&ドロップしてください</small>
          
          {showFolderContents && folderFiles.length > 0 && (
            <div className="folder-contents" style={{ marginTop: '1rem' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '0.5rem' }}>
                <h3 style={{ margin: 0, fontSize: '1rem', fontWeight: 600 }}>フォルダ内容 ({folderFiles.length}個のファイル)</h3>
                <div style={{ display: 'flex', gap: '0.5rem', alignItems: 'center' }}>
                  <button
                    type="button"
                    onClick={handleClearFolderContents}
                    style={{
                      background: '#ef4444',
                      color: 'white',
                      border: 'none',
                      borderRadius: '4px',
                      padding: '0.4rem 0.8rem',
                      cursor: 'pointer',
                      fontSize: '0.85rem',
                      fontWeight: 600,
                      transition: 'background-color 0.3s'
                    }}
                    onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#dc2626'}
                    onMouseOut={(e) => e.currentTarget.style.backgroundColor = '#ef4444'}
                    title="フォルダ内容をクリア"
                    disabled={isLoading}
                  >
                    🗑️ クリア
                  </button>
                  <button
                    type="button"
                    onClick={() => setShowFolderContents(false)}
                    style={{
                      background: 'none',
                      border: 'none',
                      color: '#666',
                      cursor: 'pointer',
                      fontSize: '1.2rem',
                      padding: '0.25rem 0.5rem'
                    }}
                    title="閉じる"
                  >
                    ✕
                  </button>
                </div>
              </div>
              <div className="folder-files-list" style={{
                maxHeight: '300px',
                overflowY: 'auto',
                border: '1px solid #e0e0e0',
                borderRadius: '6px',
                padding: '0.5rem',
                backgroundColor: '#f9fafb'
              }}>
                {folderFiles.map((file, index) => {
                  const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
                  return (
                    <div
                      key={index}
                      style={{
                        padding: '0.5rem',
                        marginBottom: '0.25rem',
                        backgroundColor: isExcel ? '#e0f2fe' : '#f3f4f6',
                        borderRadius: '4px',
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                        fontSize: '0.9rem'
                      }}
                    >
                      <div style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        <span style={{ fontWeight: isExcel ? 600 : 400, color: isExcel ? '#0369a1' : '#374151' }}>
                          {isExcel ? '📊 ' : '📄 '}
                          {file.name}
                        </span>
                        <span style={{ color: '#6b7280', marginLeft: '0.5rem', fontSize: '0.85rem' }}>
                          ({formatFileSize(file.size)})
                        </span>
                      </div>
                    </div>
                  )
                })}
              </div>
              <div style={{ marginTop: '0.5rem', fontSize: '0.85rem', color: '#666' }}>
                Excelファイル: {folderFiles.filter(f => f.name.endsWith('.xlsx') || f.name.endsWith('.xls')).length}個
              </div>
            </div>
          )}
        </div>

        <div className="form-group">
          <label>検索キーワード</label>
          {keywords.map((keyword, index) => (
            <div key={index} className="keyword-input-group">
              <input
                type="text"
                value={keyword}
                onChange={(e) => handleKeywordChange(index, e.target.value)}
                placeholder={`キーワード ${index + 1}`}
                className="form-input keyword-input"
                disabled={isLoading}
              />
              {keywords.length > 1 && (
                <button
                  type="button"
                  onClick={() => handleRemoveKeyword(index)}
                  className="remove-keyword-btn"
                  disabled={isLoading}
                >
                  ✕
                </button>
              )}
            </div>
          ))}
          <button
            type="button"
            onClick={handleAddKeyword}
            className="add-keyword-btn"
            disabled={isLoading}
          >
            + キーワードを追加
          </button>
        </div>

        {error && <div className="error-message">{error}</div>}

        <button
          type="submit"
          className="search-btn"
          disabled={isLoading}
        >
          {isLoading ? '検索中...' : '🔍 検索開始'}
        </button>
      </form>
    </div>
  )
}

export default SearchForm
