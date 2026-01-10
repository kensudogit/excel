import React, { useState, useMemo, useEffect } from 'react'
import './ResultsTable.css'
import { SearchResult } from '../types'

interface ResultsTableProps {
  results: SearchResult[]
  onCellClick: (result: SearchResult) => void
  onOpenExcel: (result: SearchResult) => void
}

const ITEMS_PER_PAGE = 20

const ResultsTable: React.FC<ResultsTableProps> = ({ results, onCellClick, onOpenExcel }) => {
  const [currentPage, setCurrentPage] = useState(1)

  // 検索結果が変更されたときにページを1にリセット
  useEffect(() => {
    setCurrentPage(1)
  }, [results.length])

  const getKeywordColor = (keyword: string) => {
    const colors: { [key: string]: string } = {
      // キーワードごとに異なる色を割り当て
    }
    return colors[keyword] || '#667eea'
  }

  // ページネーション計算
  const totalPages = Math.ceil(results.length / ITEMS_PER_PAGE)
  const startIndex = (currentPage - 1) * ITEMS_PER_PAGE
  const endIndex = startIndex + ITEMS_PER_PAGE
  const currentResults = useMemo(() => results.slice(startIndex, endIndex), [results, startIndex, endIndex])

  // ページ変更時にトップにスクロール
  const handlePageChange = (page: number) => {
    setCurrentPage(page)
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }

  // ページ番号の配列を生成
  const getPageNumbers = () => {
    const pages: (number | string)[] = []
    const maxVisiblePages = 7

    if (totalPages <= maxVisiblePages) {
      // ページ数が少ない場合は全て表示
      for (let i = 1; i <= totalPages; i++) {
        pages.push(i)
      }
    } else {
      // ページ数が多い場合は省略表示
      if (currentPage <= 3) {
        // 最初の数ページ
        for (let i = 1; i <= 4; i++) {
          pages.push(i)
        }
        pages.push('...')
        pages.push(totalPages)
      } else if (currentPage >= totalPages - 2) {
        // 最後の数ページ
        pages.push(1)
        pages.push('...')
        for (let i = totalPages - 3; i <= totalPages; i++) {
          pages.push(i)
        }
      } else {
        // 中間のページ
        pages.push(1)
        pages.push('...')
        for (let i = currentPage - 1; i <= currentPage + 1; i++) {
          pages.push(i)
        }
        pages.push('...')
        pages.push(totalPages)
      }
    }

    return pages
  }

  return (
    <div className="results-table-container">
      <div className="results-table-header">
        <div className="results-count">
          検索結果: {results.length}件 (ページ {currentPage} / {totalPages})
        </div>
      </div>
      <table className="results-table">
        <thead>
          <tr>
            <th>ファイル名</th>
            <th>シート名</th>
            <th>行</th>
            <th>列</th>
            <th>セル値</th>
            <th>キーワード</th>
            <th>操作</th>
          </tr>
        </thead>
        <tbody>
          {currentResults.map((result, index) => (
            <tr key={startIndex + index} className="result-row">
              <td>{result.file.split(/[/\\]/).pop()}</td>
              <td>{result.sheet}</td>
              <td className="text-center">{result.row}</td>
              <td className="text-center">{result.col}</td>
              <td className="cell-value">{result.value}</td>
              <td>
                <span
                  className="keyword-badge clickable-keyword"
                  style={{ backgroundColor: getKeywordColor(result.keyword) }}
                  onClick={(e) => {
                    e.stopPropagation()
                    onOpenExcel(result)
                  }}
                  title="クリックしてExcelファイルを開く"
                >
                  {result.keyword}
                </span>
              </td>
              <td>
                <div style={{ display: 'flex', gap: '0.5rem' }}>
                  <button
                    onClick={() => onOpenExcel(result)}
                    className="open-excel-btn"
                    title="Excelファイルを開く"
                  >
                    📂 開く
                  </button>
                  <button
                    onClick={() => onCellClick(result)}
                    className="view-details-btn"
                    title="セル詳細を表示"
                  >
                    📋 詳細
                  </button>
                </div>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {totalPages > 1 && (
        <div className="pagination">
          <button
            className="pagination-btn"
            onClick={() => handlePageChange(1)}
            disabled={currentPage === 1}
            title="最初のページ"
          >
            ««
          </button>
          <button
            className="pagination-btn"
            onClick={() => handlePageChange(currentPage - 1)}
            disabled={currentPage === 1}
            title="前のページ"
          >
            «
          </button>
          
          {getPageNumbers().map((page, index) => {
            if (page === '...') {
              return (
                <span key={`ellipsis-${index}`} className="pagination-ellipsis">
                  ...
                </span>
              )
            }
            
            return (
              <button
                key={page}
                className={`pagination-btn ${currentPage === page ? 'active' : ''}`}
                onClick={() => handlePageChange(page as number)}
                title={`ページ ${page}`}
              >
                {page}
              </button>
            )
          })}
          
          <button
            className="pagination-btn"
            onClick={() => handlePageChange(currentPage + 1)}
            disabled={currentPage === totalPages}
            title="次のページ"
          >
            »
          </button>
          <button
            className="pagination-btn"
            onClick={() => handlePageChange(totalPages)}
            disabled={currentPage === totalPages}
            title="最後のページ"
          >
            »»
          </button>
        </div>
      )}
    </div>
  )
}

export default ResultsTable
