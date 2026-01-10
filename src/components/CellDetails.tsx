import React from 'react'
import './CellDetails.css'
import { CellDetail } from '../types'

interface CellDetailsProps {
  cellDetail: CellDetail
  onClose: () => void
}

const CellDetails: React.FC<CellDetailsProps> = ({ cellDetail, onClose }) => {
  const getColumnLetter = (col: number): string => {
    let result = ''
    while (col > 0) {
      col--
      result = String.fromCharCode(65 + (col % 26)) + result
      col = Math.floor(col / 26)
    }
    return result
  }

  return (
    <div className="cell-details-overlay" onClick={onClose}>
      <div className="cell-details-modal" onClick={(e) => e.stopPropagation()}>
        <div className="cell-details-header">
          <h2>セル詳細情報</h2>
          <button onClick={onClose} className="close-btn">✕</button>
        </div>

        <div className="cell-details-content">
          <div className="cell-info-section">
            <h3>ファイル情報</h3>
            <div className="info-grid">
              <div className="info-item">
                <span className="info-label">ファイル名:</span>
                <span className="info-value">{cellDetail.file_name}</span>
              </div>
              <div className="info-item">
                <span className="info-label">シート名:</span>
                <span className="info-value">{cellDetail.sheet_name}</span>
              </div>
              <div className="info-item">
                <span className="info-label">セル位置:</span>
                <span className="info-value">
                  {getColumnLetter(cellDetail.target_cell.col)}
                  {cellDetail.target_cell.row}
                </span>
              </div>
              <div className="info-item">
                <span className="info-label">キーワード:</span>
                <span className="info-value keyword-highlight">
                  {cellDetail.target_cell.keyword}
                </span>
              </div>
              <div className="info-item">
                <span className="info-label">セル値:</span>
                <span className="info-value">{cellDetail.target_cell.value}</span>
              </div>
            </div>
          </div>

          <div className="context-section">
            <h3>周辺セル（前後5行）</h3>
            <div className="context-table-container">
              <table className="context-table">
                <thead>
                  <tr>
                    <th>行</th>
                    <th>列</th>
                    <th>セル位置</th>
                    <th>値</th>
                  </tr>
                </thead>
                <tbody>
                  {cellDetail.context.flatMap((row, rowIndex) =>
                    row.map((cell, colIndex) => (
                      <tr
                        key={`${rowIndex}-${colIndex}`}
                        className={
                          cell.is_target
                            ? 'target-cell'
                            : cell.is_header
                            ? 'header-cell'
                            : ''
                        }
                      >
                        <td className="text-center">{cell.row}</td>
                        <td className="text-center">{cell.col}</td>
                        <td className="text-center">
                          {getColumnLetter(cell.col)}
                          {cell.row}
                        </td>
                        <td className={cell.is_target ? 'target-value' : ''}>
                          {cell.value || '(空)'}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}

export default CellDetails
