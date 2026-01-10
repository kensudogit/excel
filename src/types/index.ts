export interface SearchResult {
  file: string
  sheet: string
  row: number
  col: number
  value: string
  keyword: string
}

export interface CellDetail {
  success: boolean
  file_name: string
  sheet_name: string
  target_cell: {
    row: number
    col: number
    value: string
    keyword: string
  }
  context: Array<Array<{
    row: number
    col: number
    value: string
    is_target: boolean
    is_header: boolean
  }>>
  max_row: number
  max_col: number
}
