"""
Excelファイル検索・抽出アプリケーション
指定したフォルダ内のExcelファイルから複数のキーワードを検索し、結果を別ブックに出力
"""
import os
import json
import re
import shutil
import subprocess
import platform
from pathlib import Path
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import pandas as pd
from datetime import datetime

# Windows環境でExcelを操作するためのライブラリ（オプション）
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

app = Flask(__name__)
CORS(app)

# ファイルアップロードサイズ制限を設定（デフォルトは16MB、100MBに拡大）
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB

# 設定
UPLOAD_FOLDER = Path('uploads')
RESULTS_FOLDER = Path('results')
UPLOAD_FOLDER.mkdir(exist_ok=True)
RESULTS_FOLDER.mkdir(exist_ok=True)


def search_keywords_in_excel(file_path, keywords):
    """
    Excelファイル内でキーワードを検索
    戻り値: [(sheet_name, row, col, cell_value, keyword), ...]
    """
    results = []
    try:
        # file_pathがPathオブジェクトの場合は文字列に変換
        file_path_str = str(file_path) if isinstance(file_path, Path) else file_path
        
        wb = openpyxl.load_workbook(file_path_str, data_only=True)
        
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            
            for row_idx, row in enumerate(sheet.iter_rows(values_only=False), start=1):
                for col_idx, cell in enumerate(row, start=1):
                    if cell.value is None:
                        continue
                    
                    cell_value = str(cell.value)
                    
                    # 各キーワードをチェック
                    for keyword in keywords:
                        if keyword.lower() in cell_value.lower():
                            results.append({
                                'sheet': sheet_name,
                                'row': row_idx,
                                'col': col_idx,
                                'value': cell_value,
                                'keyword': keyword,
                                'file': file_path_str  # パスをそのまま使用（後で上書きされる可能性がある）
                            })
        
        wb.close()
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"Error processing {file_path}: {error_trace}")
        app.logger.error(f"Error processing {file_path}: {error_trace}")
    
    return results


def create_results_workbook(search_results, keywords):
    """
    検索結果をExcelブックに出力
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "検索結果"
    
    # ヘッダー行
    headers = ['ファイル名', 'シート名', '行', '列', 'セル値', 'キーワード', 'ファイルパス']
    ws.append(headers)
    
    # ヘッダーのスタイル設定
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # データ行
    for result in search_results:
        row = [
            Path(result['file']).name,
            result['sheet'],
            result['row'],
            result['col'],
            result['value'],
            result['keyword'],
            result['file']
        ]
        ws.append(row)
        
        # キーワードに応じて行の色を変更
        keyword_colors = {
            keywords[0]: "FFE6E6" if len(keywords) > 0 else "FFFFFF",
            keywords[1]: "E6F3FF" if len(keywords) > 1 else "FFFFFF",
            keywords[2]: "E6FFE6" if len(keywords) > 2 else "FFFFFF",
        }
        fill_color = keyword_colors.get(result['keyword'], "FFFFFF")
        fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        
        for col in range(1, len(row) + 1):
            ws.cell(row=ws.max_row, column=col).fill = fill
    
    # 列幅の自動調整
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[col_letter].width = adjusted_width
    
    return wb


@app.route('/api/search', methods=['POST'])
def search_excel_files():
    """
    指定フォルダ内のExcelファイルを検索
    """
    try:
        # リクエストデータの取得
        if not request.is_json:
            return jsonify({'success': False, 'error': 'リクエストはJSON形式である必要があります'}), 400
        
        data = request.get_json()
        if data is None:
            return jsonify({'success': False, 'error': 'リクエストデータが空です'}), 400
        
        folder_path = data.get('folder_path', '')
        keywords = data.get('keywords', [])
        
        if not folder_path:
            return jsonify({'success': False, 'error': 'フォルダパスが指定されていません'}), 400
        
        if not keywords or len(keywords) == 0:
            return jsonify({'success': False, 'error': 'キーワードが指定されていません'}), 400
        
        folder = Path(folder_path)
        if not folder.exists():
            return jsonify({'success': False, 'error': f'指定されたフォルダが見つかりません: {folder_path}'}), 404
        
        if not folder.is_dir():
            return jsonify({'success': False, 'error': f'指定されたパスはフォルダではありません: {folder_path}'}), 400
        
        # Excelファイルを検索
        excel_files = list(folder.glob('*.xlsx')) + list(folder.glob('*.xls'))
        
        if not excel_files:
            return jsonify({'success': False, 'error': 'Excelファイルが見つかりませんでした'}), 404
        
        # 各ファイルを検索
        all_results = []
        for excel_file in excel_files:
            try:
                results = search_keywords_in_excel(excel_file, keywords)
                all_results.extend(results)
            except Exception as e:
                print(f"Error processing {excel_file}: {str(e)}")
                continue
        
        # 結果をExcelブックに出力
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = RESULTS_FOLDER / f'search_results_{timestamp}.xlsx'
            wb = create_results_workbook(all_results, keywords)
            wb.save(output_file)
        except Exception as e:
            print(f"Error creating workbook: {str(e)}")
            # ブック作成に失敗しても検索結果は返す
        
        # 結果をJSON形式で返す
        output_file_str = str(output_file) if 'output_file' in locals() else None
        return jsonify({
            'success': True,
            'results': all_results,
            'total_matches': len(all_results),
            'files_searched': len(excel_files),
            'output_file': output_file_str
        })
        
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"Error in search_excel_files: {error_trace}")
        app.logger.error(f"Error in search_excel_files: {error_trace}")
        return jsonify({
            'success': False,
            'error': f'検索中にエラーが発生しました: {str(e)}'
        }), 500


@app.route('/api/search-files', methods=['POST'])
def search_excel_files_upload():
    """
    アップロードされたExcelファイルを検索
    """
    try:
        app.logger.info(f"Received request to /api/search-files")
        app.logger.info(f"Request method: {request.method}")
        app.logger.info(f"Request content type: {request.content_type}")
        app.logger.info(f"Request form keys: {list(request.form.keys())}")
        app.logger.info(f"Request files keys: {list(request.files.keys())}")
        # キーワードの取得
        keywords_json = request.form.get('keywords', '[]')
        try:
            keywords = json.loads(keywords_json)
        except json.JSONDecodeError:
            return jsonify({'success': False, 'error': 'キーワードの形式が正しくありません'}), 400
        
        if not keywords or len(keywords) == 0:
            return jsonify({'success': False, 'error': 'キーワードが指定されていません'}), 400
        
        # アップロードされたファイルの取得
        if 'files' not in request.files:
            return jsonify({'success': False, 'error': 'ファイルが指定されていません'}), 400
        
        uploaded_files = request.files.getlist('files')
        if not uploaded_files or len(uploaded_files) == 0:
            return jsonify({'success': False, 'error': 'ファイルが指定されていません'}), 400
        
        # Excelファイルのみをフィルタリング
        excel_files = []
        for file in uploaded_files:
            if file.filename == '':
                continue
            if file.filename.endswith('.xlsx') or file.filename.endswith('.xls'):
                excel_files.append(file)
        
        if not excel_files:
            return jsonify({'success': False, 'error': 'Excelファイルが見つかりませんでした'}), 404
        
        # 各ファイルを一時保存して検索
        all_results = []
        temp_file_paths = []
        
        for excel_file in excel_files:
            try:
                # 一時ファイルとして保存（ファイル名の重複を防ぐため、タイムスタンプを追加）
                import time
                timestamp = int(time.time() * 1000)
                safe_filename = f"{timestamp}_{excel_file.filename}"
                temp_file = UPLOAD_FOLDER / safe_filename
                excel_file.save(str(temp_file))
                temp_file_paths.append(temp_file)
                
                # 検索実行
                results = search_keywords_in_excel(temp_file, keywords)
                # ファイル名を元のファイル名に設定（パスではなくファイル名のみ）
                original_filename = excel_file.filename
                for result in results:
                    result['file'] = original_filename
                all_results.extend(results)
            except Exception as e:
                import traceback
                error_trace = traceback.format_exc()
                print(f"Error processing {excel_file.filename}: {error_trace}")
                app.logger.error(f"Error processing {excel_file.filename}: {error_trace}")
                continue
        
        # 結果をExcelブックに出力
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_file = RESULTS_FOLDER / f'search_results_{timestamp}.xlsx'
            wb = create_results_workbook(all_results, keywords)
            wb.save(output_file)
        except Exception as e:
            print(f"Error creating workbook: {str(e)}")
        
        # 一時ファイルを削除
        for temp_file in temp_file_paths:
            try:
                if temp_file.exists():
                    temp_file.unlink()
            except Exception as e:
                print(f"Error deleting temp file {temp_file}: {str(e)}")
        
        # 結果をJSON形式で返す
        output_file_str = str(output_file) if 'output_file' in locals() else None
        app.logger.info(f"Search completed: {len(all_results)} matches found in {len(excel_files)} files")
        return jsonify({
            'success': True,
            'results': all_results,
            'total_matches': len(all_results),
            'files_searched': len(excel_files),
            'output_file': output_file_str
        })
        
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"Error in search_excel_files_upload: {error_trace}")
        app.logger.error(f"Error in search_excel_files_upload: {error_trace}")
        return jsonify({
            'success': False,
            'error': f'検索中にエラーが発生しました: {str(e)}',
            'traceback': error_trace if app.debug else None
        }), 500


@app.route('/api/get-cell-details', methods=['POST'])
def get_cell_details():
    """
    特定のセルの詳細情報を取得（周辺のセルも含む）
    """
    try:
        data = request.json
        file_path = data.get('file_path', '')
        sheet_name = data.get('sheet_name', '')
        row = data.get('row', 0)
        col = data.get('col', 0)
        keyword = data.get('keyword', '')
        context_rows = data.get('context_rows', 5)  # 前後何行表示するか
        
        if not file_path or not sheet_name or not row or not col:
            return jsonify({'success': False, 'error': '必要なパラメータが不足しています'}), 400
        
        file_path_obj = Path(file_path)
        if not file_path_obj.exists():
            return jsonify({'success': False, 'error': 'ファイルが見つかりません'}), 404
        
        wb = openpyxl.load_workbook(file_path_obj, data_only=True)
        
        if sheet_name not in wb.sheetnames:
            wb.close()
            return jsonify({'success': False, 'error': 'シートが見つかりません'}), 404
        
        sheet = wb[sheet_name]
        
        # 周辺のセル情報を取得
        context_data = []
        start_row = max(1, row - context_rows)
        end_row = min(sheet.max_row, row + context_rows)
        
        for r in range(start_row, end_row + 1):
            row_data = []
            for c in range(1, sheet.max_column + 1):
                cell = sheet.cell(row=r, column=c)
                cell_info = {
                    'row': r,
                    'col': c,
                    'value': str(cell.value) if cell.value is not None else '',
                    'is_target': (r == row and c == col),
                    'is_header': (r == 1)
                }
                row_data.append(cell_info)
            context_data.append(row_data)
        
        # ヒットしたセルの詳細情報
        target_cell = sheet.cell(row=row, column=col)
        
        result = {
            'success': True,
            'file_name': file_path_obj.name,
            'sheet_name': sheet_name,
            'target_cell': {
                'row': row,
                'col': col,
                'value': str(target_cell.value) if target_cell.value is not None else '',
                'keyword': keyword
            },
            'context': context_data,
            'max_row': sheet.max_row,
            'max_col': sheet.max_column
        }
        
        wb.close()
        return jsonify(result)
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/download-results', methods=['GET'])
def download_results():
    """
    検索結果のExcelファイルをダウンロード
    """
    try:
        file_path = request.args.get('file_path', '')
        if not file_path:
            return jsonify({'success': False, 'error': 'ファイルパスが指定されていません'}), 400
        
        file_path_obj = Path(file_path)
        if not file_path_obj.exists():
            return jsonify({'success': False, 'error': 'ファイルが見つかりません'}), 404
        
        return send_file(
            str(file_path_obj),
            as_attachment=True,
            download_name=file_path_obj.name
        )
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/open-excel-file', methods=['POST'])
def open_excel_file():
    """
    Excelファイルを開く（Windows環境）
    """
    try:
        data = request.json
        file_path = data.get('file_path', '')
        sheet_name = data.get('sheet_name', '')
        row = data.get('row', 0)
        col = data.get('col', 0)
        
        if not file_path:
            return jsonify({'success': False, 'error': 'ファイルパスが指定されていません'}), 400
        
        file_path_obj = Path(file_path)
        if not file_path_obj.exists():
            return jsonify({'success': False, 'error': 'ファイルが見つかりません'}), 404
        
        # Windows環境でExcelファイルを開く
        if platform.system() == 'Windows':
            # 特定のシートとセルに移動する場合は、COM経由でExcelを操作
            if WIN32COM_AVAILABLE and sheet_name and row > 0 and col > 0:
                try:
                    excel = win32com.client.Dispatch("Excel.Application")
                    excel.Visible = True
                    wb = excel.Workbooks.Open(str(file_path_obj))
                    
                    if sheet_name:
                        ws = wb.Worksheets(sheet_name)
                        ws.Activate()
                        if row > 0 and col > 0:
                            ws.Cells(row, col).Select()
                    
                    return jsonify({
                        'success': True,
                        'message': 'Excelファイルを開きました（シートとセルに移動しました）'
                    })
                except Exception as e:
                    # COM操作に失敗した場合は、通常の方法でファイルを開く
                    print(f"COM操作に失敗: {str(e)}")
                    os.startfile(str(file_path_obj))
                    return jsonify({
                        'success': True,
                        'message': 'Excelファイルを開きました'
                    })
            else:
                # 通常の方法でファイルを開く
                os.startfile(str(file_path_obj))
                return jsonify({
                    'success': True,
                    'message': 'Excelファイルを開きました'
                })
        else:
            # Windows以外の環境
            if platform.system() == 'Darwin':  # macOS
                subprocess.Popen(['open', str(file_path_obj)])
            else:  # Linux
                subprocess.Popen(['xdg-open', str(file_path_obj)])
            return jsonify({
                'success': True,
                'message': 'Excelファイルを開きました'
            })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/search-replace', methods=['POST'])
def search_replace_files():
    """
    フォルダ内の複数ファイルに対して一括検索・置換を実行
    正規表現対応
    """
    try:
        data = request.json
        folder_path = data.get('folder_path', '')
        search_pattern = data.get('search_pattern', '')
        replace_pattern = data.get('replace_pattern', '')
        use_regex = data.get('use_regex', False)
        file_extensions = data.get('file_extensions', ['.txt', '.csv', '.html', '.js', '.ts', '.tsx', '.jsx', '.py', '.json', '.xml', '.css'])
        preview_only = data.get('preview_only', True)  # プレビューのみか、実際に置換するか
        
        if not folder_path:
            return jsonify({'success': False, 'error': 'フォルダパスが指定されていません'}), 400
        
        if not search_pattern:
            return jsonify({'success': False, 'error': '検索パターンが指定されていません'}), 400
        
        folder = Path(folder_path)
        if not folder.exists() or not folder.is_dir():
            return jsonify({'success': False, 'error': '指定されたフォルダが見つかりません'}), 404
        
        # 対象ファイルを取得
        target_files = []
        for ext in file_extensions:
            target_files.extend(list(folder.glob(f'*{ext}')))
            target_files.extend(list(folder.rglob(f'**/*{ext}')))  # サブディレクトリも検索
        
        # 重複を除去
        target_files = list(set(target_files))
        
        if not target_files:
            return jsonify({'success': False, 'error': '対象ファイルが見つかりませんでした'}), 404
        
        results = []
        total_replacements = 0
        
        # 正規表現のコンパイル
        if use_regex:
            try:
                pattern = re.compile(search_pattern)
            except re.error as e:
                return jsonify({'success': False, 'error': f'正規表現エラー: {str(e)}'}), 400
        else:
            # 通常の文字列検索（エスケープ処理）
            escaped_pattern = re.escape(search_pattern)
            pattern = re.compile(escaped_pattern)
        
        # 各ファイルを処理
        for file_path in target_files:
            try:
                # Excelファイルかどうかを判定
                is_excel = file_path.suffix.lower() in ['.xlsx', '.xls']
                
                if is_excel:
                    # Excelファイルの処理
                    try:
                        wb = openpyxl.load_workbook(file_path, data_only=True)
                        file_result = {
                            'file_path': str(file_path),
                            'file_name': file_path.name,
                            'matches': [],
                            'total_matches': 0,
                            'replaced': False
                        }
                        
                        # バックアップを作成（置換実行前）
                        if not preview_only:
                            backup_path = file_path.with_suffix(file_path.suffix + '.bak')
                            shutil.copy2(file_path, backup_path)
                            file_result['backup_path'] = str(backup_path)
                        
                        # 各シートを処理
                        for sheet_name in wb.sheetnames:
                            ws = wb[sheet_name]
                            
                            # 各セルを走査
                            for row in ws.iter_rows():
                                for cell in row:
                                    if cell.value is None:
                                        continue
                                    
                                    # セルの値を文字列に変換
                                    cell_value = str(cell.value)
                                    
                                    # 検索実行
                                    matches = list(pattern.finditer(cell_value))
                                    
                                    if matches:
                                        for match in matches:
                                            file_result['total_matches'] += 1
                                            file_result['matches'].append({
                                                'line': cell.row,
                                                'start': match.start(),
                                                'end': match.end(),
                                                'match_text': match.group(),
                                                'line_content': cell_value,
                                                'context_before': cell_value[max(0, match.start()-50):match.start()],
                                                'context_after': cell_value[match.end():min(len(cell_value), match.end()+50)],
                                                'sheet': sheet_name,
                                                'column': cell.column_letter
                                            })
                                            
                                            # 置換実行（プレビューモードでない場合）
                                            if not preview_only:
                                                # セルの値を置換
                                                if use_regex:
                                                    new_value = pattern.sub(replace_pattern, cell_value)
                                                else:
                                                    new_value = cell_value.replace(search_pattern, replace_pattern)
                                                
                                                # セルに新しい値を設定
                                                cell.value = new_value
                                                total_replacements += 1
                        
                        # Excelファイルを保存（置換実行した場合）
                        if not preview_only and file_result['total_matches'] > 0:
                            wb.save(file_path)
                            file_result['replaced'] = True
                        
                        # 結果が1つでもあれば追加
                        if file_result['total_matches'] > 0:
                            results.append(file_result)
                        
                        wb.close()
                        
                    except Exception as excel_error:
                        results.append({
                            'file_path': str(file_path),
                            'file_name': file_path.name,
                            'error': f'Excelファイル処理エラー: {str(excel_error)}',
                            'matches': [],
                            'total_matches': 0
                        })
                else:
                    # テキストファイルの処理（既存の処理）
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        content = f.read()
                    
                    # 検索実行
                    matches = list(pattern.finditer(content))
                    
                    if matches:
                        file_result = {
                            'file_path': str(file_path),
                            'file_name': file_path.name,
                            'matches': [],
                            'total_matches': len(matches),
                            'replaced': False
                        }
                        
                        # 各マッチの情報を取得
                        for match in matches:
                            start_pos = match.start()
                            end_pos = match.end()
                            
                            # 該当行を取得
                            line_number = content[:start_pos].count('\n') + 1
                            line_start = content.rfind('\n', 0, start_pos) + 1
                            line_end = content.find('\n', end_pos)
                            if line_end == -1:
                                line_end = len(content)
                            line_content = content[line_start:line_end]
                            
                            file_result['matches'].append({
                                'line': line_number,
                                'start': start_pos,
                                'end': end_pos,
                                'match_text': match.group(),
                                'line_content': line_content,
                                'context_before': content[max(0, start_pos-50):start_pos],
                                'context_after': content[end_pos:min(len(content), end_pos+50)]
                            })
                        
                        # 置換実行（プレビューモードでない場合）
                        if not preview_only:
                            # バックアップを作成
                            backup_path = file_path.with_suffix(file_path.suffix + '.bak')
                            shutil.copy2(file_path, backup_path)
                            
                            # 置換実行
                            if use_regex:
                                new_content = pattern.sub(replace_pattern, content)
                            else:
                                new_content = content.replace(search_pattern, replace_pattern)
                            
                            # ファイルに書き込み
                            with open(file_path, 'w', encoding='utf-8') as f:
                                f.write(new_content)
                            
                            file_result['replaced'] = True
                            file_result['backup_path'] = str(backup_path)
                            total_replacements += len(matches)
                        
                        results.append(file_result)
                    
            except Exception as e:
                results.append({
                    'file_path': str(file_path),
                    'file_name': file_path.name,
                    'error': str(e),
                    'matches': [],
                    'total_matches': 0
                })
        
        return jsonify({
            'success': True,
            'results': results,
            'total_files': len(target_files),
            'files_with_matches': len([r for r in results if r.get('total_matches', 0) > 0]),
            'total_replacements': total_replacements,
            'preview_only': preview_only
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/get-file-path', methods=['POST'])
def get_file_path():
    """
    アップロードされたファイルから絶対パスを取得
    """
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'ファイルが指定されていません'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'error': 'ファイルが指定されていません'}), 400
        
        # 一時ファイルとして保存してパスを取得
        import tempfile
        original_filename = file.filename
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, file.filename)
        file.save(temp_file)
        
        # 絶対パスを取得
        abs_path = os.path.abspath(temp_file)
        
        # ファイルの親ディレクトリのパスを返す
        parent_dir = os.path.dirname(abs_path)
        
        # 一時ファイルを削除
        try:
            os.remove(temp_file)
        except:
            pass
        
        return jsonify({
            'success': True,
            'file_path': abs_path,
            'parent_dir': parent_dir,
            'filename': original_filename,
            'message': '注意: これは一時ファイルのパスです。元のファイルのパスを取得するには、ファイルを直接ドロップするか、フォルダパスを手動で入力してください。'
        })
        
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        app.logger.error(f"Error in get_file_path: {error_trace}")
        return jsonify({
            'success': False,
            'error': f'ファイルパス取得中にエラーが発生しました: {str(e)}'
        }), 500


@app.route('/api/browse-folder', methods=['POST'])
def browse_folder():
    """
    フォルダ選択ダイアログを開く（Windows環境）
    注意: Webサーバー環境ではGUIダイアログが開けないため、この機能は制限されます
    """
    try:
        # Webサーバー環境ではGUIダイアログを開くことができない
        # 代わりに、エラーメッセージを返す
        
        # 環境変数からデフォルトフォルダを取得（設定されている場合）
        default_folder = os.environ.get('DEFAULT_SEARCH_FOLDER', '')
        
        if default_folder and Path(default_folder).exists():
            return jsonify({
                'success': True,
                'folder_path': default_folder,
                'message': 'デフォルトフォルダを使用します'
            })
        
        # GUIダイアログを試みる（ローカル環境でのみ動作）
        try:
            import tkinter as tk
            from tkinter import filedialog
            
            # ディスプレイが利用可能かチェック
            if platform.system() == 'Windows':
                try:
                    # Tkinterのルートウィンドウを非表示で作成
                    root = tk.Tk()
                    root.withdraw()  # メインウィンドウを非表示
                    root.attributes('-topmost', True)  # 最前面に表示
                    
                    # フォルダ選択ダイアログを開く
                    folder_path = filedialog.askdirectory(title='検索対象フォルダを選択')
                    
                    root.destroy()
                    
                    if folder_path:
                        # 完全パスを正規化
                        folder_path = os.path.abspath(folder_path)
                        return jsonify({
                            'success': True,
                            'folder_path': folder_path,
                            'message': f'フォルダが選択されました: {folder_path}'
                        })
                    else:
                        return jsonify({
                            'success': False,
                            'error': 'フォルダが選択されませんでした'
                        })
                except Exception as tk_error:
                    # Tkinter関連のエラー
                    app.logger.error(f"Tkinter error: {str(tk_error)}")
                    return jsonify({
                        'success': False,
                        'error': 'GUI環境が利用できません。フォルダパスを手動で入力してください。'
                    })
            else:
                return jsonify({
                    'success': False,
                    'error': 'フォルダ選択機能はWindows環境でのみ利用可能です'
                })
                
        except ImportError:
            # tkinterが利用できない場合
            return jsonify({
                'success': False,
                'error': 'フォルダ選択機能は利用できません。フォルダパスを手動で入力してください。'
            })
        except Exception as e:
            # GUI関連のエラー
            error_msg = str(e)
            app.logger.error(f"Browse folder error: {error_msg}")
            if 'display' in error_msg.lower() or 'DISPLAY' in error_msg:
                return jsonify({
                    'success': False,
                    'error': 'GUI環境が利用できません。フォルダパスを手動で入力してください。'
                })
            else:
                return jsonify({
                    'success': False,
                    'error': f'フォルダ選択中にエラーが発生しました: {error_msg}'
                })
            
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        app.logger.error(f"Unexpected error in browse_folder: {error_trace}")
        return jsonify({
            'success': False,
            'error': f'予期しないエラーが発生しました: {str(e)}'
        }), 500


@app.route('/api/get-folder-path', methods=['POST'])
def get_folder_path():
    """
    アップロードされたファイルからフォルダの完全パスを取得
    注意: ブラウザのセキュリティ制限により、元のファイルパスは取得できません
    このエンドポイントは、バックエンドのフォルダ選択ダイアログを使用することを推奨します
    """
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'ファイルが指定されていません'}), 400
        
        file = request.files['file']
        folder_name = request.form.get('folder_name', '')
        
        if file.filename == '':
            return jsonify({'success': False, 'error': 'ファイルが指定されていません'}), 400
        
        # ブラウザのセキュリティ制限により、元のファイルパスは取得できません
        # 代わりに、バックエンドのフォルダ選択ダイアログを使用することを推奨
        return jsonify({
            'success': False,
            'error': 'ブラウザのセキュリティ制限により、フォルダの完全パスを取得できません。\n「フォルダ選択」ボタンを使用して、サーバー側でフォルダを選択してください。',
            'folder_name': folder_name,
            'suggestion': 'バックエンドのフォルダ選択ダイアログ（/api/browse-folder）を使用してください'
        })
        
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        app.logger.error(f"Error in get_folder_path: {error_trace}")
        return jsonify({
            'success': False,
            'error': f'フォルダパス取得中にエラーが発生しました: {str(e)}'
        }), 500


@app.route('/api/health', methods=['GET'])
def health_check():
    """ヘルスチェック"""
    return jsonify({'status': 'ok', 'message': 'Excel Search API is running'})


if __name__ == '__main__':
    import logging
    logging.basicConfig(level=logging.INFO)
    app.logger.setLevel(logging.INFO)
    print("Starting Flask server on http://localhost:5001")
    print("API endpoints:")
    print("  - POST /api/search")
    print("  - POST /api/search-files")
    print("  - POST /api/get-cell-details")
    print("  - POST /api/open-excel-file")
    print("  - POST /api/search-replace")
    print("  - GET /api/health")
    app.run(debug=True, port=5001, host='0.0.0.0')
