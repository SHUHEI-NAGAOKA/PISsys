import pandas as pd
import openpyxl
import os

# --- .xlsファイルを.xlsxファイルに変換する関数 (変更なし) ---
def convert_xls_to_xlsx(xls_file_path, output_folder=None):
    """
    .xls形式のExcelファイルを読み込み、.xlsx形式で保存します。

    Args:
        xls_file_path (str): 変換したい.xlsファイルの完全なパス。
        output_folder (str, optional): 変換後の.xlsxファイルを保存するフォルダのパス。
                                       指定しない場合、元の.xlsファイルと同じフォルダに保存されます。

    Returns:
        str or None: 変換後の.xlsxファイルの完全なパス。変換に失敗した場合は None。
    """
    if not os.path.exists(xls_file_path):
        print(f"エラー: 元のファイル '{xls_file_path}' が見つかりません。")
        return None

    if not xls_file_path.lower().endswith('.xls'):
        print(f"エラー: 指定されたファイル '{xls_file_path}' は.xls形式ではありません。")
        return None

    base_name = os.path.basename(xls_file_path)
    file_name_without_ext = os.path.splitext(base_name)[0]
    xlsx_file_name = f"{file_name_without_ext}.xlsx"

    if output_folder:
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            print(f"出力フォルダ '{output_folder}' を作成しました。")
        output_path = os.path.join(output_folder, xlsx_file_name)
    else:
        output_path = os.path.join(os.path.dirname(xls_file_path), xlsx_file_name)

    try:
        print(f"'{xls_file_path}' を読み込み中...")
        xls = pd.ExcelFile(xls_file_path)

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  シート '{sheet_name}' を変換しました。")

        print(f"ファイルの変換が完了しました: '{output_path}'")
        return output_path

    except Exception as e:
        print(f"ファイルの変換中にエラーが発生しました: {e}")
        print("ファイルが破損しているか、アクセス権の問題があるか、")
        print("または xlrd ライブラリのインストールが不完全である可能性があります。")
        return None

# --- セル値を抽出するヘルパー関数 ---
def _get_cell_value(worksheet, cell_address):
    """
    指定されたワークシートから特定のセル（結合セル対応）の値を抽出します。
    """
    try:
        target_cell = worksheet[cell_address]
        
        is_merged = False
        for merged_range in worksheet.merged_cells.ranges:
            if cell_address in merged_range: 
                is_merged = True
                top_left_cell_address = merged_range.coord.split(':')[0]
                # print(f"  セル '{cell_address}' は結合セル '{merged_range.coord}' の一部です。")
                # print(f"  結合元のセル '{top_left_cell_address}' の値を参照します。")
                return worksheet[top_left_cell_address].value
        
        if not is_merged:
            # print(f"  セル '{cell_address}' は結合セルではないか、結合元のセルです。")
            return target_cell.value
            
    except Exception as e:
        print(f"警告: セル '{cell_address}' の処理中にエラーが発生しました: {e}")
        return None # エラーの場合はNoneを返す

# --- 全てのシートから辞書形式で指定された複数のセルペアを抽出する関数 (変更点あり) ---
def extract_specified_cell_pairs_from_all_sheets(excel_file_path, cell_address_map):
    """
    Excelファイルの全てのシートから、辞書形式で指定された複数のセルペアの値を抽出します。

    Args:
        excel_file_path (str): 処理するExcelファイルの完全なパス（.xlsx形式である必要があります）。
        cell_address_map (dict): 抽出したいセルのペアを定義する辞書。
                                 キーと値がそれぞれ抽出したいセルアドレスとなる。
                                 例: {'A8': 'V8', 'A10': 'V10'}

    Returns:
        dict: 各シート名をキーとし、そのシートから抽出されたセルの値を格納した辞書。
              セルの値は、元のセルアドレスをキーとする別の辞書として格納されます。
              例: {
                  'Sheet1': {'A8': 'A8の値', 'V8': 'V8の値', 'A10': 'A10の値', 'V10': 'V10の値'},
                  'Sheet2': {'A8': 'A8の値2', 'V8': 'V8の値2', 'A10': 'A10の値2', 'V10': 'V10の値2'}
              }
              ファイルが見つからない場合やエラーの場合は空の辞書を返します。
    """
    extracted_data_by_sheet = {}

    if not os.path.exists(excel_file_path):
        print(f"エラー: Excelファイル '{excel_file_path}' が見つかりません。")
        return {}

    if not excel_file_path.lower().endswith('.xlsx'):
        print(f"エラー: 抽出対象のファイル '{excel_file_path}' は.xlsx形式ではありません。")
        print("この関数は.xlsxファイルのみをサポートします。")
        return {}

    try:
        workbook = openpyxl.load_workbook(excel_file_path, data_only=True)
        
        for sheet_name in workbook.sheetnames:
            print(f"\n--- シート '{sheet_name}' を処理中 ---")
            ws = workbook[sheet_name]
            
            # このシートの抽出結果を格納する辞書
            sheet_extracted_cells = {} 

            # cell_address_map のキーと値を両方抽出
            for primary_cell, secondary_cell in cell_address_map.items():
                
                # primary_cell (例: A8) の値を取得
                value_primary = _get_cell_value(ws, primary_cell)
                sheet_extracted_cells[primary_cell] = value_primary
                print(f"  セル '{primary_cell}' の値: {value_primary}")

                # secondary_cell (例: V8) の値を取得
                value_secondary = _get_cell_value(ws, secondary_cell)
                sheet_extracted_cells[secondary_cell] = value_secondary
                print(f"  セル '{secondary_cell}' の値: {value_secondary}")
            
            extracted_data_by_sheet[sheet_name] = sheet_extracted_cells

    except Exception as e:
        print(f"Excelファイルの読み込み中にエラーが発生しました: {e}")
        print("ファイル形式が間違っている可能性があります（.xlsではなく.xlsxである必要があります）。")
        return {}
    
    return extracted_data_by_sheet

# --- メインの実行ブロック ---
if __name__ == "__main__":
    # --- 1. 変換したい.xlsファイルのパスを設定 ---
    input_xls_file_path = '/Users/nagaokashuuhei/Desktop/sys_practice/PISsys/history_200 (1).xls'

    # --- 2. 変換後の.xlsxファイルを保存するフォルダを設定 (任意) ---
    output_directory_for_conversion = '/Users/nagaokashuuhei/Desktop/sys_practice/PISsys。' 

    # --- 3. 変換後の.xlsxファイルから抽出したいセルのペアを辞書形式で設定 ---
    # キーと値がそれぞれ抽出したいセルアドレスになります。
    # 例: 'A8'の値を抽出し、それに対応する'V8'の値も抽出する
    target_cell_pairs_to_extract = {
        'A8': 'V8',
        'A10': 'V10',
        'A12': 'V12' # 別のペアの例
    } 

    # --- 変換処理の実行 ---
    converted_xlsx_path = convert_xls_to_xlsx(input_xls_file_path, output_directory_for_conversion)

    if converted_xlsx_path:
        print(f"\n変換されたExcelファイル: {converted_xlsx_path}")
        print("\n--- 変換されたファイルからのセル抽出を開始します ---")
        
        # --- 抽出処理の実行 ---
        all_sheets_cells_data = extract_specified_cell_pairs_from_all_sheets(
            converted_xlsx_path, 
            target_cell_pairs_to_extract
        )

        if all_sheets_cells_data:
            print("\n--- 全シートからの複数セルペア抽出結果 ---")
            for sheet, cells_data in all_sheets_cells_data.items():
                print(f"シート '{sheet}':")
                for cell_addr, value in cells_data.items():
                    print(f"  セル '{cell_addr}' の値 -> {value}")
                print("-" * 20) 
        else:
            print("\nセルの抽出に失敗しました。変換されたファイルの内容を確認してください。")
    else:
        print("\nファイルの変換に失敗したため、セル抽出処理をスキップします。")