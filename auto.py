import os
import pandas as pd
from datetime import datetime
import glob

def create_excel_file_list():
    """
    指定フォルダ内のExcelファイル一覧を作成し、新しいExcelファイルとして保存する
    """
    # 対象フォルダのパス
    target_folder = r"C:\Users\yukik\Desktop\excel_auto2025052"
    
    # フォルダが存在するかチェック
    if not os.path.exists(target_folder):
        print(f"フォルダが見つかりません: {target_folder}")
        return
    
    # Excelファイルの拡張子パターン
    excel_extensions = ['*.xlsx', '*.xls', '*.xlsm', '*.xlsb']
    
    # Excelファイルのリストを作成
    excel_files = []
    
    for extension in excel_extensions:
        pattern = os.path.join(target_folder, extension)
        files = glob.glob(pattern)
        excel_files.extend(files)
    
    # ファイル情報を整理
    file_info_list = []
    
    for file_path in excel_files:
        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)
        # ファイルサイズをKBに変換
        file_size_kb = round(file_size / 1024, 2)
        
        # 更新日時を取得
        modification_time = os.path.getmtime(file_path)
        modification_date = datetime.fromtimestamp(modification_time).strftime('%Y-%m-%d %H:%M:%S')
        
        file_info_list.append({
            'ファイル名': file_name,
            'ファイルサイズ(KB)': file_size_kb,
            '更新日時': modification_date,
            'フルパス': file_path
        })
    
    # DataFrameを作成
    df = pd.DataFrame(file_info_list)
    
    # ファイル名でソート
    df = df.sort_values('ファイル名').reset_index(drop=True)
    
    # 出力ファイル名を生成（現在の日時を含む）
    current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_filename = f"Excel名称リスト_{current_time}.xlsx"
    output_path = os.path.join(target_folder, output_filename)
    
    # Excelファイルとして保存
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # メインシートに保存
            df.to_excel(writer, sheet_name='Excelファイル一覧', index=False)
            
            # サマリー情報を別シートに保存
            summary_data = {
                '項目': ['総ファイル数', '作成日時', '対象フォルダ'],
                '値': [len(df), datetime.now().strftime('%Y-%m-%d %H:%M:%S'), target_folder]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='サマリー', index=False)
            
            # ワークシートの調整
            workbook = writer.book
            worksheet = writer.sheets['Excelファイル一覧']
            
            # 列幅を自動調整
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"✅ Excel名称リストを作成しました: {output_filename}")
        print(f"📁 保存場所: {output_path}")
        print(f"📊 検出されたExcelファイル数: {len(df)}")
        
        # 検出されたファイル一覧を表示
        if len(df) > 0:
            print("\n📋 検出されたExcelファイル:")
            for i, row in df.iterrows():
                print(f"  {i+1:2d}. {row['ファイル名']} ({row['ファイルサイズ(KB)']}KB)")
        else:
            print("❌ 指定フォルダにExcelファイルが見つかりませんでした")
            
    except Exception as e:
        print(f"❌ エラーが発生しました: {str(e)}")

def main():
    """
    メイン実行関数
    """
    print("=" * 60)
    print("🔍 Excelファイル名称リスト作成ツール")
    print("=" * 60)
    
    create_excel_file_list()
    
    print("\n" + "=" * 60)
    print("処理完了")
    input("Enterキーを押して終了...")

if __name__ == "__main__":
    main()
