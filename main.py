import re
from pathlib import Path
import openpyxl
import openpyxl.cell

def main():
    # 各項目を標準入力から取得
    str_GrepCondition = input('条件(正規表現が使えます): ')      # 項目[条件]:     検索文字列
    str_GrepFilename  = input('ファイル(正規表現が使えます): ')  # 項目[ファイル]: 検索対象のファイル名
    str_GrepFolder    = input('フォルダ: ')                      # 項目[フォルダ]: 検索対象のフォルダパス

    for objPath_File in Path(str_GrepFolder).iterdir():
        # ファイルでないならスキップ
        if objPath_File.is_file() == False:
            continue

        # 項目[ファイル]に合うファイル名でないならスキップ
        if not re.fullmatch(str_GrepFilename, objPath_File.name):
            continue

        str_FilePath = objPath_File.as_posix()
        
        # ファイルをExcelブックとして読み込み
        try:
            try:
                objExcel = openpyxl.load_workbook(objPath_File, read_only=True)

            # Excelファイルでないならスキップ
            except Exception as objExp:
                print(f'{str_FilePath} {type(objExp)}: ', objExp)
                continue

            for objSheet_Target in objExcel.worksheets:
                for lst_Row in objSheet_Target:
                    for objCell in lst_Row:
                        # 空であればスキップ
                        if objCell.value is None:
                            continue
                        
                        # 項目[条件]に合うセル値でなければスキップ
                        if not re.search(str_GrepCondition, objCell.value):
                            continue

                        # ヒットしたため標準出力に表示
                        int_Row = objCell.row
                        int_Col = objCell.column
                        str_SheetName = objSheet_Target.title
                        str_CellValue = objCell.value
                        str_Print = f'{str_FilePath} {str_SheetName}({int_Row}, {int_Col}): {str_CellValue}'
                        print(str_Print)
        
        # エラーが発生した場合はExcelブックを閉じる
        except Exception as objExp:
            print(f'{objPath_File.as_posix()} {type(objExp)}: ', objExp)
            objExcel.close()

if __name__ == '__main__':
    main()