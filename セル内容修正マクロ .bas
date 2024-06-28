Attribute VB_Name = "Module1"
Sub Main()
    ' 1. "エクセルファイルパス"を開く。
    Dim sourceWorkbook As Workbook
    Set sourceWorkbook = Workbooks.Open("エクセルファイルパス")
    
    ' 2. "C:\work\OPEN_DATA"直下のSQLファイルを名前の昇順で一つ開く。
    Dim filePath As String
    filePath = Dir("C:\work\フォルダパス\*.SQL")
    
    While filePath <> ""
        ' 3. 2で開いたファイルの拡張子を取り除いた文字列と部分一致する.xlsファイルを"フォルダパス"のフォルダ直下から開く。
        Dim targetWorkbook As Workbook
        Set targetWorkbook = Workbooks.Open("C:フォルダパス\" & "ビュー_" & Replace(filePath, ".SQL", ".xls"))
    
        ' 1-1. 1で開いたEXCELファイルの"変更履歴"シートの内容をコピーし、3で開いたEXCELファイルの"変更履歴"シートへ貼り付ける。
        
        sourceWorkbook.Sheets("変更履歴").Range("A1:BF45").CurrentRegion.Copy Destination:=targetWorkbook.Sheets("変更履歴").Range("A1:BF45")
    
        ' 1-2. 3で開いたファイルの先頭文字"ビュー_"と拡張子を取り除いた文字列をAG4セルとP6セルへ貼り付ける
        Dim sheetName As String
        sheetName = Replace(filePath, ".SQL", "")
        
        targetWorkbook.Sheets("データ項目定義").Range("AG4").Value = sheetName
        targetWorkbook.Sheets("データ項目定義").Range("P6").Value = sheetName
    
        ' 1-3. 3で開いたEXCELファイルの"データ項目定義"シートのCF1セルの内容を"2024/06/17"とCF2セルの内容を"今井"にする。
        targetWorkbook.Sheets("データ項目定義").Range("CF1").Value = "2024/06/17"
        targetWorkbook.Sheets("データ項目定義").Range("CF2").Value = "名前"
    
        ' 1-4. 3で開いたEXCELファイルの"20ビュー定義書"シートの4行目から下の内容を削除し、2で開いたテキスト内容を文字コードをSJIS変換し、B4セルへ貼り付ける。
        Dim textContent As String
        Open "C:\work\OPEN_DATA\" & filePath For Input As #1
        Do Until EOF(1)
            Line Input #1, textContent
        Loop
        Close #1
        targetWorkbook.Sheets("20ビュー生成定義").Range("B4").Value = StrConv(textContent, vbFromUnicode)
    
        ' 1-5. 3で開いたEXCELファイルの"データ項目定義"シートのBI1セルの内容を"2024/06/17"とBI2セルの内容を"今井"にする。
        targetWorkbook.Sheets("データ項目定義").Range("CF1").Value = "2024/06/17"
        targetWorkbook.Sheets("データ項目定義").Range("CF2").Value = "名前"
    
        ' 1-6. 3で開いたEXCELファイルの"50インデックス定義"シートのBI1セルの内容を"2024/06/17"とBI2セルの内容を"今井"にする。
        targetWorkbook.Sheets("50インデックス定義").Range("BI1").Value = "2024/06/17"
        targetWorkbook.Sheets("50インデックス定義").Range("BI2").Value = "名前"
    
        ' 1-7. 3で開いたEXCELファイルの"データ項目定義"シートを先頭から2番目に移動する
        targetWorkbook.Sheets("データ項目定義").Move After:=targetWorkbook.Sheets(2)
    
        ' 4. 1.と2.と3.で開いたファイルをそれぞれ保存し、ファイルを閉じる。リソースを開放する。
        
'        sourceWorkbook.Save
        targetWorkbook.Save
'        sourceWorkbook.Close
        targetWorkbook.Close
'        Set sourceWorkbook = Nothing
        Set targetWorkbook = Nothing
        
        ' 次のファイルを開く
        filePath = Dir
    Wend
End Sub
