Attribute VB_Name = "Module1"
Sub Main()
    ' 1. "C:\work\DB出力情報"フォルダ内のファイルを開く
    Dim filePath As String
    filePath = "C:フォルダパス\"
    
    Dim fileName As String
    fileName = Dir(filePath & "*.xls*")
    
    While fileName <> ""
        Dim wb As Workbook
        Set wb = Workbooks.Open(filePath & fileName)
        
        Dim ws As Worksheet
        Set ws = wb.Sheets("ビュー定義書")
        
        ' 1-1. A7セルからセルが空の行までの内容を変数に格納する
        Dim rowCount As Long
        rowCount = ws.Range("A7").End(xlDown).Row
        
        Dim data1 As Variant
        data1 = ws.Range("A7:A" & rowCount).Value
        
        Dim data2 As Variant
        data2 = ws.Range("B7:B" & rowCount).Value
        
        Dim data3 As Variant
        data3 = ws.Range("C7:C" & rowCount).Value
        
        Dim data4 As Variant
        data4 = ws.Range("D7:D" & rowCount).Value
        
        Dim data5 As Variant
        data5 = ws.Range("E7:E" & rowCount).Value
        
        Dim data6 As Variant
        data6 = ws.Range("F7:F" & rowCount).Value
        
        Dim data7 As Variant
        data7 = ws.Range("G7:G" & rowCount).Value
        
        ' 2. "フォルダパス"フォルダ内の既存ファイルを開く
        Dim targetPath As String
        targetPath = "C:\work\テーブルView定義\"
        
        Dim targetFileName As String
        targetFileName = "ビュー_" & Left(fileName, Len(fileName) - 5) & ".xls"
        
        Dim targetWb As Workbook
        
        Set targetWb = Workbooks.Open(targetPath & targetFileName)

        Set targetWs = targetWb.Sheets("データ項目定義")
        
        ' 2-1. A14セルへ1-1.で取得した変数の内容を1行ずつ貼り付ける
        targetWs.Range("A14").Resize(UBound(data1, 1) + 1, 1).Value = data1
        
        ' 2-2. B14セルへ1-2.で取得した変数の内容を1行ずつ貼り付ける
        targetWs.Range("B14").Resize(UBound(data2, 1) + 1, 1).Value = data2
        
        ' 2-3. AE14セルへ1-3.で取得した変数の内容を1行ずつ貼り付ける
        targetWs.Range("AE14").Resize(UBound(data3, 1) + 1, 1).Value = data3
        
        ' 2-4. AL14セルへ1-4.で取得した変数の内容を1行ずつ貼り付ける
        targetWs.Range("AL14").Resize(UBound(data4, 1) + 1, 1).Value = data4
        
        ' 2-5. AQ14セルへ1-5.で取得した変数の内容を1行ずつ貼り付ける
        targetWs.Range("AQ14").Resize(UBound(data5, 1) + 1, 1).Value = data5
        
        ' 2-6. BA14セルへ1-6.で取得した変数の内容を1行ずつ貼り付ける
        targetWs.Range("BA14").Resize(UBound(data6, 1) + 1, 1).Value = data6
        
        ' 2-7. BF14セルへ1-7.で取得した変数の内容を1行ずつ貼り付ける
        targetWs.Range("BF14").Resize(UBound(data7, 1) + 1, 1).Value = data7
        
        ' 3. 1.と2.で開いたファイルをそれぞれ保存し、ファイルを閉じる
        targetWb.Save
        targetWb.Close
        
        wb.Save
        wb.Close
        
        ' リソースを開放する
        Set targetWs = Nothing
        Set targetWb = Nothing
        Set ws = Nothing
        Set wb = Nothing
        
        ' 次のファイルを開く
        fileName = Dir
    Wend
End Sub
