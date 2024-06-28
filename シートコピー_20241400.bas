Attribute VB_Name = "Module1"
Sub CopySheetFromReferenceFile()
    ' 参照ファイルのパス
    Dim referenceFilePath As String
    referenceFilePath = "ファイルパス"
    
    ' シート名
    Dim sheetName As String
    sheetName1 = "変更履歴"
    sheetName2 = "データ項目定義"
    sheetName3 = "20ビュー生成定義"
    sheetName4 = "50インデックス定義"
    
    ' 保存先フォルダのパス
    Dim destinationFolderPath As String
    destinationFolderPath = "保存先フォルダパス"
    
    ' 参照ファイルを開く
    Dim referenceWorkbook As Workbook
    Set referenceWorkbook = Workbooks.Open(referenceFilePath)
    
    ' 参照ファイル内のシートをコピー
    ' 変更履歴
    Dim referenceSheet1 As Worksheet
    Set referenceSheet1 = referenceWorkbook.sheets(sheetName1)
    ' データ項目定義
    Dim referenceSheet2 As Worksheet
    Set referenceSheet2 = referenceWorkbook.sheets(sheetName2)
    ' 20ビュー生成定義
    Dim referenceSheet3 As Worksheet
    Set referenceSheet3 = referenceWorkbook.sheets(sheetName3)
    ' 50インデックス定義
    Dim referenceSheet4 As Worksheet
    Set referenceSheet4 = referenceWorkbook.sheets(sheetName4)
    
    ' 保存先フォルダ内のファイルをループ
    Dim destinationFile As String
    destinationFile = Dir(destinationFolderPath & "\*.xls*")
    
    Do While destinationFile <> ""
        ' 保存先ファイルを開く
        Dim destinationWorkbook As Workbook
        Set destinationWorkbook = Workbooks.Open(destinationFolderPath & "\" & destinationFile)
        
        ' シートをコピー
        referenceSheet1.Copy Before:=destinationWorkbook.sheets(1)
        referenceSheet2.Copy Before:=destinationWorkbook.sheets(2)
        referenceSheet3.Copy Before:=destinationWorkbook.sheets(3)
        referenceSheet4.Copy Before:=destinationWorkbook.sheets(4)
        
        ' 保存してクローズ
        destinationWorkbook.Save
        destinationWorkbook.Close
        
        ' 次のファイルへ
        destinationFile = Dir
    Loop
    
    ' 参照ファイルをクローズ
    referenceWorkbook.Close
    
    ' リソースを解放
    Set referenceSheet = Nothing
    Set referenceWorkbook = Nothing
End Sub
