Attribute VB_Name = "Module1"
Sub CopySheetFromReferenceFile()
    ' �Q�ƃt�@�C���̃p�X
    Dim referenceFilePath As String
    referenceFilePath = "�t�@�C���p�X"
    
    ' �V�[�g��
    Dim sheetName As String
    sheetName1 = "�ύX����"
    sheetName2 = "�f�[�^���ڒ�`"
    sheetName3 = "20�r���[������`"
    sheetName4 = "50�C���f�b�N�X��`"
    
    ' �ۑ���t�H���_�̃p�X
    Dim destinationFolderPath As String
    destinationFolderPath = "�ۑ���t�H���_�p�X"
    
    ' �Q�ƃt�@�C�����J��
    Dim referenceWorkbook As Workbook
    Set referenceWorkbook = Workbooks.Open(referenceFilePath)
    
    ' �Q�ƃt�@�C�����̃V�[�g���R�s�[
    ' �ύX����
    Dim referenceSheet1 As Worksheet
    Set referenceSheet1 = referenceWorkbook.sheets(sheetName1)
    ' �f�[�^���ڒ�`
    Dim referenceSheet2 As Worksheet
    Set referenceSheet2 = referenceWorkbook.sheets(sheetName2)
    ' 20�r���[������`
    Dim referenceSheet3 As Worksheet
    Set referenceSheet3 = referenceWorkbook.sheets(sheetName3)
    ' 50�C���f�b�N�X��`
    Dim referenceSheet4 As Worksheet
    Set referenceSheet4 = referenceWorkbook.sheets(sheetName4)
    
    ' �ۑ���t�H���_���̃t�@�C�������[�v
    Dim destinationFile As String
    destinationFile = Dir(destinationFolderPath & "\*.xls*")
    
    Do While destinationFile <> ""
        ' �ۑ���t�@�C�����J��
        Dim destinationWorkbook As Workbook
        Set destinationWorkbook = Workbooks.Open(destinationFolderPath & "\" & destinationFile)
        
        ' �V�[�g���R�s�[
        referenceSheet1.Copy Before:=destinationWorkbook.sheets(1)
        referenceSheet2.Copy Before:=destinationWorkbook.sheets(2)
        referenceSheet3.Copy Before:=destinationWorkbook.sheets(3)
        referenceSheet4.Copy Before:=destinationWorkbook.sheets(4)
        
        ' �ۑ����ăN���[�Y
        destinationWorkbook.Save
        destinationWorkbook.Close
        
        ' ���̃t�@�C����
        destinationFile = Dir
    Loop
    
    ' �Q�ƃt�@�C�����N���[�Y
    referenceWorkbook.Close
    
    ' ���\�[�X�����
    Set referenceSheet = Nothing
    Set referenceWorkbook = Nothing
End Sub
