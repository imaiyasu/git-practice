Attribute VB_Name = "Module1"
Sub Main()
    ' 1. "C:\work\DB�o�͏��"�t�H���_���̃t�@�C�����J��
    Dim filePath As String
    filePath = "C:�t�H���_�p�X\"
    
    Dim fileName As String
    fileName = Dir(filePath & "*.xls*")
    
    While fileName <> ""
        Dim wb As Workbook
        Set wb = Workbooks.Open(filePath & fileName)
        
        Dim ws As Worksheet
        Set ws = wb.Sheets("�r���[��`��")
        
        ' 1-1. A7�Z������Z������̍s�܂ł̓��e��ϐ��Ɋi�[����
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
        
        ' 2. "�t�H���_�p�X"�t�H���_���̊����t�@�C�����J��
        Dim targetPath As String
        targetPath = "C:\work\�e�[�u��View��`\"
        
        Dim targetFileName As String
        targetFileName = "�r���[_" & Left(fileName, Len(fileName) - 5) & ".xls"
        
        Dim targetWb As Workbook
        
        Set targetWb = Workbooks.Open(targetPath & targetFileName)

        Set targetWs = targetWb.Sheets("�f�[�^���ڒ�`")
        
        ' 2-1. A14�Z����1-1.�Ŏ擾�����ϐ��̓��e��1�s���\��t����
        targetWs.Range("A14").Resize(UBound(data1, 1) + 1, 1).Value = data1
        
        ' 2-2. B14�Z����1-2.�Ŏ擾�����ϐ��̓��e��1�s���\��t����
        targetWs.Range("B14").Resize(UBound(data2, 1) + 1, 1).Value = data2
        
        ' 2-3. AE14�Z����1-3.�Ŏ擾�����ϐ��̓��e��1�s���\��t����
        targetWs.Range("AE14").Resize(UBound(data3, 1) + 1, 1).Value = data3
        
        ' 2-4. AL14�Z����1-4.�Ŏ擾�����ϐ��̓��e��1�s���\��t����
        targetWs.Range("AL14").Resize(UBound(data4, 1) + 1, 1).Value = data4
        
        ' 2-5. AQ14�Z����1-5.�Ŏ擾�����ϐ��̓��e��1�s���\��t����
        targetWs.Range("AQ14").Resize(UBound(data5, 1) + 1, 1).Value = data5
        
        ' 2-6. BA14�Z����1-6.�Ŏ擾�����ϐ��̓��e��1�s���\��t����
        targetWs.Range("BA14").Resize(UBound(data6, 1) + 1, 1).Value = data6
        
        ' 2-7. BF14�Z����1-7.�Ŏ擾�����ϐ��̓��e��1�s���\��t����
        targetWs.Range("BF14").Resize(UBound(data7, 1) + 1, 1).Value = data7
        
        ' 3. 1.��2.�ŊJ�����t�@�C�������ꂼ��ۑ����A�t�@�C�������
        targetWb.Save
        targetWb.Close
        
        wb.Save
        wb.Close
        
        ' ���\�[�X���J������
        Set targetWs = Nothing
        Set targetWb = Nothing
        Set ws = Nothing
        Set wb = Nothing
        
        ' ���̃t�@�C�����J��
        fileName = Dir
    Wend
End Sub
