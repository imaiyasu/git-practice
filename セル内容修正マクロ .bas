Attribute VB_Name = "Module1"
Sub Main()
    ' 1. "�G�N�Z���t�@�C���p�X"���J���B
    Dim sourceWorkbook As Workbook
    Set sourceWorkbook = Workbooks.Open("�G�N�Z���t�@�C���p�X")
    
    ' 2. "C:\work\OPEN_DATA"������SQL�t�@�C���𖼑O�̏����ň�J���B
    Dim filePath As String
    filePath = Dir("C:\work\�t�H���_�p�X\*.SQL")
    
    While filePath <> ""
        ' 3. 2�ŊJ�����t�@�C���̊g���q����菜����������ƕ�����v����.xls�t�@�C����"�t�H���_�p�X"�̃t�H���_��������J���B
        Dim targetWorkbook As Workbook
        Set targetWorkbook = Workbooks.Open("C:�t�H���_�p�X\" & "�r���[_" & Replace(filePath, ".SQL", ".xls"))
    
        ' 1-1. 1�ŊJ����EXCEL�t�@�C����"�ύX����"�V�[�g�̓��e���R�s�[���A3�ŊJ����EXCEL�t�@�C����"�ύX����"�V�[�g�֓\��t����B
        
        sourceWorkbook.Sheets("�ύX����").Range("A1:BF45").CurrentRegion.Copy Destination:=targetWorkbook.Sheets("�ύX����").Range("A1:BF45")
    
        ' 1-2. 3�ŊJ�����t�@�C���̐擪����"�r���[_"�Ɗg���q����菜�����������AG4�Z����P6�Z���֓\��t����
        Dim sheetName As String
        sheetName = Replace(filePath, ".SQL", "")
        
        targetWorkbook.Sheets("�f�[�^���ڒ�`").Range("AG4").Value = sheetName
        targetWorkbook.Sheets("�f�[�^���ڒ�`").Range("P6").Value = sheetName
    
        ' 1-3. 3�ŊJ����EXCEL�t�@�C����"�f�[�^���ڒ�`"�V�[�g��CF1�Z���̓��e��"2024/06/17"��CF2�Z���̓��e��"����"�ɂ���B
        targetWorkbook.Sheets("�f�[�^���ڒ�`").Range("CF1").Value = "2024/06/17"
        targetWorkbook.Sheets("�f�[�^���ڒ�`").Range("CF2").Value = "���O"
    
        ' 1-4. 3�ŊJ����EXCEL�t�@�C����"20�r���[��`��"�V�[�g��4�s�ڂ��牺�̓��e���폜���A2�ŊJ�����e�L�X�g���e�𕶎��R�[�h��SJIS�ϊ����AB4�Z���֓\��t����B
        Dim textContent As String
        Open "C:\work\OPEN_DATA\" & filePath For Input As #1
        Do Until EOF(1)
            Line Input #1, textContent
        Loop
        Close #1
        targetWorkbook.Sheets("20�r���[������`").Range("B4").Value = StrConv(textContent, vbFromUnicode)
    
        ' 1-5. 3�ŊJ����EXCEL�t�@�C����"�f�[�^���ڒ�`"�V�[�g��BI1�Z���̓��e��"2024/06/17"��BI2�Z���̓��e��"����"�ɂ���B
        targetWorkbook.Sheets("�f�[�^���ڒ�`").Range("CF1").Value = "2024/06/17"
        targetWorkbook.Sheets("�f�[�^���ڒ�`").Range("CF2").Value = "���O"
    
        ' 1-6. 3�ŊJ����EXCEL�t�@�C����"50�C���f�b�N�X��`"�V�[�g��BI1�Z���̓��e��"2024/06/17"��BI2�Z���̓��e��"����"�ɂ���B
        targetWorkbook.Sheets("50�C���f�b�N�X��`").Range("BI1").Value = "2024/06/17"
        targetWorkbook.Sheets("50�C���f�b�N�X��`").Range("BI2").Value = "���O"
    
        ' 1-7. 3�ŊJ����EXCEL�t�@�C����"�f�[�^���ڒ�`"�V�[�g��擪����2�ԖڂɈړ�����
        targetWorkbook.Sheets("�f�[�^���ڒ�`").Move After:=targetWorkbook.Sheets(2)
    
        ' 4. 1.��2.��3.�ŊJ�����t�@�C�������ꂼ��ۑ����A�t�@�C�������B���\�[�X���J������B
        
'        sourceWorkbook.Save
        targetWorkbook.Save
'        sourceWorkbook.Close
        targetWorkbook.Close
'        Set sourceWorkbook = Nothing
        Set targetWorkbook = Nothing
        
        ' ���̃t�@�C�����J��
        filePath = Dir
    Wend
End Sub
