Attribute VB_Name = "macro"
Dim im As invoiceManager

'�t�@�C���p�X�ݒ�
Sub �������t�@�C���I��()

    Set im = New invoiceManager
    rs = MsgBox("�������t�@�C���p�X��ݒ肵�܂����H", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "�������t�@�C���p�X�̐ݒ�𒆒f���܂����B"
        Exit Sub
    Else
        Call im.getFilePath
    End If

End Sub

'�t�@�C���p�X�ݒ�
Sub ���[���쐬()

    Set im = New invoiceManager
    rs = MsgBox("���[�����쐬���܂����H", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "���[���쐬�𒆒f���܂����B"
        Exit Sub
    Else
        Call im.CreateEmailWithPDFInvoices
    End If

End Sub


' �e�X�g�p�@���t�惊�X�g�V�[�g�̃V�[�g���𐿋����t�@�C���̃V�[�g���ɂ���
Sub CopyStringsToSheetNames()
    Dim wbA As Workbook
    Dim wbB As Workbook
    Dim wsA As Worksheet
    Dim i As Integer
    Dim sheetName As String
    
    ' ���ɊJ���Ă���A�t�@�C����B�t�@�C�����擾
    Set wbA = Workbooks("���[�����M_���}�N��.xlsm") ' A�t�@�C�����J���Ă���ꍇ
    Set wsA = wbA.Sheets("���t�惊�X�g") ' A�t�@�C���́u���t�惊�X�g�v�V�[�g���w��

    Set wbB = Workbooks("������.xlsx") ' B�t�@�C�����J���Ă���ꍇ
    
    ' D2����D101�܂ł̃��[�v����
    For i = 2 To 101
        ' A�t�@�C����D�񂩂當������擾
        sheetName = wsA.Cells(i, 4).Value ' D���4��ڂɊY��
        
            
            ' �V�����V�[�g��B�t�@�C���ɒǉ�
         wbB.Sheets.Add(After:=wbB.Sheets(wbB.Sheets.Count)).Name = sheetName
        

NextSheet:
    Next i
    
    MsgBox "D2����D101�܂ł̃V�[�g����B�t�@�C���ɒǉ�����܂����B"
End Sub
