Attribute VB_Name = "macro"
'STEP1
Sub �������f�[�^�ǎ�()

    Dim gfc As GetFolderContents
    Set gfc = New GetFolderContents
    
    rs = MsgBox("�������f�[�^�ǎ�����s���܂����H", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "�������f�[�^�ǎ�𒆒f���܂����B"
        Exit Sub
    Else
        Call gfc.ListFoldersAndFiles
    End If

End Sub

'STEP2
Sub �o��()

    Dim ox As organizeXML
    Set ox = New organizeXML
    
    rs = MsgBox("�f�[�^���������s���܂����H", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "�f�[�^�����𒆒f���܂����B"
        Exit Sub
    Else
        Call ox.organizeXML
    End If

End Sub

Sub �t�@�C���o��() 'STEP3

    Dim dc As dataConvert
    Set dc = New dataConvert
    
    rs = MsgBox("CSV�捞�����s���܂����H", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "CSV�捞�𒆒f���܂����B"
        Exit Sub
    Else
        Call dc.fileOutput
    End If

End Sub

