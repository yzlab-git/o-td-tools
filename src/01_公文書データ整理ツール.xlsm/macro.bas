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
        MsgBox "�������f�[�^�ǎ悪�������܂����B"
    End If

End Sub

'STEP2
Sub �������f�[�^����()

    Dim ox As organizeXML
    Set ox = New organizeXML
    
    rs = MsgBox("�������f�[�^���������s���܂����H", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "�������f�[�^�����𒆒f���܂����B"
        Exit Sub
    Else
        Call ox.organizeXML
        MsgBox "�������f�[�^�������������܂����B"
    End If

End Sub
