Attribute VB_Name = "macro"
Sub �������f�[�^����()

    Dim sc As startCategorize
    Set sc = New startCategorize
    
    rs = MsgBox("�������f�[�^���������s���܂����H", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "�������f�[�^�����𒆒f���܂����B"
        Exit Sub
    Else
        Call sc.startCategorize
        MsgBox "�������f�[�^�������������܂����B"
    End If

End Sub

Sub ������()

    Dim sc As startCategorize
    Set sc = New startCategorize
    
    rs = MsgBox("�����������s���܂����H" & vbLf & vbLf & "���ȉ��̏����폜���܂��B" & vbLf _
    & "�E�ǎ�t�H���_" & vbLf & "�E�o�̓t�H���_" & vbLf & "�E���s����" & vbLf & "�E��������" & vbLf & "�E����������", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "�������𒆒f���܂����B"
        Exit Sub
    Else
        Call sc.deleteValue
        MsgBox "���������������܂����B"
    End If

End Sub
