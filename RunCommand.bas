Attribute VB_Name = "RunCommand"
'*******************************************************************************
'  ���[�N�u�b�N��������
'*******************************************************************************
Option Explicit

Private Const cnsYEN = "\"

' �R�}���h��`
Const CMD_CODE_DEF As Byte = &H1
Const CMD_AUT_HEIGHT As Byte = &H2
Const CMD_AUT_GEN As Byte = &H4

'�R�[�h��`���f���s�t���O
Public execCmd As Byte
    
'*******************************************************************************
' �u�v�{�^������������
'*******************************************************************************
Sub Button1_Click()
    execCmd = CMD_CODE_DEF
    Call ProcMain
End Sub

'*******************************************************************************
' �u�ꊇ�����v�{�^������������
'*******************************************************************************
Sub Button2_Click()
    execCmd = CMD_CODE_DEF Or _
                CMD_AUT_HEIGHT
    Call ProcMain
End Sub

'*******************************************************************************
' �uDB���ڂƂ̐��������킹�v�{�^������������
'*******************************************************************************
Sub AutoGen_Click()
    execCmd = CMD_AUT_GEN Or _
                CMD_AUT_HEIGHT
    Call ProcMain
End Sub

'*******************************************************************************
' ���C������
'*******************************************************************************
Sub ProcMain()

    Dim xlAPP As Application        ' Excel.Application
    Dim strPathName As String       ' �w��t�H���_��
    Dim strFileName As String       ' ���o�����t�@�C����
    Dim swESC As Boolean            ' Esc�L�[����
    
    Call debugLog(INIT_LOG, 0)
    
    ' ��t�H���_�̎Q�ƣ���t�H���_���̎擾(modFolderPicker1�Ɏ��e)
    strPathName = FolderDialog("�t�H���_���w�肵�ĉ�����", True)
    If strPathName = "" Then Exit Sub
    
    ' �w��t�H���_����Excel���[�N�u�b�N�̃t�@�C�������Q�Ƃ���(1����)
    strFileName = Dir(strPathName & "\*.xls", vbNormal)
    If strFileName = "" Then
        MsgBox "���̃t�H���_�ɂ�Excel���[�N�u�b�N�͑��݂��܂���B"
        Exit Sub
    End If
    
    Set xlAPP = Application
    With xlAPP
        .ScreenUpdating = False             ' ��ʕ`���~
        .EnableEvents = False               ' �C�x���g�����~
        .EnableCancelKey = xlErrorHandler   ' Esc�L�[�ŃG���[�g���b�v����
        .Cursor = xlWait                    ' �J�[�\���������v�ɂ���
    End With
    On Error GoTo Button1_Click_ESC
    
    Dim commands As Collection
    Set commands = GetCommands()
        
    ' �w��t�H���_�̑SExcel���[�N�u�b�N�ɂ��ČJ��Ԃ�
    Do While strFileName <> ""
        ' Esc�L�[�Ō�����
        DoEvents
        If swESC = True Then
            ' ���f����̂������b�Z�[�W�Ŋm�F
            If MsgBox("���f�L�[��������܂����B�����ŏI�����܂����H", _
                vbInformation + vbYesNo) = vbYes Then
                GoTo Button1_Click_EXIT
            Else
                swESC = False
            End If
        End If

        '-----------------------------------------------------------------------
        ' ���������P�t�@�C���P�ʂ̏���
        Call OneWorkbookProc(xlAPP, commands, strPathName, strFileName)
        
        '-----------------------------------------------------------------------
        ' ���̃t�@�C�������Q��
        strFileName = Dir
    Loop
    GoTo Button1_Click_EXIT
    
'----------------
' Esc�L�[�E�o�p�s���x��
Button1_Click_ESC:
    If Err.Number = 18 Then
        ' Esc�L�[�ł̃G���[Raise
        swESC = True
        Resume
    ElseIf Err.Number = 1004 Then
        ' �B���V�[�g�����ΏۂȂ��̎��s���G���[�͖���
        Resume Next
    Else
        ' ���̑��̃G���[�̓��b�Z�[�W�\����I��
        MsgBox Err.Description
    End If

'----------------
' �����I��
Button1_Click_EXIT:
    MsgBox "�I�����܂���"
    With xlAPP
        .StatusBar = False                  ' �X�e�[�^�X�o�[�𕜋A
        .EnableEvents = True                ' �C�x���g����ĊJ
        .EnableCancelKey = xlInterrupt      ' Esc�L�[�����߂�
        .Cursor = xlDefault                 ' �J�[�\������̫�Ăɂ���
        .ScreenUpdating = True              ' ��ʕ`��ĊJ
    End With
    Set xlAPP = Nothing
End Sub

'*******************************************************************************
' ���s�R�}���h���X�g�̐���
'*******************************************************************************
Private Function GetCommands() As Collection
    '------------------------------------------------
    ' �R�}���h����
    ' �V�[�g�ɑ΂��鏈���R�}���h�̃C���X�^���X�𐶐�����B
    ' �e�����R�}���h�Ɋ��蓖�Ă�ꂽ�r�b�g�l�ɂ���āA�Y���R�}���h�̃C���X�^���X�𐶐�����B
    '------------------------------------------------
    Dim commands As Collection      ' �R�}���h���X�g
    Set commands = New Collection

    If (execCmd And CMD_CODE_DEF) > 0 Then
        Call commands.Add(New CmdCodeDef)
    End If
    If (execCmd And CMD_AUT_GEN) > 0 Then
        Call commands.Add(New CmdAutoGen)
    End If
    If (execCmd And CMD_AUT_HEIGHT) > 0 Then
        Call commands.Add(New CmdRowAdjust)
    End If
    
    Set GetCommands = commands
End Function


'*******************************************************************************
' �P�̃��[�N�u�b�N�̏���
'*******************************************************************************
Private Sub OneWorkbookProc(xlAPP As Application, _
                            commands As Collection, _
                            strPathName As String, _
                            strFileName As String)
    '---------------------------------------------------------------------------
    Dim objWBK As Workbook          ' ���[�N�u�b�NObject
    Dim aSheet As Worksheet          ' ���[�N�u�b�NObject
    ' �X�e�[�^�X�o�[�ɏ����t�@�C������\��
    xlAPP.StatusBar = strFileName & " �R�[�h��`���f��"
    ' ���[�N�u�b�N���J��
    Set objWBK = Workbooks.Open(fileName:=strPathName & cnsYEN & strFileName, _
                                UpdateLinks:=False, _
                                ReadOnly:=False)
    '---------------------------------------------------------------------------
    ' �V�[�g���̏���
    '---------------------------------------------------------------------------
    For Each aSheet In objWBK.Sheets
        Call OneSheetProc(commands, aSheet)
    Next

    objWBK.Close SaveChanges:=True
    xlAPP.StatusBar = False
    Set objWBK = Nothing
End Sub


'*******************************************************************************
' �P�̃V�[�g�̏���
'*******************************************************************************
Private Sub OneSheetProc(commands As Collection, _
                            targetSheet As Worksheet)
    
    Dim cmd As Variant
    For Each cmd In commands
        Call cmd.ExecCommand(targetSheet)
    Next
    
End Sub


                                                      

