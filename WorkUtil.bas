Attribute VB_Name = "WorkUtil"
'*******************************************************************************
'  ���[�N�u�b�N��������
'*******************************************************************************
Option Explicit

Private Const cnsYEN = "\"

'*******************************************************************************
'  ���V�[�g��
'*******************************************************************************
Private Const SHEET_NAME_COVER_OLD = "�\��"
Private Const SHEET_NAME_HISTORY_OLD = "��������"
Private Const SHEET_NAME_INOUT_OLD = "1.���o�͒�`"
Private Const SHEET_NAME_TYPE_LIST_OLD = "2.�v�f�ꗗ"

'*******************************************************************************
'  �V�V�[�g��
'*******************************************************************************
Private Const SHEET_NAME_COVER_NEW = "�\��"
Private Const SHEET_NAME_HISTORY_NEW = "��������"
Private Const SHEET_NAME_REF_LIST_NEW = "1.�Q�ƃt�@�C���ꗗ"
Private Const SHEET_NAME_COMMON_NEW = "2.���ʏ��"
Private Const SHEET_NAME_TYPE_LIST_NEW = "3.�^�ꗗ"
Private Const SHEET_NAME_INOUT_NEW = "4.���o�͒�`"
Private Const TEMPLATE_SHEET_NAME = "�^��`_�J�[�g�f�ޗ\�񉞓�"

'*******************************************************************************
'  �^�V�[�g�v���t�B�b�N�X
'*******************************************************************************
Private Const TARGET_SHEET_PREFIX_OLD = "4."
Private Const TARGET_SHEET_PREFIX_NEW = "�^��`_"
Private Const TEMPLATE_TYPE_DEF_SHEET_MAX_ROW = 50
Private Const TEMPLATE_REF_LIST_SHEET_MAX_ROW = 50
Private Const TEMPLATE_TYPE_LIST_SHEET_MAX_ROW = 50
'�e���v���[�g�t�@�C���̃t�@�C���p�X
Private Const TEMPLATE_FILE_PATH = "\\192.168.10.250\common\AP�`�[�����L�b��\03_�O������\02_�K���\01_�݌v���K��\06_���V�X�e���֘A\���V�X�e���C���^�[�t�F�[�X�d�l��_IFAWS000I_�J�[�g�f�ޗ\��API.xlsx"

'*******************************************************************************
'  �j�����[�h
' 1 -- �V�[�g�폜
' 2 -- �V�[�g���ύX
'*******************************************************************************
Private Const DESTRUCTION_MODE_DELETE = 1
Private Const DESTRUCTION_MODE_RENAME = 2
Private Const DESTRUCTION_MODE = DESTRUCTION_MODE_DELETE


'*******************************************************************************
'���V�X�e���C���^�[�t�F�[�X�d�l��_���^��`�V�[�g�i�v�f�ꗗ�V�[�g�j_���`
'*******************************************************************************
Private Const COL_OLD_TYPEDEF_NO = 1 ' NO��
Private Const COL_OLD_TYPEDEF_NAME = 2 ' ���ږ�
Private Const COL_OLD_TYPEDEF_ID = 3 ' ����ID
Private Const COL_OLD_TYPEDEF_EXPLAIN = 4 ' ����
Private Const COL_OLD_TYPEDEF_REQIRED = 5 ' �K�{
Private Const COL_OLD_TYPEDEF_COND_REQ = 6 ' �K�{����
Private Const COL_OLD_TYPEDEF_MIN_LOOP = 7 ' �ŏ��J��Ԃ�
Private Const COL_OLD_TYPEDEF_MAX_LOOP = 8 ' �ő�J��Ԃ�
Private Const COL_OLD_TYPEDEF_TYPE = 9 ' �^
Private Const COL_OLD_TYPEDEF_TYPE_DETAIL = 10 ' �^�ڍ�
Private Const COL_OLD_TYPEDEF_MINLENGTH = 11 ' �ŏ�����
Private Const COL_OLD_TYPEDEF_MAXLENGTH = 12 ' �ő包��
Private Const COL_OLD_TYPEDEF_DATA_SAMPLE = 13 ' �f�[�^��
Private Const COL_OLD_TYPEDEF_BIKO = 14 ' ���l

'���O
Private Const COL_OLD_TYPEDEF_DB_TYPE = COL_OLD_TYPEDEF_BIKO + 1 ' DB�^
Private Const COL_OLD_TYPEDEF_NOT_NULL = COL_OLD_TYPEDEF_BIKO + 4 ' DB_Notnull


'*******************************************************************************
'���V�X�e���C���^�[�t�F�[�X�d�l��_�V�^��`�V�[�g_���`
'*******************************************************************************
Private Const COL_NEW_TYPEDEF_NO = 1 ' NO��
Private Const COL_NEW_TYPEDEF_NAME = 2 ' ���ږ�
Private Const COL_NEW_TYPEDEF_ID = 3 ' ����ID
Private Const COL_NEW_TYPEDEF_EXPLAIN = 4 ' ����
Private Const COL_NEW_TYPEDEF_MIN_LOOP = 5 ' �ŏ��J��Ԃ�
Private Const COL_NEW_TYPEDEF_MAX_LOOP = 6 ' �ő�J��Ԃ�
Private Const COL_NEW_TYPEDEF_COND_REQ = 7 ' �K�{����
Private Const COL_NEW_TYPEDEF_LIMIT_NUM = 8 ' �J��Ԃ��񐔐���
Private Const COL_NEW_TYPEDEF_TYPE = 9 ' �^
Private Const COL_NEW_TYPEDEF_TYPE_NAME = 10 ' �^��
Private Const COL_NEW_TYPEDEF_FILE_ID = 11 ' �t�@�C��ID
Private Const COL_NEW_TYPEDEF_CHAR_FORMAT = 12 ' ������E����
Private Const COL_NEW_TYPEDEF_MINLENGTH = 13 ' �ŏ�����
Private Const COL_NEW_TYPEDEF_MAXLENGTH = 14 ' �ő包��
Private Const COL_NEW_TYPEDEF_DATA_SAMPLE = 15 ' �f�[�^��
Private Const COL_NEW_TYPEDEF_BIKO = 16 ' ���l
Private Const COL_NEW_TYPEDEF_MIN_NO = COL_NEW_TYPEDEF_NO ' �ō��̗�ԍ�
Private Const COL_NEW_TYPEDEF_MAX_NO = COL_NEW_TYPEDEF_BIKO ' �ŉE�̗�ԍ�

'*******************************************************************************
'���V�X�e���C���^�[�t�F�[�X�d�l��_�^�ꗗ�V�[�g_���`
'*******************************************************************************
Private Const COL_TYPELIST_NO = 1 ' NO��
Private Const COL_TYPELIST_LNAME = 2 ' �^�_����
Private Const COL_TYPELIST_PNAME = 3 ' �^������
Private Const COL_TYPELIST_BIKO = 4 ' ���l
Private Const COL_TYPELIST_MIN_NO = COL_TYPELIST_NO ' �ō��̗�ԍ�
Private Const COL_TYPELIST_MAX_NO = COL_TYPELIST_BIKO ' �ŉE�̗�ԍ�

'*******************************************************************************
'���V�X�e���C���^�[�t�F�[�X�d�l��_�Q�ƃt�@�C���ꗗ�V�[�g_���`
'*******************************************************************************
Private Const COL_REFLIST_NO = 1 ' ����
Private Const COL_REFLIST_FILE_NAME = 2 ' �t�@�C����
Private Const COL_REFLIST_FILE_ID = 3 ' �t�@�C��ID
Private Const COL_REFLIST_BIKO = 4 ' ���l
Private Const COL_REFLIST_MIN_NO = COL_REFLIST_NO ' �ō��̗�ԍ�
Private Const COL_REFLIST_MAX_NO = COL_REFLIST_BIKO ' �ŉE�̗�ԍ�

'*******************************************************************************
'���V�X�e���C���^�[�t�F�[�X�d�l��_�����o�͒�`�V�[�g_���`
'*******************************************************************************
Private Const COL_OLD_IO_NO = 1 ' ����
Private Const COL_OLD_IO_IO = 2 ' �t�@�C����
Private Const COL_OLD_IO_ROOT_ELEMENT_ID = 4 ' �ŏ�ʗv�fID
Private Const COL_OLD_IO_ELEMENT_NAME = 5 ' �ŏ�ʗv�f��
Private Const COL_OLD_IO_BIKO = 6 ' ���l

'*******************************************************************************
'���V�X�e���C���^�[�t�F�[�X�d�l��_�V���o�͒�`�V�[�g_���`
'*******************************************************************************
Private Const COL_NEW_IO_NO = 1 ' ����
Private Const COL_NEW_IO_IO = 2 ' ���o�͋敪
Private Const COL_NEW_IO_ROOT_ELEMENT_ID = 3 ' �ŏ�ʗv�fID
Private Const COL_NEW_IO_ROOT_ELEMENT_NAME = 4 ' �ŏ�ʗv�f��
Private Const COL_NEW_IO_TYPE_NAME = 5 ' �^��
Private Const COL_NEW_IO_FILE_ID = 6 ' �t�@�C��ID
Private Const COL_NEW_IO_BIKO = 7 ' ���l
Private Const COL_NEW_IO_MIN_NO = COL_NEW_IO_NO ' �ō��̗�ԍ�
Private Const COL_NEW_IO_MAX_NO = COL_NEW_IO_BIKO ' �ŉE�̗�ԍ�

'*******************************************************************************
'���ʌ^�ꗗ��`�V�[�g
'*******************************************************************************
Private Const SHEET_NAME_TYPEDEF = "���ʌ^��`"


'*******************************************************************************
' ���V�[�g������v�f�ꗗ�쐬
'*******************************************************************************
Public Sub ���V�[�g������v�f�ꗗ�쐬()
    Call Do�V�[�g������v�f�ꗗ�쐬(TARGET_SHEET_PREFIX_OLD)
End Sub

'*******************************************************************************
' �֐��J�n���O
'*******************************************************************************
Public Sub LogStart(subName As String, symbol As String)
    Call debugLog("Start:" & subName & " # " & symbol, OUTPUT_LOG)
End Sub

'*******************************************************************************
' �֐��I�����O
'*******************************************************************************
Public Sub LogEnd(subName As String, symbol As String)
    Call debugLog("End:" & subName & " # " & symbol, OUTPUT_LOG)
End Sub

'*******************************************************************************
' �G���[�n���h�����O���O
'*******************************************************************************
Public Sub LogErrorHandle(desc As String, subName As String, file As String)
    Call debugLog("Error:" & desc & " in " & subName & "@" & file, OUTPUT_LOG)
End Sub

'*******************************************************************************
' �G���[�n���h�����O���O
'*******************************************************************************
Public Sub LogOutput(logtext As String, file As String)
    Call debugLog(logtext & "@" & file, OUTPUT_LOG)
End Sub

'*******************************************************************************
' �V�V�[�g������v�f�ꗗ�쐬
'*******************************************************************************
Public Sub �V�V�[�g������v�f�ꗗ�쐬()
    
    Call Do�V�[�g������v�f�ꗗ�쐬(TARGET_SHEET_PREFIX_NEW)
End Sub
'*******************************************************************************
' �V�[�g����v�f�ꗗ�쐬�����֐�
'*******************************************************************************
Public Sub Do�V�[�g������v�f�ꗗ�쐬(targetPrefix As String)
    Call debugLog(INIT_LOG, 0)
    Dim xlAPP As Application        ' Excel.Application
    Dim strPathName As String       ' �w��t�H���_��
    Dim strFileName As String       ' ���o�����t�@�C����
    Dim swESC As Boolean            ' Esc�L�[����
    
    Call LogStart("Do�V�[�g������v�f�ꗗ�쐬", "")
    ' ��t�H���_�̎Q�ƣ���t�H���_���̎擾(modFolderPicker1�Ɏ��e)
    strPathName = FolderDialog("�t�H���_���w�肵�ĉ�����", True)
    If strPathName = "" Then Exit Sub
    
    ' �w��t�H���_����Excel���[�N�u�b�N�̃t�@�C�������Q�Ƃ���(1����)
    strFileName = Dir(strPathName & "\*.xls", vbNormal)
    If strFileName = "" Then
        MsgBox "���̃t�H���_�ɂ�Excel���[�N�u�b�N�͑��݂��܂���B"
        Exit Sub
    End If
    
    Dim files As Collection
    Set files = New Collection
    Dim vFileName As Variant
    Do While strFileName <> ""
        Call files.Add(strFileName)
        ' ���̃t�@�C�������Q��
        strFileName = Dir
    Loop
    'For Each vFileName In files
    
    Set xlAPP = Application
    With xlAPP
        .ScreenUpdating = False             ' ��ʕ`���~
        .EnableEvents = False               ' �C�x���g�����~
        .EnableCancelKey = xlErrorHandler   ' Esc�L�[�ŃG���[�g���b�v����
        .Cursor = xlWait                    ' �J�[�\���������v�ɂ���
    End With
    On Error GoTo Error_Handler
        
    ' �w��t�H���_�̑SExcel���[�N�u�b�N�ɂ��ČJ��Ԃ�
    For Each vFileName In files
        ' Esc�L�[�Ō�����
        DoEvents
        If swESC = True Then
            ' ���f����̂������b�Z�[�W�Ŋm�F
            If MsgBox("���f�L�[��������܂����B�����ŏI�����܂����H", _
                vbInformation + vbYesNo) = vbYes Then
                GoTo Error_Handler
            Else
                swESC = False
            End If
        End If

        '-----------------------------------------------------------------------
        ' ���������P�t�@�C���P�ʂ̏���
        Call WB�v�f�ꗗ�쐬����(xlAPP, strPathName, CStr(vFileName), targetPrefix)
        
        '-----------------------------------------------------------------------
    Next
    
    GoTo Sub_EXIT
    
'----------------
' Esc�L�[�E�o�p�s���x��
Error_Handler:
    
    If Err.Number = 18 Then
        Call LogErrorHandle("Esc�L�[�ł̃G���[Raise", "Do�V�[�g������v�f�ꗗ�쐬", "")
        ' Esc�L�[�ł̃G���[Raise
        swESC = True
        Resume
    ElseIf Err.Number = 1004 Then
        ' �B���V�[�g�����ΏۂȂ��̎��s���G���[�͖���
        Resume Next
    Else
        Call LogErrorHandle(Err.Description, "Do�V�[�g������v�f�ꗗ�쐬", "")
        ' ���̑��̃G���[�̓��b�Z�[�W�\����I��
        MsgBox Err.Description
    End If

    Resume Sub_EXIT
'----------------
' �����I��
Sub_EXIT:
    With xlAPP
        .StatusBar = False                  ' �X�e�[�^�X�o�[�𕜋A
        .EnableEvents = True                ' �C�x���g����ĊJ
        .EnableCancelKey = xlInterrupt      ' Esc�L�[�����߂�
        .Cursor = xlDefault                 ' �J�[�\������̫�Ăɂ���
        .ScreenUpdating = True              ' ��ʕ`��ĊJ
    End With
    Set xlAPP = Nothing
    Call LogEnd("Do�V�[�g������v�f�ꗗ�쐬", "")
    MsgBox "�I�����܂���"
End Sub


'*******************************************************************************
' �P�̃��[�N�u�b�N�̏���
'*******************************************************************************
Private Sub WB�v�f�ꗗ�쐬����(xlAPP As Application, _
                            strPathName As String, _
                            strFileName As String, _
                            targetPrefix As String)
    On Error GoTo Error_Handler
    '---------------------------------------------------------------------------
    Dim objWBK As Workbook          ' ���[�N�u�b�NObject
    
    Call LogStart("WB�v�f�ꗗ�쐬����", "")
    ' �X�e�[�^�X�o�[�ɏ����t�@�C������\��
    xlAPP.StatusBar = strFileName & " ������..."
    ' ���[�N�u�b�N���J��
    Set objWBK = OpenWorkBook(strPathName & cnsYEN & strFileName, False, False)
    If Not objWBK Is Nothing Then
    
        Call Do�v�f�ꗗ�쐬����(objWBK, targetPrefix)
        objWBK.Close SaveChanges:=True

    End If
    
    GoTo Sub_EXIT

Error_Handler:
    'Resume Sub_EXIT
    Err.Raise 5001
   
Sub_EXIT:
    Call LogEnd("WB�v�f�ꗗ�쐬����", "")
    xlAPP.StatusBar = False
    Set objWBK = Nothing

End Sub

'*******************************************************************************
' �v�f�ꗗ(�^�ꗗ)�쐬����
'*******************************************************************************
Private Sub Do�v�f�ꗗ�쐬����(objWBK As Workbook, targetPrefix As String)
    '---------------------------------------------------------------------------
    Dim aSheet As Worksheet         ' ���[�N�u�b�NObject
    Dim elmListSheet As Worksheet
    
    Call LogStart("Do�v�f�ꗗ�쐬����", "")
    
    '�v�f�ꗗ�i�^�ꗗ�j�V�[�g�擾
    Set elmListSheet = GetTypeListSheet(objWBK)
    
    If elmListSheet Is Nothing Then
        Exit Sub
    End If
    '---------------------------------------------------------------------------
    ' �v�f���X�g�̃G���g���[���폜
    '---------------------------------------------------------------------------
    Dim i As Integer
    Dim maxNum As Integer: maxNum = elmListSheet.Cells(elmListSheet.Rows.count, 1).End(xlUp).row
    
    If maxNum < TEMPLATE_TYPE_LIST_SHEET_MAX_ROW Then
        maxNum = TEMPLATE_TYPE_LIST_SHEET_MAX_ROW
    End If
    For i = maxNum To 2 Step -1
        elmListSheet.Rows(i).Delete
    Next
    
    '---------------------------------------------------------------------------
    ' �V�[�g���̏���
    '---------------------------------------------------------------------------
    i = 2
    For Each aSheet In objWBK.Sheets
        If InStr(aSheet.Name, targetPrefix) = 1 Then
            elmListSheet.Cells(i, COL_TYPELIST_LNAME).Value = GetTypeNameFromSheetName(aSheet, targetPrefix)
            elmListSheet.Cells(i, COL_TYPELIST_NO).Value = "=Row()-1"
            With elmListSheet
              .Range(.Cells(i, COL_TYPELIST_NO), .Cells(i, COL_TYPELIST_MAX_NO)).Borders.LineStyle = xlContinuous
            End With
            i = i + 1
        End If
    Next
    
    Call LogEnd("Do�v�f�ꗗ�쐬����", "")
End Sub

'*******************************************************************************
' �Q�ƃt�@�C���ꗗ�쐬����
'*******************************************************************************
Private Sub Do�Q�ƃt�@�C���ꗗ�쐬����(objWBK As Workbook, list As Collection)
    '---------------------------------------------------------------------------
    Dim aSheet As Worksheet         ' ���[�N�u�b�NObject
    Dim refSheet As Worksheet
    
    Call LogStart("Do�Q�ƃt�@�C���ꗗ�쐬����", "")
    
    '�Q�ƃt�@�C���ꗗ�擾
    Set refSheet = GetSheet(objWBK, SHEET_NAME_REF_LIST_NEW)
    If refSheet Is Nothing Then
        Exit Sub
    End If
    
    '---------------------------------------------------------------------------
    ' �Q�ƃt�@�C���ꗗ�̃G���g���[���폜
    '---------------------------------------------------------------------------
    Dim j As Integer
    For j = TEMPLATE_REF_LIST_SHEET_MAX_ROW To 2 Step -1
        refSheet.Rows(j).Delete
    Next
    
    '---------------------------------------------------------------------------
    ' �V�[�g���̏���
    '---------------------------------------------------------------------------
    Dim rowIdx As Integer: rowIdx = 2
    Dim item As Variant
    For Each item In list
        refSheet.Cells(rowIdx, COL_REFLIST_NO).Value = "=Row()-1"
        refSheet.Cells(rowIdx, COL_REFLIST_FILE_NAME).Value = GenCommonTypeFileName(CStr(item))
        refSheet.Cells(rowIdx, COL_REFLIST_FILE_ID).Value = Mid(item, InStr(item, ":") + 1)
        With refSheet
          .Range(.Cells(rowIdx, COL_REFLIST_MIN_NO), .Cells(rowIdx, COL_REFLIST_MAX_NO)).Borders.LineStyle = xlContinuous
        End With
        rowIdx = rowIdx + 1
    Next
    Call LogEnd("Do�Q�ƃt�@�C���ꗗ�쐬����", "")
End Sub

'*******************************************************************************
' ���ʌ^�E�݌v���t�@�C�����̎擾
'*******************************************************************************
Private Function GenCommonTypeFileName(item As String)
    GenCommonTypeFileName = "���V�X�e���C���^�[�t�F�[�X�d�l��_IFASW" & _
                            Mid(item, InStr(item, ":") + 1) & _
                            "I_" & _
                            Mid(item, 1, InStr(item, ":") - 1) & ".xlsx"

End Function
'*******************************************************************************
' �v�f�ꗗ����V�[�g�쐬
'*******************************************************************************
Public Sub �v�f�ꗗ����V�[�g�쐬()
    Call LogStart("�v�f�ꗗ����V�[�g�쐬", "")
    
    Call debugLog(INIT_LOG, 0)
    Dim xlAPP As Application        ' Excel.Application
    Dim strPathName As String       ' �w��t�H���_��
    Dim strFileName As String       ' ���o�����t�@�C����
    Dim swESC As Boolean            ' Esc�L�[����
    Dim templateWBK As Workbook     ' Excel.Workbook(�e���v���[�g�t�@�C��)
    
    ' ��t�H���_�̎Q�ƣ���t�H���_���̎擾(modFolderPicker1�Ɏ��e)
    strPathName = FolderDialog("�t�H���_���w�肵�ĉ�����", True)
    If strPathName = "" Then Exit Sub
    
    ' �w��t�H���_����Excel���[�N�u�b�N�̃t�@�C�������Q�Ƃ���(1����)
    strFileName = Dir(strPathName & "\*.xls", vbNormal)
    If strFileName = "" Then
        MsgBox "���̃t�H���_�ɂ�Excel���[�N�u�b�N�͑��݂��܂���B"
        Exit Sub
    End If
    
    Dim files As Collection
    Set files = New Collection
    Dim vFileName As Variant
    Do While strFileName <> ""
        Call files.Add(strFileName)
        ' ���̃t�@�C�������Q��
        strFileName = Dir
    Loop
    'For Each vFileName In files
    
    Set xlAPP = Application
    With xlAPP
        .ScreenUpdating = False             ' ��ʕ`���~
        .EnableEvents = False               ' �C�x���g�����~
        .EnableCancelKey = xlErrorHandler   ' Esc�L�[�ŃG���[�g���b�v����
        .Cursor = xlWait                    ' �J�[�\���������v�ɂ���
    End With
    On Error GoTo Error_Handler
    
    ' �e���v���[�g���[�N�u�b�N���J��
    Set templateWBK = OpenWorkBook(TEMPLATE_FILE_PATH, False, True)

    If Not templateWBK Is Nothing Then
    
        ' �w��t�H���_�̑SExcel���[�N�u�b�N�ɂ��ČJ��Ԃ�
        For Each vFileName In files
            ' Esc�L�[�Ō�����
            DoEvents
            If swESC = True Then
                ' ���f����̂������b�Z�[�W�Ŋm�F
                If MsgBox("���f�L�[��������܂����B�����ŏI�����܂����H", _
                    vbInformation + vbYesNo) = vbYes Then
                    GoTo Error_Handler
                Else
                    swESC = False
                End If
            End If
    
            '-----------------------------------------------------------------------
            ' ���������P�t�@�C���P�ʂ̏���
            Call WB�^�ꗗ�V�[�g�쐬����(xlAPP, templateWBK, strPathName, CStr(vFileName))
            
            '-----------------------------------------------------------------------
        Next
        templateWBK.Close SaveChanges:=False
    
    End If
    
    GoTo Sub_EXIT
    
'----------------
' Esc�L�[�E�o�p�s���x��
Error_Handler:
    If Err.Number = 18 Then
        ' Esc�L�[�ł̃G���[Raise
        swESC = True
        Resume
    ElseIf Err.Number = 1004 Then
        ' �B���V�[�g�����ΏۂȂ��̎��s���G���[�͖���
        Resume Next
    Else
        Call LogErrorHandle(Err.Description, "�v�f�ꗗ����V�[�g�쐬", "")
        ' ���̑��̃G���[�̓��b�Z�[�W�\����I��
        MsgBox Err.Description
    End If

    Resume Sub_EXIT
'----------------
' �����I��
Sub_EXIT:

    With xlAPP
        .StatusBar = False                  ' �X�e�[�^�X�o�[�𕜋A
        .EnableEvents = True                ' �C�x���g����ĊJ
        .EnableCancelKey = xlInterrupt      ' Esc�L�[�����߂�
        .Cursor = xlDefault                 ' �J�[�\������̫�Ăɂ���
        .ScreenUpdating = True              ' ��ʕ`��ĊJ
    End With
    Set xlAPP = Nothing
    Call LogEnd("�v�f�ꗗ����V�[�g�쐬", "")
    MsgBox "�I�����܂���"
End Sub


'*******************************************************************************
' �^�ꗗ�V�[�g�쐬_�P���[�N�u�b�N�̏���
'*******************************************************************************
Private Sub WB�^�ꗗ�V�[�g�쐬����(xlAPP As Application, _
                            templateWBK As Workbook, _
                            strPathName As String, _
                            strFileName As String)
    On Error GoTo Error_Handler
    '---------------------------------------------------------------------------
    Call LogStart("WB�^�ꗗ�V�[�g�쐬����", strFileName)
    Dim objWBK As Workbook          ' ���[�N�u�b�NObject

    ' �X�e�[�^�X�o�[�ɏ����t�@�C������\��
    xlAPP.StatusBar = strFileName & " �^�ꗗ���f��"
    ' ���[�N�u�b�N���J��
    Set objWBK = OpenWorkBook(strPathName & cnsYEN & strFileName, False, False)
    
    If Not objWBK Is Nothing Then
    
        '�v�f�ꗗ�i�^�ꗗ�j�V�[�g�擾
        Dim elmListSheet As Worksheet
        Set elmListSheet = GetTypeListSheet(objWBK)
    
        Dim i As Integer, _
            typeName As String, _
            targetSheetName As String, _
            targetSheet As Worksheet
    
        For i = 2 To elmListSheet.Cells(elmListSheet.Rows.count, 1).End(xlUp).row
            '�^��
            typeName = elmListSheet.Cells(i, COL_NEW_TYPEDEF_NAME)
            '�V�[�g��
            targetSheetName = TARGET_SHEET_PREFIX_NEW & typeName
    
            ' ���݂��Ȃ��ꍇ�Ƀe���v���[�g����쐬����B
            If Not ContainsSheet(objWBK, targetSheetName) Then
                ' �R�s�[���s
                templateWBK.Worksheets(TEMPLATE_SHEET_NAME).Copy After:=objWBK.Worksheets(objWBK.Worksheets.count)
                objWBK.Worksheets(objWBK.Worksheets.count).Name = targetSheetName
                Set targetSheet = objWBK.Sheets(targetSheetName)
                If Not targetSheet Is Nothing Then
                    '---------------------------------------------------------------------------
                    ' �v�f���X�g�̃G���g���[���폜
                    '---------------------------------------------------------------------------
                    Dim j As Integer
                    For j = targetSheet.Cells(targetSheet.Rows.count, 1).End(xlUp).row To 2 Step -1
                        targetSheet.Rows(j).Delete
                    Next
                End If
            End If
        Next
    
        objWBK.Close SaveChanges:=True
    
    End If
    
    GoTo Sub_EXIT
    
Error_Handler:
    Call LogErrorHandle("5002", "WB�^�ꗗ�V�[�g�쐬����", strFileName)
    'Resume Sub_EXIT
    Err.Raise 5002
   
Sub_EXIT:
    Call LogEnd("WB�^�ꗗ�V�[�g�쐬����", strFileName)
    xlAPP.StatusBar = False
    Set objWBK = Nothing
    Set elmListSheet = Nothing
End Sub

'*******************************************************************************
' �^�ꗗ�V�[�g�쐬_�P���[�N�u�b�N�̏���
'*******************************************************************************
Private Sub WB�f�[�^�x�[�X����(xlAPP As Application, _
                            templateWBK As Workbook, _
                            strPathName As String, _
                            strFileName As String)
    On Error GoTo Error_Handler
    '---------------------------------------------------------------------------
    Call LogStart("WB�^�ꗗ�V�[�g�쐬����", strFileName)
    Dim objWBK As Workbook          ' ���[�N�u�b�NObject

    ' �X�e�[�^�X�o�[�ɏ����t�@�C������\��
    xlAPP.StatusBar = strFileName & " �^�ꗗ���f��"
    ' ���[�N�u�b�N���J��
    Set objWBK = OpenWorkBook(strPathName & cnsYEN & strFileName, False, False)
    
    If Not objWBK Is Nothing Then
    
        '�v�f�ꗗ�i�^�ꗗ�j�V�[�g�擾
        Dim elmListSheet As Worksheet
        Set elmListSheet = GetTypeListSheet(objWBK)
    
        Dim i As Integer, _
            typeName As String, _
            targetSheetName As String, _
            targetSheet As Worksheet
    
        For i = 2 To elmListSheet.Cells(elmListSheet.Rows.count, 1).End(xlUp).row
            '�^��
            typeName = elmListSheet.Cells(i, COL_NEW_TYPEDEF_NAME)
            '�V�[�g��
            targetSheetName = TARGET_SHEET_PREFIX_NEW & typeName
    
            ' ���݂��Ȃ��ꍇ�Ƀe���v���[�g����쐬����B
            If Not ContainsSheet(objWBK, targetSheetName) Then
                ' �R�s�[���s
                templateWBK.Worksheets(TEMPLATE_SHEET_NAME).Copy After:=objWBK.Worksheets(objWBK.Worksheets.count)
                objWBK.Worksheets(objWBK.Worksheets.count).Name = targetSheetName
                Set targetSheet = objWBK.Sheets(targetSheetName)
                If Not targetSheet Is Nothing Then
                    '---------------------------------------------------------------------------
                    ' �v�f���X�g�̃G���g���[���폜
                    '---------------------------------------------------------------------------
                    Dim j As Integer
                    For j = targetSheet.Cells(targetSheet.Rows.count, 1).End(xlUp).row To 2 Step -1
                        targetSheet.Rows(j).Delete
                    Next
                End If
            End If
        Next
    
        objWBK.Close SaveChanges:=True
    
    End If
    
    GoTo Sub_EXIT
    
Error_Handler:
    Call LogErrorHandle("5002", "WB�^�ꗗ�V�[�g�쐬����", strFileName)
    'Resume Sub_EXIT
    Err.Raise 5002
   
Sub_EXIT:
    Call LogEnd("WB�^�ꗗ�V�[�g�쐬����", strFileName)
    xlAPP.StatusBar = False
    Set objWBK = Nothing
    Set elmListSheet = Nothing
End Sub


'*******************************************************************************
' �P�̃��[�N�u�b�N���ɃV�[�g�����݂��邩�ǂ����̏���
'*******************************************************************************
Private Function ContainsSheet(targetWb As Workbook, sname As String) As Boolean
    Dim ws As Worksheet, flag As Boolean
    flag = False
    For Each ws In targetWb.Worksheets
        If ws.Name = sname Then
            flag = True
            Exit For
        End If
    Next ws
    ContainsSheet = flag
End Function

'*******************************************************************************
' �v�f�ꗗ�V�[�g�̎擾
' �V�t�H�[�}�b�g�̖��̂�����΂������D��
'*******************************************************************************
Private Function GetTypeListSheet(targetWb As Workbook) As Worksheet
    Dim ws As Worksheet, retWs As Worksheet
    If ContainsSheet(targetWb, SHEET_NAME_TYPE_LIST_NEW) Then
        Set retWs = GetSheet(targetWb, SHEET_NAME_TYPE_LIST_NEW)
    Else
        Set retWs = GetSheet(targetWb, SHEET_NAME_TYPE_LIST_OLD)
    End If
    Set GetTypeListSheet = retWs
End Function

'*******************************************************************************
' ���[�N�u�b�N����w��̖��O�̃��[�N�V�[�g��T���ĕԋp
'*******************************************************************************
Private Function GetSheet(targetWb As Workbook, targetSheet As String) As Worksheet
    Dim ws As Worksheet, retWs As Worksheet
    For Each ws In targetWb.Worksheets
        If ws.Name = targetSheet Then
            Set retWs = ws
            Exit For
        End If
    Next ws
    Set GetSheet = retWs
End Function

'*******************************************************************************
' ���[�N�u�b�N����W��V�[�g��T��
'*******************************************************************************
Private Function GetSummarySheet(targetWb As Workbook) As Worksheet
    Dim ws As Worksheet, retWs As Worksheet
    For Each ws In targetWb.Worksheets
        If InStr(ws.Name, "3.1") = 1 Then
            Set retWs = ws
            Exit For
        End If
    Next ws
    Set GetSummarySheet = retWs
End Function

'*******************************************************************************
' ���V�[�g������v�f�����擾
'*******************************************************************************
Private Function GetTypeNameFromSheetName(targetSheet As Worksheet, targetPrefix As String) As String
    Dim retVal As String, pos As Long
    If InStr(targetSheet.Name, targetPrefix) = 1 Then
        '���V�[�g��
        If targetPrefix = TARGET_SHEET_PREFIX_OLD Then
            '�Ō��.(�h�b�g�̈ʒu)
            pos = InStrRev(targetSheet.Name, ".")
            If pos > 0 Then
                ' �v�f�ꗗ�V�[�g�ɒǋL
                retVal = Mid(targetSheet.Name, pos + 1)
            End If
        End If

        '�V�V�[�g��
        If targetPrefix = TARGET_SHEET_PREFIX_NEW Then
            '�Ō��.(�h�b�g�̈ʒu)
            pos = InStrRev(targetSheet.Name, TARGET_SHEET_PREFIX_NEW)
            If pos = 1 Then
                ' �v�f�ꗗ�V�[�g�ɒǋL
                retVal = Mid(targetSheet.Name, Len(TARGET_SHEET_PREFIX_NEW) + 1)
            End If
        End If
    End If
    GetTypeNameFromSheetName = retVal
End Function

'*******************************************************************************
' ���ʌ^���o����
'*******************************************************************************
Public Sub ���ʌ^���o����()
    Call debugLog(INIT_LOG, 0)
    Dim xlAPP As Application        ' Excel.Application
    Dim strPathName As String       ' �w��t�H���_��
    Dim strFileName As String       ' ���o�����t�@�C����
    Dim swESC As Boolean            ' Esc�L�[����
    Dim templateWBK As Workbook     ' Excel.Workbook(�e���v���[�g�t�@�C��)
    Dim targetFiles As Collection   ' �Ώۃt�@�C���ꗗ
    
    ' ��t�H���_�̎Q�ƣ���t�H���_���̎擾(modFolderPicker1�Ɏ��e)
    strPathName = FolderDialog("�t�H���_���w�肵�ĉ�����", True)
    If strPathName = "" Then Exit Sub
    
    ' �w��t�H���_����Excel���[�N�u�b�N�̃t�@�C�������Q�Ƃ���(1����)
    strFileName = Dir(strPathName & "\*.xlsx", vbNormal)
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
    On Error GoTo Error_Handler
    
    Call LogStart("���ʌ^���o����", "")
    ' �e���v���[�g���[�N�u�b�N���J��
    Set templateWBK = OpenWorkBook(TEMPLATE_FILE_PATH, False, True)
    
    If Not templateWBK Is Nothing Then

        Set targetFiles = New Collection
        targetFiles.Add (strFileName)
        ' �w��t�H���_�̑SExcel���[�N�u�b�N�ɂ��ČJ��Ԃ�
        Do While strFileName <> ""
            targetFiles.Add (strFileName)
            ' ���̃t�@�C�������Q��
            strFileName = Dir
        Loop
        
        ' �w��t�H���_�̑SExcel���[�N�u�b�N�ɂ��ČJ��Ԃ�
        Dim aFile As Variant
        For Each aFile In targetFiles
            ' Esc�L�[�Ō�����
            DoEvents
            If swESC = True Then
                ' ���f����̂������b�Z�[�W�Ŋm�F
                If MsgBox("���f�L�[��������܂����B�����ŏI�����܂����H", _
                    vbInformation + vbYesNo) = vbYes Then
                    GoTo Error_Handler
                Else
                    swESC = False
                End If
            End If
            '-----------------------------------------------------------------------
            ' ���������P�t�@�C���P�ʂ̏���
            Call WB���ʌ^���o����(xlAPP, templateWBK, strPathName, CStr(aFile))
            '-----------------------------------------------------------------------
    
        Next
    
        templateWBK.Close SaveChanges:=False
    
    End If

    GoTo Sub_EXIT
    
'----------------
' Esc�L�[�E�o�p�s���x��
Error_Handler:
    If Err.Number = 18 Then
        ' Esc�L�[�ł̃G���[Raise
        swESC = True
        Resume
    ElseIf Err.Number = 1004 Then
        ' �B���V�[�g�����ΏۂȂ��̎��s���G���[�͖���
        Resume Next
    Else
        Call LogErrorHandle(Err.Description, "���ʌ^���o����", "")
        ' ���̑��̃G���[�̓��b�Z�[�W�\����I��
        MsgBox Err.Description
    End If
    Resume Sub_EXIT
'----------------
' �����I��
Sub_EXIT:
    Call LogEnd("���ʌ^���o����", "")
    With xlAPP
        .StatusBar = False                  ' �X�e�[�^�X�o�[�𕜋A
        .EnableEvents = True                ' �C�x���g����ĊJ
        .EnableCancelKey = xlInterrupt      ' Esc�L�[�����߂�
        .Cursor = xlDefault                 ' �J�[�\������̫�Ăɂ���
        .ScreenUpdating = True              ' ��ʕ`��ĊJ
    End With
    Set xlAPP = Nothing
    MsgBox "�I�����܂���"
End Sub

'*******************************************************************************
' �P�̃��[�N�u�b�N�̈ڍs����
'*******************************************************************************
Private Sub WB���ʌ^���o����(xlAPP As Application, _
                            templateWBK As Workbook, _
                            strPathName As String, _
                            strFileName As String)
    
    On Error GoTo Error_Handler
    Call LogStart("WB���ʌ^���o����", strFileName)
    '---------------------------------------------------------------------------
    Dim objWBK As Workbook          ' ���[�N�u�b�NObject

    ' �X�e�[�^�X�o�[�ɏ����t�@�C������\��
    xlAPP.StatusBar = strFileName & " �������D�D�D"
    ' ���[�N�u�b�N���J��
    Set objWBK = OpenWorkBook(strPathName & cnsYEN & strFileName, False, False)
    
    If Not objWBK Is Nothing Then
    
        If Not ContainsSheet(objWBK, SHEET_NAME_COVER_NEW) Then
            objWBK.Close SaveChanges:=False
            GoTo Sub_EXIT
            Exit Sub
        End If
        
        '---------------------------------------------------------------------------
        ' �V�[�g���ɏ���
        '---------------------------------------------------------------------------
        Call �ڍs_���ʌ^��`(objWBK, strPathName)
        
        objWBK.Close SaveChanges:=True

    End If
    
    GoTo Sub_EXIT
    
Error_Handler:
    Call LogErrorHandle(Err.Description, "WB���ʌ^���o����", strFileName)
    If Not objWBK Is Nothing Then
        objWBK.Close SaveChanges:=False
    End If

    '-- �G���[�����������ꍇ
    'Resume Sub_EXIT
    Err.Raise 5003
   
Sub_EXIT:
    Call LogEnd("WB���ʌ^���o����", strFileName)
    xlAPP.StatusBar = False
    Set objWBK = Nothing

End Sub

'*******************************************************************************
' ���t�H�[�}�b�g����V�t�H�[�}�b�g�ֈڍs����
'*******************************************************************************
Public Sub �ڍs()
    Call debugLog(INIT_LOG, 0)
    Dim xlAPP As Application        ' Excel.Application
    Dim strPathName As String       ' �w��t�H���_��
    Dim strFileName As String       ' ���o�����t�@�C����
    Dim swESC As Boolean            ' Esc�L�[����
    Dim templateWBK As Workbook     ' Excel.Workbook(�e���v���[�g�t�@�C��)
    Dim targetFiles As Collection   ' �Ώۃt�@�C���ꗗ
    
    ' ��t�H���_�̎Q�ƣ���t�H���_���̎擾(modFolderPicker1�Ɏ��e)
    strPathName = FolderDialog("�t�H���_���w�肵�ĉ�����", True)
    If strPathName = "" Then Exit Sub
    
    ' �w��t�H���_����Excel���[�N�u�b�N�̃t�@�C�������Q�Ƃ���(1����)
    strFileName = Dir(strPathName & "\*.xlsx", vbNormal)
    If strFileName = "" Then
        MsgBox "���̃t�H���_�ɂ�Excel���[�N�u�b�N�͑��݂��܂���B"
        Exit Sub
    End If

    Set xlAPP = Application
    With xlAPP
        .ScreenUpdating = False             ' ��ʕ`���~
        .EnableEvents = False               ' �C�x���g�����~
        .EnableCancelKey = xlErrorHandler   ' Esc�L�[�ŃG���[�g���b�v����
'        .Cursor = xlWait                    ' �J�[�\���������v�ɂ���
    End With
    On Error GoTo Error_Handler
 
    Call LogStart("�ڍs", strFileName)
    
    ' �e���v���[�g���[�N�u�b�N���J��
    Set templateWBK = OpenWorkBook(TEMPLATE_FILE_PATH, False, True)
    
    If Not templateWBK Is Nothing Then

        Set targetFiles = New Collection
        ' �w��t�H���_�̑SExcel���[�N�u�b�N�ɂ��ČJ��Ԃ�
        Do While strFileName <> ""
            targetFiles.Add (strFileName)
            ' ���̃t�@�C�������Q��
            strFileName = Dir
        Loop
        
        ' �w��t�H���_�̑SExcel���[�N�u�b�N�ɂ��ČJ��Ԃ�
        Dim aFile As Variant
        For Each aFile In targetFiles
            ' Esc�L�[�Ō�����
            DoEvents
            If swESC = True Then
                ' ���f����̂������b�Z�[�W�Ŋm�F
                If MsgBox("���f�L�[��������܂����B�����ŏI�����܂����H", _
                    vbInformation + vbYesNo) = vbYes Then
                    GoTo Error_Handler
                Else
                    swESC = False
                End If
            End If
            '-----------------------------------------------------------------------
            ' ���������P�t�@�C���P�ʂ̏���
            Call WB�ڍs(xlAPP, templateWBK, strPathName, CStr(aFile))
            '-----------------------------------------------------------------------
    
        Next
    
        templateWBK.Close SaveChanges:=False
    
    End If
    
    GoTo Sub_EXIT
    
'----------------
' Esc�L�[�E�o�p�s���x��
Error_Handler:
    If Err.Number = 18 Then
        ' Esc�L�[�ł̃G���[Raise
        swESC = True
        Resume
    ElseIf Err.Number = 1004 Then
        ' �B���V�[�g�����ΏۂȂ��̎��s���G���[�͖���
        Resume Next
    Else
        Call LogErrorHandle(Err.Description, "�ڍs", strFileName)
        ' ���̑��̃G���[�̓��b�Z�[�W�\����I��
        MsgBox Err.Description
    End If
    Resume Sub_EXIT
'----------------
' �����I��
Sub_EXIT:
    Call LogEnd("�ڍs", strFileName)
    With xlAPP
        .StatusBar = False                  ' �X�e�[�^�X�o�[�𕜋A
        .EnableEvents = True                ' �C�x���g����ĊJ
        .EnableCancelKey = xlInterrupt      ' Esc�L�[�����߂�
        .Cursor = xlDefault                 ' �J�[�\������̫�Ăɂ���
        .ScreenUpdating = True              ' ��ʕ`��ĊJ
    End With
    Set xlAPP = Nothing
    MsgBox "�I�����܂���"
End Sub

'*******************************************************************************
' �P�̃��[�N�u�b�N�̈ڍs����
'*******************************************************************************
Private Sub WB�ڍs(xlAPP As Application, _
                            templateWBK As Workbook, _
                            strPathName As String, _
                            strFileName As String)
    
    On Error GoTo Error_Handler
    '---------------------------------------------------------------------------
    Dim objWBK As Workbook          ' ���[�N�u�b�NObject

    Call LogStart("WB�ڍs", strFileName)
    
    ' �X�e�[�^�X�o�[�ɏ����t�@�C������\��
    xlAPP.StatusBar = strFileName & " �������D�D�D"
    ' ���[�N�u�b�N���J��
    Set objWBK = OpenWorkBook(strPathName & cnsYEN & strFileName, False, False)
    
    If Not objWBK Is Nothing Then
        
        If Not ContainsSheet(objWBK, SHEET_NAME_INOUT_OLD) Then
            objWBK.Close SaveChanges:=False
            GoTo Sub_EXIT
            Exit Sub
        End If
        
        '---------------------------------------------------------------------------
        ' �V�[�g���Ɉڍs
        '---------------------------------------------------------------------------
        Call �ڍs_�\��(objWBK, templateWBK)
        Call �ڍs_��������(objWBK, templateWBK)
        Call �ڍs_�Q�ƃt�@�C���ꗗ(objWBK, templateWBK)
        Call �ڍs_���ʏ��(objWBK, templateWBK)
        Call �ڍs_�^�ꗗ(objWBK, templateWBK)
        Call �ڍs_���o�͒�`(objWBK, templateWBK)
        Call �ڍs_�^�V�[�g(objWBK, templateWBK)
        
        '---------------------------------------------------------------------------
        ' �^�ꗗ�̐���
        '---------------------------------------------------------------------------
        Call Do�v�f�ꗗ�쐬����(objWBK, TARGET_SHEET_PREFIX_NEW)
            
        Call �ڍs_���ʌ^��`(objWBK, strPathName)
        '---------------------------------------------------------------------------
        ' �W��V�[�g�̍폜
        '---------------------------------------------------------------------------
        Call DiscardSheet(GetSummarySheet(objWBK))
    
        
        objWBK.Close SaveChanges:=True

    End If
    
    GoTo Sub_EXIT
    
Error_Handler:
    Call LogErrorHandle(Err.Description, "WB�ڍs", strFileName)
    
    If Not objWBK Is Nothing Then
        objWBK.Close SaveChanges:=False
    End If

    '-- �G���[�����������ꍇ
    'Resume Sub_EXIT
    Err.Raise 5003
   
Sub_EXIT:
    Call LogEnd("WB�ڍs", strFileName)
    xlAPP.StatusBar = False
    Set objWBK = Nothing

End Sub

'*******************************************************************************
' �u�\���v�̈ڍs����
'*******************************************************************************
Private Sub �ڍs_�\��(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("�ڍs_�\��", targetWBK.Name)
    
    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_COVER_NEW
    Dim OLD_SHEET_NAME As String: OLD_SHEET_NAME = SHEET_NAME_COVER_OLD
    
    '-------------------
    '���ڍs�v�ۃ`�F�b�N
    '-------------------
    '�V�[�g�����݂��Ă��邩
    'B21�̃Z���ɒl�����邩�ǂ���
    Dim targetWS As Worksheet
    Set targetWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    If (Not targetWS Is Nothing) Then
        If (targetWS.Range("B21").Value <> "") Then
            Exit Sub
        End If
    End If
    
    '-------------------
    '���ڍs���s
    '-------------------
    Dim templateWS As Worksheet, newWS As Worksheet
    Set templateWS = GetSheet(templateWBK, NEW_SHEET_NAME)
    Set targetWS = GetSheet(targetWBK, OLD_SHEET_NAME)
    '�ΏۃV�[�g���Ȃ��ꍇ��old�V�[�g��T��
    If targetWS Is Nothing Then
        Set targetWS = GetSheet(targetWBK, "old_" & OLD_SHEET_NAME)
        If targetWS Is Nothing Then
            MsgBox "�ڍs���V�[�g������܂���"
            Err.Raise 6001
        End If
    End If
    
    '���l�[��
    If targetWS.Name = OLD_SHEET_NAME Then
        targetWS.Name = "old_" & targetWS.Name
    End If
    
    '�V�V�[�g�����݂��Ă�����폜
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    If Not newWS Is Nothing Then
        Application.DisplayAlerts = False
        newWS.Delete
        Application.DisplayAlerts = True
    End If

    '�e���v���[�g����R�s�[
    templateWS.Copy Before:=Sheets(1)
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    '�l�̈ڍs
    newWS.Range("B21").Value = "=CONCATENATE(""��"",TEXT(MAX(��������!A2:A1048576),""0.00""),""�� "")"
    newWS.Range("B23").Value = "=CONCATENATE(YEAR(MAX(��������!B2:B1048576)),""."",MONTH(MAX(��������!B2:B1048576)),""."",DAY(MAX(��������!B2:B1048576)))"
    newWS.Range("B16").Value = targetWS.Range("B16").Value
    newWS.Range("E16").Value = targetWS.Range("E16").Value

    '---------------------------------------------------------------------------
    ' �V�[�g�j��
    '---------------------------------------------------------------------------
    Call DiscardSheet(targetWS)
    
    Call LogEnd("�ڍs_�\��", targetWBK.Name)

End Sub

'*******************************************************************************
' �u���������v�̈ڍs����
'*******************************************************************************
Private Sub �ڍs_��������(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("�ڍs_��������", targetWBK.Name)

    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_HISTORY_NEW
    Dim OLD_SHEET_NAME As String: OLD_SHEET_NAME = SHEET_NAME_HISTORY_OLD
    
    '-------------------
    '���ڍs�v�ۃ`�F�b�N
    '-------------------
    '�����Ȃ�
    
    '-------------------
    '���ڍs���s
    '-------------------
    Dim targetWS As Worksheet
    Set targetWS = GetSheet(targetWBK, OLD_SHEET_NAME)
    '�ΏۃV�[�g���Ȃ��ꍇ�̓G���[
    If targetWS Is Nothing Then
        MsgBox "�ڍs���V�[�g������܂���"
        Err.Raise 6002
    End If
    
    '�����ύX�̂�
    targetWS.Range("A2:A30").NumberFormatLocal = "G/�W��"
    
    Call LogEnd("�ڍs_��������", targetWBK.Name)

       
End Sub

'*******************************************************************************
' �u�Q�ƃt�@�C���ꗗ�v�̈ڍs����
'*******************************************************************************
Private Sub �ڍs_�Q�ƃt�@�C���ꗗ(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("�ڍs_�Q�ƃt�@�C���ꗗ", targetWBK.Name)


    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_REF_LIST_NEW '"1.�Q�ƃt�@�C���ꗗ"
    Dim newWS As Worksheet
    
    '-------------------
    '���ڍs�v�ۃ`�F�b�N
    '-------------------
    '�u1.�Q�ƃt�@�C���ꗗ�v�̃V�[�g������Ύ��s���Ȃ�
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    '�ΏۃV�[�g���Ȃ��ꍇ�ɃR�s�[
    If Not newWS Is Nothing Then
        Exit Sub
    End If

    '-------------------
    '���ڍs���s
    '-------------------
    '�e���v���[�g����R�s�[
    templateWBK.Worksheets(NEW_SHEET_NAME).Copy After:=GetSheet(targetWBK, SHEET_NAME_HISTORY_NEW)
    Set newWS = targetWBK.Worksheets(NEW_SHEET_NAME)
    '---------------------------------------------------------------------------
    ' �s�v�R�����g�폜
    '---------------------------------------------------------------------------
    newWS.Range("A1:D30").ClearComments
    '---------------------------------------------------------------------------
    ' �T���v���폜
    '---------------------------------------------------------------------------
    newWS.Range("A2:D2").Value = ""

    '---------------------------------------------------------------------------
    ' �s�̍����𒲐�����B
    '---------------------------------------------------------------------------
    Call AdjustRowHeight(newWS)

    Call LogEnd("�ڍs_�Q�ƃt�@�C���ꗗ", targetWBK.Name)

End Sub

'*******************************************************************************
' �u���ʏ��v�̈ڍs����
'*******************************************************************************
Private Sub �ڍs_���ʏ��(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("�ڍs_���ʏ��", targetWBK.Name)

    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_COMMON_NEW '"2.���ʏ��"
    Dim newWS As Worksheet
    
    '-------------------
    '���ڍs�v�ۃ`�F�b�N
    '-------------------
    '�u2.���ʏ��v�̃V�[�g������Ύ��s���Ȃ�
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    '�ΏۃV�[�g���Ȃ��ꍇ�ɃR�s�[
    If Not newWS Is Nothing Then
        Exit Sub
    End If

    '-------------------
    '���ڍs���s
    '-------------------
    '�e���v���[�g����R�s�[
    templateWBK.Worksheets(NEW_SHEET_NAME).Copy After:=targetWBK.Worksheets(SHEET_NAME_REF_LIST_NEW) '�u1.�Q�ƃt�@�C���ꗗ�v�̌��
    Set newWS = targetWBK.Worksheets(NEW_SHEET_NAME)
    '�폜
    newWS.Range("C2").Value = ""
    
    Call LogEnd("�ڍs_���ʏ��", targetWBK.Name)

End Sub

'*******************************************************************************
' �u�^�ꗗ�v�̈ڍs����
'*******************************************************************************
Private Sub �ڍs_�^�ꗗ(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("�ڍs_�^�ꗗ", targetWBK.Name)
    
    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_TYPE_LIST_NEW '"3.�^�ꗗ"
    Dim OLD_SHEET_NAME As String: OLD_SHEET_NAME = SHEET_NAME_TYPE_LIST_OLD '"2.�v�f�ꗗ"

    Dim newWS As Worksheet, oldWS As Worksheet
    
    '-------------------
    '���ڍs�v�ۃ`�F�b�N
    '-------------------
    '�u3.�^�ꗗ�v�̃V�[�g������Ύ��s���Ȃ�
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    '�ΏۃV�[�g���Ȃ��ꍇ�ɃR�s�[
    If Not newWS Is Nothing Then
        Exit Sub
    End If
    
    '-------------------
    '���ڍs���s
    '-------------------
    '�e���v���[�g����R�s�[
    templateWBK.Worksheets(NEW_SHEET_NAME).Copy After:=targetWBK.Worksheets(SHEET_NAME_COMMON_NEW) '�u2.���ʏ��v�̌��
    Set newWS = targetWBK.Worksheets(NEW_SHEET_NAME)

    '---------------------------------------------------------------------------
    ' �v�f���X�g�̃G���g���[���폜
    '---------------------------------------------------------------------------
    Dim i As Integer
    For i = newWS.Cells(newWS.Rows.count, 1).End(xlUp).row To 2 Step -1
        newWS.Rows(i).Delete
    Next
    
    '---------------------------------------------------------------------------
    ' �s�v�R�����g�폜
    '---------------------------------------------------------------------------
    newWS.Range("B1:F6").ClearComments
                    
    '---------------------------------------------------------------------------
    ' �s�̍����𒲐�����B
    '---------------------------------------------------------------------------
    Call AdjustRowHeight(newWS)
                    
    '---------------------------------------------------------------------------
    ' �V�[�g�j��
    '---------------------------------------------------------------------------
    Call DiscardSheet(GetSheet(targetWBK, OLD_SHEET_NAME))
    
    Call LogEnd("�ڍs_�^�ꗗ", targetWBK.Name)
    
End Sub

'*******************************************************************************
' �ΏۃV�[�g�̍s�̍����𒲐�����B
'*******************************************************************************
Private Sub AdjustRowHeight(aSheet As Worksheet)
    Dim cmd As CmdRowAdjust
    Set cmd = New CmdRowAdjust
    Call cmd.ExecCommand(aSheet)
End Sub

'*******************************************************************************
' �u���o�͒�`�v�̈ڍs����
'*******************************************************************************
Private Sub �ڍs_���o�͒�`(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("�ڍs_���o�͒�`", targetWBK.Name)
    
    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_INOUT_NEW  '"4.���o�͒�`"
    Dim OLD_SHEET_NAME As String: OLD_SHEET_NAME = SHEET_NAME_INOUT_OLD  '"1.���o�͒�`"
    Dim newWS As Worksheet, oldWS As Worksheet
    
    '-------------------
    '���ڍs�v�ۃ`�F�b�N
    '-------------------
    '�ΏۃV�[�g�����ɂ���ꍇ�͏I��
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    If Not newWS Is Nothing Then
        Exit Sub
    End If
    
    '���ƂȂ�V�[�g���Ȃ��ꍇ�͏I��
    Set oldWS = GetSheet(targetWBK, OLD_SHEET_NAME)
    If oldWS Is Nothing Then
        Exit Sub
    End If
    
    '-------------------
    '���ڍs���s
    '-------------------
    '�e���v���[�g����R�s�[
    templateWBK.Worksheets(NEW_SHEET_NAME).Copy After:=targetWBK.Worksheets(SHEET_NAME_TYPE_LIST_NEW) '�u3.�^�ꗗ�v�̌��
    '�V�V�[�g���擾
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)

    '---------------------------------------------------------------------------
    ' �V�V�[�g�̓��o�͒�`�̃G���g���[���폜
    '---------------------------------------------------------------------------
    Dim i As Integer
    For i = newWS.Cells(newWS.Rows.count, 1).End(xlUp).row To 2 Step -1
        newWS.Rows(i).Delete
    Next
    
    Dim j As Integer
    For j = 2 To oldWS.Cells(oldWS.Rows.count, 1).End(xlUp).row
        newWS.Cells(j, COL_NEW_IO_NO).Value = "=ROW()-1"
        newWS.Cells(j, COL_NEW_IO_IO).Value = oldWS.Cells(j, COL_OLD_IO_IO).Value
        newWS.Cells(j, COL_NEW_IO_ROOT_ELEMENT_ID).Value = oldWS.Cells(j, COL_OLD_IO_ROOT_ELEMENT_ID).Value
        newWS.Cells(j, COL_NEW_IO_ROOT_ELEMENT_NAME).Value = oldWS.Cells(j, COL_OLD_IO_ELEMENT_NAME).Value
        newWS.Cells(j, COL_NEW_IO_TYPE_NAME).Value = oldWS.Cells(j, COL_OLD_IO_ELEMENT_NAME).Value
        With newWS
          .Range(.Cells(i, COL_NEW_IO_MIN_NO), .Cells(i, COL_NEW_IO_MAX_NO)).Borders.LineStyle = xlContinuous
        End With
    Next

    '---------------------------------------------------------------------------
    ' �V�[�g�j��
    '---------------------------------------------------------------------------
    Call DiscardSheet(oldWS)
    
    Call LogEnd("�ڍs_���o�͒�`", targetWBK.Name)

End Sub

'*******************************************************************************
' �u�^�V�[�g�v�̈ڍs����
'*******************************************************************************
Private Sub �ڍs_�^�V�[�g(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("�ڍs_�^�V�[�g", targetWBK.Name)
    
    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_REF_LIST_NEW '"1.�Q�ƃt�@�C���ꗗ"
    
    '-------------------
    '���ڍs�v�ۃ`�F�b�N
    '-------------------
    '���̖��̂̃V�[�g�����݂���ꍇ�Ɏ��{����B
    Dim exists As Boolean: exists = False
    Dim aSheet As Worksheet
    For Each aSheet In targetWBK.Sheets
        If InStr(aSheet.Name, TARGET_SHEET_PREFIX_OLD) = 1 Then
            exists = True
            Exit For
        End If
    Next
    
    If Not exists Then
        Exit Sub
    End If
    
    '-------------------
    '���ڍs���s
    '-------------------
    Dim newWS As Worksheet, oldWS As Worksheet
    
    Dim pos As Integer, typeName As String, newSheetName As String
    Dim discardSheets As Collection
    Set discardSheets = New Collection
    
    '���̖��̂̃V�[�g����������
    For Each oldWS In targetWBK.Sheets
        If InStr(oldWS.Name, TARGET_SHEET_PREFIX_OLD) = 1 And oldWS.Name <> SHEET_NAME_INOUT_NEW Then
            '�Ō��.(�h�b�g�̈ʒu)
            pos = InStrRev(oldWS.Name, ".")
            If pos > 0 Then
                ' �v�f�ꗗ�V�[�g�ɒǋL
                typeName = Mid(oldWS.Name, pos + 1)
                newSheetName = TARGET_SHEET_PREFIX_NEW & typeName
                ' ���݂��Ȃ��ꍇ�Ƀe���v���[�g����쐬����B
                If Not ContainsSheet(targetWBK, newSheetName) Then
                    ' �R�s�[���s
                    templateWBK.Worksheets(TEMPLATE_SHEET_NAME).Copy After:=targetWBK.Worksheets(targetWBK.Worksheets.count)
                    targetWBK.Worksheets(targetWBK.Worksheets.count).Name = newSheetName
                    Set newWS = targetWBK.Sheets(newSheetName)

                    '---------------------------------------------------------------------------
                    ' �T���v���폜
                    '---------------------------------------------------------------------------
                    With newWS.Range("A2:P30")
                        .Value = ""             '�l�폜
                        .ClearComments          '�R�����g�폜
                    End With
  
                    '---------------------------------------------------------------------------
                    ' �f�[�^�ڍs
                    '---------------------------------------------------------------------------
                    Dim j As Integer
                    For j = 2 To oldWS.Cells(oldWS.Rows.count, 1).End(xlUp).row
                        
                        newWS.Cells(j, COL_NEW_TYPEDEF_NO).Value = "=ROW()-1"    '����
                        newWS.Cells(j, COL_NEW_TYPEDEF_NAME).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_NAME).Value   '���ږ�
                        newWS.Cells(j, COL_NEW_TYPEDEF_ID).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_ID).Value   '����ID
                        newWS.Cells(j, COL_NEW_TYPEDEF_EXPLAIN).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_EXPLAIN).Value   '����
                        newWS.Cells(j, COL_NEW_TYPEDEF_MIN_LOOP).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_MIN_LOOP).Value   '�ŏ��J��Ԃ�
                        newWS.Cells(j, COL_NEW_TYPEDEF_MIN_LOOP).HorizontalAlignment = xlRight
                        '�ő�J��Ԃ�
                        If oldWS.Cells(j, COL_OLD_TYPEDEF_MAX_LOOP).Value = "n" Or oldWS.Cells(j, COL_OLD_TYPEDEF_MAX_LOOP).Value = "��" Then
                            newWS.Cells(j, COL_NEW_TYPEDEF_MAX_LOOP).Value = "*"
                        Else
                            newWS.Cells(j, COL_NEW_TYPEDEF_MAX_LOOP).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_MAX_LOOP).Value
                        End If
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAX_LOOP).HorizontalAlignment = xlRight
                        
                        
                        newWS.Cells(j, COL_NEW_TYPEDEF_COND_REQ).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_COND_REQ).Value   '�K�{����
                        
                        '�^
                        If oldWS.Cells(j, COL_OLD_TYPEDEF_TYPE).Value = "���l" Then
                            newWS.Cells(j, COL_NEW_TYPEDEF_TYPE).Value = "����"
                        Else
                            newWS.Cells(j, COL_NEW_TYPEDEF_TYPE).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_TYPE).Value
                        End If
                        
                        '�^��
                        If oldWS.Cells(j, COL_OLD_TYPEDEF_TYPE).Value = "�v�f" Then
                            newWS.Cells(j, COL_NEW_TYPEDEF_TYPE_NAME).Value = oldWS.Cells(j, 2).Value
                        Else
                            newWS.Cells(j, COL_NEW_TYPEDEF_TYPE_NAME).Value = ""
                        End If
                        
                        newWS.Cells(j, COL_NEW_TYPEDEF_CHAR_FORMAT).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_TYPE_DETAIL).Value  '������E����
                        newWS.Cells(j, COL_NEW_TYPEDEF_MINLENGTH).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_MINLENGTH).Value  '�ŏ�����
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAXLENGTH).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_MAXLENGTH).Value  '�ő包��
                        newWS.Cells(j, COL_NEW_TYPEDEF_DATA_SAMPLE).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_DATA_SAMPLE).Value  '�f�[�^��
                        newWS.Cells(j, COL_NEW_TYPEDEF_BIKO).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_BIKO).Value  '���l
                        '���O��4����x�R�s�[
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAX_NO + 1).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_BIKO + 1).Value
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAX_NO + 2).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_BIKO + 2).Value
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAX_NO + 3).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_BIKO + 3).Value
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAX_NO + 4).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_BIKO + 4).Value
                        
                        '�r��
                        With newWS
                            .Range(.Cells(j, COL_NEW_TYPEDEF_MIN_NO), .Cells(j, COL_NEW_TYPEDEF_MAX_NO)).Borders.LineStyle = xlContinuous
                        End With
                    Next
                    '---------------------------------------------------------------------------
                    ' �s�v�ȍs�̍폜
                    '---------------------------------------------------------------------------
                    If newWS.Cells(newWS.Rows.count, 1).End(xlUp).row < TEMPLATE_TYPE_DEF_SHEET_MAX_ROW Then
                        For j = TEMPLATE_TYPE_LIST_SHEET_MAX_ROW To newWS.Cells(newWS.Rows.count, 1).End(xlUp).row + 1 Step -1
                            newWS.Rows(j).Delete
                        Next
                    End If
                    
                    Dim cmd As CmdRowAdjust
                    Set cmd = New CmdRowAdjust
                    Call cmd.ExecCommand(newWS)
                    Call discardSheets.Add(oldWS)
                    
                End If
            End If
        End If
    Next
    '
    '---------------------------------------------------------------------------
    ' ���V�[�g�̔j��
    '---------------------------------------------------------------------------
    For Each oldWS In discardSheets
        Call DiscardSheet(oldWS)
    Next
    
    Call LogEnd("�ڍs_�^�V�[�g", targetWBK.Name)
    
End Sub

'*******************************************************************************
' ���V�[�g�̔j������
'*******************************************************************************
Private Sub DiscardSheet(oldWS As Worksheet)
    If Not oldWS Is Nothing Then
        If DESTRUCTION_MODE = DESTRUCTION_MODE_DELETE Then
            '�V�[�g�폜
            Application.DisplayAlerts = False
            oldWS.Delete
            Application.DisplayAlerts = True
            
        ElseIf DESTRUCTION_MODE = DESTRUCTION_MODE_RENAME Then
            '�V�[�g����ύX
            oldWS.Name = "old_" & oldWS.Name
        End If
    End If
End Sub

'*******************************************************************************
' ���ʌ^��`�̈ڍs����
'*******************************************************************************
Private Function GetCommonTypes() As Collection
    
    Dim defSheet As Worksheet, typeName As String, typeId As String
    Set defSheet = ThisWorkbook.Worksheets(SHEET_NAME_TYPEDEF)
    
    Set GetCommonTypes = New Collection
    
    Dim i As Integer
    For i = 3 To defSheet.Cells(defSheet.Rows.count, 3).End(xlUp).row
        typeId = defSheet.Cells(i, 2)
        typeName = defSheet.Cells(i, 3)
        If typeId <> "" And typeName <> "" Then
            GetCommonTypes.Add (typeName & ":" & typeId)
        End If
    Next
End Function

'*******************************************************************************
' ���ʌ^��`�̈ڍs����
'*******************************************************************************
Private Sub �ڍs_���ʌ^��`(targetWBK As Workbook, strPathName As String)

    Call LogStart("�ڍs_���ʌ^��`", targetWBK.Name)
    
    Dim CommonTypes As Collection
    Set CommonTypes = GetCommonTypes()
    
    Dim aSheet As Worksheet
    ' �ꗗ�쐬�p�Q�ƌ^��`���X�g
    Dim refTypeNames As Collection
    Set refTypeNames = New Collection
    
    Dim changed As Boolean: changed = False

    '-------------------
    '���ڍs���s
    '-------------------

    '---------------------------------------------------------------------------
    '�@�e�^��`�V�[�g���̋��ʌ^�̕����Ƀt�@�C��ID��ݒ肷�鏈��
    '---------------------------------------------------------------------------
    Call Do�Q�ƌ^��`���o����(targetWBK, refTypeNames, CommonTypes, strPathName)
    
    If refTypeNames.count > 0 Then
        '---------------------------------------------------------------------------
        '�@�Q�ƃt�@�C���ꗗ�̍쐬����
        '---------------------------------------------------------------------------
        Call Do�Q�ƃt�@�C���ꗗ�쐬����(targetWBK, refTypeNames)
          
        '---------------------------------------------------------------------------
        '�@���ʌ^�ƂȂ�V�[�g��ʖ��ۑ����鏈��
        '---------------------------------------------------------------------------
        Call Do���ʌ^�ʖ��ۑ�����(targetWBK, refTypeNames, strPathName)
    
        '---------------------------------------------------------------------------
        '�@ID�Ⴂ�̃t�@�C��������ꍇ�Ƀ��l�[�����鏈��
        '---------------------------------------------------------------------------
        Call Do���ʌ^���l�[������(targetWBK, refTypeNames, strPathName)
        
    End If
    
    
    Call LogEnd("�ڍs_���ʌ^��`", targetWBK.Name)
    
        
End Sub

'*******************************************************************************
' �Q�ƌ^��`���o����
'*******************************************************************************
Sub Do�Q�ƌ^��`���o����(targetWBK As Workbook, refTypeNames As Collection, CommonTypes As Collection, strPathName As String)
    Dim aSheet As Worksheet
    If targetWBK Is Nothing Then
        Exit Sub
    End If
    For Each aSheet In targetWBK.Sheets

        If InStr(aSheet.Name, TARGET_SHEET_PREFIX_NEW) = 1 Then
            '---------------------------------------------------------------------------
            ' �s���̏���
            '---------------------------------------------------------------------------
            Dim j As Integer
            Dim refTypeName As String, itm As Variant, bufWb As Workbook
            For j = 2 To aSheet.Cells(aSheet.Rows.count, 1).End(xlUp).row
                refTypeName = aSheet.Cells(j, COL_NEW_TYPEDEF_TYPE_NAME)
                If Len(refTypeName) > 0 Then
                    Dim exists As Boolean: exists = False
                    
                    For Each itm In CommonTypes
                        '��v������̂������ID���L������B
                        If InStr(CStr(itm), refTypeName & ":") = 1 Then
                            exists = True
                            '�Z���̒l���قȂ�ꍇ�̂ݍX�V
                            Call UpdateCell(aSheet, j, COL_NEW_TYPEDEF_FILE_ID, Mid(itm, Len(refTypeName) + 2))

                            If ContainsItem(refTypeNames, CStr(itm)) = False Then
                                Call refTypeNames.Add(itm)
                                If GetSheet(targetWBK, TARGET_SHEET_PREFIX_NEW & refTypeName) Is Nothing Then
                                    Set bufWb = OpenWorkBook(strPathName & cnsYEN & GenCommonTypeFileName(CStr(itm)), False, True, False)
                                    If Not bufWb Is Nothing Then
                                        Call Do�Q�ƌ^��`���o����(bufWb, refTypeNames, CommonTypes, strPathName)
                                        bufWb.Close SaveChanges:=False
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If Not exists Then
                        '�Z���̒l���قȂ�ꍇ�̂ݍX�V
                        Call UpdateCell(aSheet, j, COL_NEW_TYPEDEF_FILE_ID, "")
                    End If

                    If aSheet.Cells(j, COL_NEW_TYPEDEF_FILE_ID) = "" Then
                        If GetSheet(targetWBK, TARGET_SHEET_PREFIX_NEW & refTypeName) Is Nothing Then
                            Call LogErrorHandle("�V�[�g�u" & aSheet.Name & "�v�ɋL�ڂ̌^�u" & refTypeName & "�v�͑Ώۂ̃V�[�g�����݂��܂���B", "", "")
                        End If
                    End If
                End If
            Next
            
            Dim typeWb As Workbook
            For Each itm In CommonTypes
                If InStr(CStr(itm), Mid(aSheet.Name, 5) & ":") = 1 Then
                    If ContainsItem(refTypeNames, CStr(itm)) = False Then
                        Call refTypeNames.Add(itm)
                    End If
                End If
            Next
        End If
    
    Next

End Sub

'*******************************************************************************
' �X�V����
'*******************************************************************************
Sub UpdateCell(aSheet As Worksheet, rowId As Integer, colId As Integer, val As Variant)
    If CStr(aSheet.Cells(rowId, colId)) <> CStr(val) Then
        aSheet.Cells(rowId, colId) = val
    End If
End Sub


'*******************************************************************************
' �R���N�V�������ɃA�C�e�����܂܂�邩�ǂ���
'*******************************************************************************
Private Function ContainsItem(list As Collection, item As String) As Boolean
    Dim t As Variant, flag As Boolean: flag = False
    
    For Each t In list
        If CStr(t) = item Then
            flag = True
            Exit For
        End If
    Next
    ContainsItem = flag
End Function

'*******************************************************************************
' ���ʌ^�̃V�[�g��ʖ��ۑ�����
'*******************************************************************************
Private Sub Do���ʌ^�ʖ��ۑ�����(objWBK As Workbook, list As Collection, strPathName As String)
    '---------------------------------------------------------------------------
    Dim item As Variant
    
    For Each item In list
        Call �ʖ��ۑ�����( _
            objWBK, _
            GetSheet(objWBK, TARGET_SHEET_PREFIX_NEW & Mid(CStr(item), 1, InStr(CStr(item), ":") - 1)), _
            GenCommonTypeFileName(CStr(item)), _
            strPathName)
    Next
    
End Sub

'*******************************************************************************
' ���ʌ^�̃V�[�g��ʖ��ۑ�����
'*******************************************************************************
Private Sub �ʖ��ۑ�����(objWBK As Workbook, tgtSheet As Worksheet, fileName As String, strPathName As String)

    Call LogStart("�ʖ��ۑ�����", objWBK.Name & ":" & fileName)
    
    '�Q�ƌ^��`�ꗗ�̌^����������Book�ɑ��݂����ꍇ�ɁA�ʖ��ۑ����������{����B
    If tgtSheet Is Nothing Then
        Exit Sub
    End If
        
    '�p�X���擾���� ���Q
    Dim filePath As String
    Dim testFileName As String: testFileName = fileName
    Dim pos As Integer
    Dim counter As Integer: counter = 1
    Dim newWorkBook As String
    
    filePath = strPathName & "\" & testFileName
    
    '���Ƀt�@�C�������݂��Ă����珈�����s��Ȃ��B
    '�u�b�N�����擾
    '�Ō��_(�A���_�[�X�R�A�̈ʒu)
    Dim bookName As String
    pos = InStrRev(objWBK.Name, "_")
    If pos > 0 Then
        bookName = Mid(objWBK.Name, pos + 1)
    End If
        
    If Not Dir(filePath) <> "" Then
        'Workbooks.Add
        'newWorkBook = ActiveWorkbook.Name
        'tgtSheet.Copy Before:=Workbooks(newWorkBook).Sheets(1)
        tgtSheet.Copy
        ActiveWorkbook.SaveAs fileName:=filePath        '�ʖ���t���ău�b�N��ۑ�����
        ActiveWorkbook.Close                            '�ʖ��u�b�N�����
        Call debugLog(bookName & ":���ʃV�[�g�R�s�[�쐬�F" & filePath, OUTPUT_LOG)
    End If

    'book�ŗL�t�@�C�����쐬
    Dim filePath2 As String
    pos = InStrRev(fileName, ".")
    filePath2 = strPathName & "\" & Mid(fileName, 1, pos - 1) & "_" & bookName
    tgtSheet.Move
    ActiveWorkbook.SaveAs fileName:=filePath2
    ActiveWorkbook.Close
    Call debugLog(bookName & ":���ʃV�[�g�쐬�F" & filePath2, OUTPUT_LOG)

    Call LogEnd("�ʖ��ۑ�����", objWBK.Name & ":" & fileName)
    
End Sub

'*******************************************************************************
' ���ʌ^��ID�Ⴂ�̃t�@�C��������΃��l�[������B
'*******************************************************************************
Private Sub Do���ʌ^���l�[������(objWBK As Workbook, CommonTypes As Collection, strPathName As String)

    Dim buf As String, item As Variant, searchPath As String
    For Each item In CommonTypes
        searchPath = "���V�X�e���C���^�[�t�F�[�X�d�l��_IFASW*" & _
                            "I_" & _
                            Mid(item, 1, InStr(item, ":") - 1) & ".xlsx"
                            
        buf = Dir(strPathName & cnsYEN & searchPath, vbNormal)
        If buf <> "" Then
            If buf <> GenCommonTypeFileName(CStr(item)) Then
                Dim buf2 As String
                buf2 = Dir(strPathName & cnsYEN & GenCommonTypeFileName(CStr(item)))
                If buf2 = "" Then
                    Call debugLog("���l�[��:�u" & buf & "�v=>�u" & GenCommonTypeFileName(CStr(item)) & "�v", OUTPUT_LOG)
                    '���l�[��
                    Name strPathName & cnsYEN & buf As strPathName & cnsYEN & GenCommonTypeFileName(CStr(item))
                Else
                    Call debugLog(buf & "�͕s�v�ȃt�@�C���ł��B" & buf2 & "�����݂��܂��B", OUTPUT_LOG)
                End If
            End If
        End If
CONTINUE:
    Next
End Sub

'*******************************************************************************
' �w��̃��[�N�u�b�N���I�[�v������B
'*******************************************************************************
Function OpenWorkBook(filePath As String, updateLinks_ As Boolean, readOnly_ As Boolean, Optional msgFlg As Boolean) As Workbook
    If IsMissing(msgFlg) Then
        msgFlg = True
    End If
    Dim buf As String, wb As Workbook
    ''�t�@�C���̑��݃`�F�b�N
    buf = Dir(filePath)
    If buf = "" Then
        If msgFlg Then
            MsgBox filePath & vbCrLf & "�͑��݂��܂���", vbExclamation
        End If
        Exit Function
    End If
    ''�����u�b�N�̃`�F�b�N
    For Each wb In Workbooks
        If wb.Name = buf Then
            If msgFlg Then
                MsgBox buf & vbCrLf & "�͂��łɊJ���Ă��܂�", vbExclamation
            End If
            Exit Function
        End If
    Next wb
    
    Set OpenWorkBook = Workbooks.Open(fileName:=filePath, _
                                updateLinks:=updateLinks_, _
                                readOnly:=readOnly_)
End Function
