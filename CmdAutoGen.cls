VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmdAutoGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' �R�[�h��`���̓��e�����ɁA�݌v���̓���̃Z���ɒl��ݒ肷��
' DB��`���Q�Ƃ���
'*******************************************************************************

Option Explicit

'*******************************************************************************
' �萔��`
'*******************************************************************************

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

Private Const COL_NEW_TYPEDEF_TABLE_NAME = 21 ' �e�[�u����
Private Const COL_NEW_TYPEDEF_COLUMN_NAME = 22 ' ���ږ�


'�C���^�t�F�[�X�d�l��_�^��`�V�[�g�̃V�[�g���ړ���
Private Const TARGET_SHEET_PREFIX = "�^��`_"

' �R�[�h��`����
Private m_codeDefDic As CodeDictionary
' DB���ڒ�`����
Private m_dbDefDic As DBDictionary
'*******************************************************************************
' �R���X�g���N�^
'*******************************************************************************
Private Sub Class_Initialize()
    Set m_codeDefDic = New CodeDictionary
    Set m_dbDefDic = New DBDictionary
    Debug.Print ("CmdRowAdjust Constructor is called.")
End Sub

'*******************************************************************************
' �f�X�g���N�^
'*******************************************************************************
Private Sub Class_Terminate()
    Debug.Print ("CmdRowAdjust Destructor is called.")
End Sub

'*******************************************************************************
' ������
'*******************************************************************************
Public Sub ExecCommand(targetSheet As Worksheet)

    Dim defineCellValue As String
    Dim defineKey As String
    Dim i As Integer
    Dim crPos As Integer
    Dim coronPos As Integer
    Dim tail As String
    
    Dim sheetName As String
    
    sheetName = targetSheet.Name
    
    ' �������擾
    Dim defCont As String, codeCont As String, varchar2Size As Integer
    Dim colType As String, pkey As Boolean, notnull As Boolean, explain As String, sample As String
    '���[�h���擾
    Dim modeSheet As Worksheet, mode As String
    Set modeSheet = ThisWorkbook.Worksheets("DB���ڃ`�F�b�N")
    
    mode = modeSheet.Range("C14").Value
    
    If InStr(sheetName, TARGET_SHEET_PREFIX) = 1 Then
        For i = 2 To targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).row

            ' No��̒�`���Ȃ��Ȃ�܂Ń��[�v
            If targetSheet.Cells(i, COL_NEW_TYPEDEF_NO) = "" Then
                Exit For
            End If
            
            '����ID���󔒂ɂ���B
            targetSheet.Cells(i, COL_NEW_TYPEDEF_ID).Value = ""
            
            '�v�f�̏ꍇ��
            If targetSheet.Cells(i, COL_NEW_TYPEDEF_TYPE) = "�v�f" Then
                GoTo CONTINUE
            End If
            
            '���o���̋L�q������ꍇ
            defCont = ""
            codeCont = ""
            If targetSheet.Cells(i, COL_NEW_TYPEDEF_TABLE_NAME) <> "" And targetSheet.Cells(i, COL_NEW_TYPEDEF_COLUMN_NAME) <> "" Then
                defCont = m_dbDefDic.GetValue(targetSheet.Cells(i, COL_NEW_TYPEDEF_TABLE_NAME), targetSheet.Cells(i, COL_NEW_TYPEDEF_COLUMN_NAME))
                codeCont = m_codeDefDic.GetValue(targetSheet.Cells(i, COL_NEW_TYPEDEF_TABLE_NAME), targetSheet.Cells(i, COL_NEW_TYPEDEF_COLUMN_NAME))
            Else
                defCont = m_dbDefDic.GetValue("", targetSheet.Cells(i, COL_NEW_TYPEDEF_NAME))
            End If
            
            
            ' �G���[����Ȃ��ꍇ
            If defCont <> "" And InStr(defCont, "ERR:") <> 1 Then
            
                Dim tmp As Variant
                tmp = Split(defCont, "@")
                colType = tmp(0)
                If tmp(1) = "��" Then
                    pkey = True
                Else
                    pkey = False
                End If
                
                If tmp(3) = "��" Then
                    notnull = True
                Else
                    notnull = False
                End If
                explain = tmp(4)
                sample = tmp(5)
                
                If notnull Then
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MIN_LOOP, 1)
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAX_LOOP, 1)
                Else
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MIN_LOOP, 0)
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAX_LOOP, 1)
                End If

                Dim dbType As String  'DB�̃f�[�^�^
                Dim dbLength As String 'DB���ڂ̒���
                
                Dim sizeOpenPos As Integer: sizeOpenPos = InStr(colType, "(")
                '�S�p�ɑΉ�
                If sizeOpenPos = 0 Then
                    sizeOpenPos = InStr(colType, "�i")
                End If
                
                Dim sizeClosePos As Integer: sizeClosePos = InStr(colType, ")")
                '�S�p�ɑΉ�
                If sizeClosePos = 0 Then
                    sizeClosePos = InStr(colType, "�j")
                End If
                '-------
                ' �⑫
                ' �uCHAR(8)�v�Ƃ����悤�ȕ\�L����A"CHAR" �Ƃ����^���� 8�Ƃ������������߂鏈��
                '-------
                If sizeOpenPos > 0 Then
                    dbType = StrConv(Mid(colType, 1, sizeOpenPos - 1), vbUpperCase)
                    dbLength = Mid(colType, sizeOpenPos + 1, sizeClosePos - sizeOpenPos - 1)
                Else
                    dbType = colType
                    dbLength = 0
                End If

                '===================
                ' �T���v���l�̏C��
                '===================
                sample = RebuildSampleText(sample, dbType, dbLength, codeCont)

                If dbType = "CHAR" Then
                    varchar2Size = dbLength
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_TYPE, "������")
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MINLENGTH, varchar2Size)
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, varchar2Size)
                    
                    '�T���v���̏C��
                    sample = Padding(sample, "X", varchar2Size)
                
                ElseIf dbType = "VARCHAR2" Or dbType = "VARCHAR" Or dbType = "VACHAR2" Then '�ςȊԈႢ�ɑΉ��i�ۑ�96�j
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_TYPE, "������")
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MINLENGTH, "1")
                    varchar2Size = dbLength
                    If (varchar2Size Mod 2) > 0 Then
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, (varchar2Size))
                        
                    ElseIf HankakuColumn(targetSheet.Cells(i, COL_NEW_TYPEDEF_COLUMN_NAME)) Then
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, (varchar2Size))
                        
                    Else
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, (varchar2Size / 2))
                    End If
                    
                ElseIf dbType = "NUMBER" Then
                    If InStr(dbLength, ",") > 0 Then
                        varchar2Size = Mid(dbLength, 1, InStr(dbLength, ",") - 1)
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_TYPE, "���l")
                    Else
                        varchar2Size = dbLength
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_TYPE, "����")
                    End If
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MINLENGTH, 1)
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, varchar2Size)
 
                ElseIf dbType = "DATE" Then
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_TYPE, "����")
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MINLENGTH, 1)
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, 12)

                Else
                    Call debugLog("�V�[�g���u" & targetSheet.Name & "�v/ �s�u" & i & "�v/ �\�����ʌ^�ł�:" & targetSheet.Name & ":" & colType & ":" & dbType, ERROR_LOG)
                End If
                
                '����
                Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_EXPLAIN, explain)
                
                '�f�[�^��
                Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_DATA_SAMPLE, sample)
                
                ' ������E����
                If codeCont <> "" Then
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, codeCont & "�̂����ꂩ")
                Else
                    If dbType = "CHAR" Then
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "���p�p��")
                        
                        If dbLength = 8 And targetSheet.Cells(i, COL_NEW_TYPEDEF_NAME) Like "*��" Then
                            sample = "20160505"
                        End If
                        
                        If dbLength = 8 And Match("^(\d{4})(\d{2})(\d{2})$", sample) Then
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "yyyyMMdd")
                            '�f�[�^��
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_DATA_SAMPLE, "'20160505")
                        End If
                    
                    ElseIf dbType = "VARCHAR2" Then
                        varchar2Size = dbLength
                        If (varchar2Size Mod 2) > 0 Then
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "���p�p��")
                            
                        ElseIf HankakuColumn(targetSheet.Cells(i, COL_NEW_TYPEDEF_COLUMN_NAME)) Then
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "���p�p��")
                            
                        ElseIf Match("^(\d{4})(\d{2})(\d{2})$", sample) Then
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "yyyyMMdd")
                            '�f�[�^��
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_DATA_SAMPLE, "'20160505")

                        Else
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "")
                        End If
                        
                    ElseIf dbType = "NUMBER" Then
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "���p����")
                    
                    ElseIf dbType = "DATE" Then
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "")
                        '�f�[�^��
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_DATA_SAMPLE, "2002-05-30T09:30:10+06:00")

                    End If
                End If
                
            End If
CONTINUE:
        Next
    End If
    
End Sub

Private Function RebuildSampleText(origSample As String, dbType As String, dbLength As String, codeDef As String)

End Function

Private Function Padding(target As String, c As String, length As Integer) As String
    Dim base As String, i As Integer
    Dim sjisStr As String
    sjisStr = StrConv(target, vbFromUnicode)
    
    If LenB(sjisStr) < length Then
        For i = 1 To length
            base = base & c
        Next
        
        Padding = target & Mid(base, 1, length - LenB(sjisStr))
    Else
        Padding = target
    End If
    
End Function



'*******************************************************************************
' �Z���ɑ΂��Đݒ肷��B
'*******************************************************************************
Private Sub SetCellValue(mode As String, shet As Worksheet, rowIdx As Integer, colIdx As Integer, celVal As Variant)
    If mode = "1" Then
        shet.Cells(rowIdx, colIdx) = celVal
    
    ElseIf mode = "2" Then
        If shet.Cells(rowIdx, colIdx) = "" Then
            shet.Cells(rowIdx, colIdx) = celVal
        End If

    ElseIf mode = "3" Then
        If shet.Cells(rowIdx, colIdx) <> celVal Then
            Call OutputCheckLog(shet, rowIdx, colIdx, celVal)
        End If
    End If
    
End Sub

Private Function HankakuColumn(columnName As String) As Boolean
    If columnName Like "*�ԍ�" Or columnName Like "*�敪" Or columnName Like "*ID" Or columnName Like "*�h�c" Or columnName Like "*�p����" Then
        HankakuColumn = True
    End If
    HankakuColumn = False
End Function

Private Function Match(strPattern As String, target As String) As Boolean
    Dim RegexObj, i As Long, msg As String
    Set RegexObj = CreateObject("VBScript.RegExp")

    With RegexObj
        .pattern = strPattern
        .IgnoreCase = False
        .Global = True
        Match = .Test(target)
    End With
    Set RegexObj = Nothing
End Function

'*******************************************************************************
' �`�F�b�N���O�o��
'*******************************************************************************
Public Sub OutputCheckLog(shet As Worksheet, rowIdx As Integer, colIdx As Integer, newVal As Variant)
    Dim oldVal As Variant
    Set oldVal = shet.Cells(rowIdx, colIdx)
    Dim colName As String
    colName = shet.Cells(rowIdx, COL_NEW_TYPEDEF_NAME)
    Call debugLog("�V�[�g���u" & shet.Name & "�v/ ���ځu" & colName & "�v/ �Z���urow=" & rowIdx & "," & "col=" & colIdx & "�v/ �ύX�O�u" & oldVal & "�v/�ύX��u" & newVal & "�v", OUTPUT_LOG)
End Sub