VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************
' �萔��`
'*******************************************************************************

'���ڒ�`���̃t�@�C���p�X
Private Const CODE_DEFINE_FILE_PATH = "\\192.168.10.250\Common\AP�`�[�����L�b��\90_���̑�\99_cijnext\����Adam\�̔�API\�݌v�Ώ�DB��`\02_���ڒ�`��\���ڒ�`��.xlsx"
Private Const CODE_DEFINE_SHEET_NAME = "�������ڒ�`"

'���ڒ�`��_���`
Private Const COL_TABLE_NAME = 1              ' �e�[�u�����i�_���j
Private Const COL_TABLE_NAME_P = 2              ' �e�[�u�����i�����j
Private Const COL_COLUMN_NAME = 3               ' ���ږ�
Private Const COL_COLUMN_TYPE = 5               ' ���ڃ^�C�v
Private Const COL_COLUMN_PKEY = 6               ' PKEY
Private Const COL_COLUMN_FKEY = 7               ' FKEY
Private Const COL_COLUMN_NOTNULL = 8            ' NOTNULL
Private Const COL_COLUMN_EXPLAIN = 9            ' ����
Private Const COL_COLUMN_SAMPLE = 17

Private Const ELEMENT_COUNT = 5

'*******************************************************************************
' �t�B�[���h��`
'*******************************************************************************

' �R�[�h��`����
Private m_defDic As Object


'*******************************************************************************
' �R���X�g���N�^
'*******************************************************************************
Private Sub Class_Initialize()
End Sub

'*******************************************************************************
' �f�X�g���N�^
'*******************************************************************************
Private Sub Class_Terminate()
    Debug.Print ("Destructor is called.")
End Sub

'*******************************************************************************
' �R�[�h��`�����玫���쐬
'*******************************************************************************
Public Function GetDictionary() As Object
    If m_defDic Is Nothing Then
        Call LoadDictionary
    End If
    Set GetDictionary = m_defDic
End Function
'*******************************************************************************
' �R�[�h��`�����玫���쐬
'*******************************************************************************
Private Sub LoadDictionary()
 
    Dim xlAPP As Application        ' Excel.Application
    Dim dic As Object               ' Scripting.Dictionary
    Dim CodeDefineBook As Workbook
    Dim objWBK As Workbook          ' ���[�N�u�b�NObject

    
 On Error GoTo Error_Handler
 
    Set dic = CreateObject("Scripting.Dictionary")
    Set xlAPP = Application

    Application.ScreenUpdating = False
    ' �X�e�[�^�X�o�[�ɏ�������\��
    xlAPP.StatusBar = "���ڒ�`���[�h���D�D�D�D"
    
    ' ���[�N�u�b�N���J��
    Set CodeDefineBook = Workbooks.Open(fileName:=CODE_DEFINE_FILE_PATH, _
                                UpdateLinks:=False, _
                                ReadOnly:=True)
    CodeDefineBook.Activate
    
    
    Dim i As Integer
    Dim targetSheet As Worksheet
    Set targetSheet = CodeDefineBook.Sheets(CODE_DEFINE_SHEET_NAME)
    For i = 4 To targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).row
        Dim key_table_and_colmun As Variant '�e�[�u�����ƃJ����������Ȃ�L�[
        Dim key_colmun As Variant  '�J����������Ȃ�L�[
        Dim val As String
        
        ' �R�[�h��̒�`���Ȃ��Ȃ�܂Ń��[�v
        If targetSheet.Cells(i, COL_TABLE_NAME) = "" Then
            Exit For
        End If
        If targetSheet.Cells(i, COL_TABLE_NAME_P) <> "" Then
            ' �g�p�e�[�u���͕�����`����Ă��邱�Ƃ�����ׁA�R���N�V�����Ŏ擾
            key_table_and_colmun = targetSheet.Cells(i, COL_TABLE_NAME) & "@" & targetSheet.Cells(i, COL_COLUMN_NAME)
            key_colmun = targetSheet.Cells(i, COL_COLUMN_NAME)
            
            val = MakeValue(targetSheet, i)
            
            '�e�[�u�����ƃJ����������Ȃ�L�[�ɃG���g��
            If dic.Exists(key_table_and_colmun) Then
                ' �^�^�C�v�܂Ń`�F�b�N
                If Mid(dic(key_table_and_colmun), 1, InStr(dic(key_table_and_colmun), "@")) <> Mid(val, 1, InStr(val, "@")) Then
                    Call debugLog("ERR:" & key_table_and_colmun & "�͒�`���d�����Ă��܂��B", ERROR_LOG)
                    dic.Add(key_table_and_colmun) = MakeErrorValue("ERR:" & key_table_and_colmun & "�͒�`���d�����Ă��܂��B")
                End If
            Else
                dic(key_table_and_colmun) = val
            End If
            
            '�J�����������ł����삷��悤�ɃJ�����������̃L�[���G���g������
            If dic.Exists(key_colmun) Then
                If Mid(dic(key_table_and_colmun), 1, InStr(dic(key_table_and_colmun), "@")) <> Mid(val, 1, InStr(val, "@")) Then
                    Call debugLog("ERR:" & key_colmun & "�͒�`���d�����Ă��܂��B", ERROR_LOG)
                    dic(key_colmun) = MakeErrorValue("ERR:" & key_colmun & "�͒�`���d�����Ă��܂��B")
                End If
            Else
                dic(key_colmun) = val
            End If
            
        End If
    Next i
    
    CodeDefineBook.Close
    xlAPP.StatusBar = False
    Application.ScreenUpdating = True
    Set m_defDic = dic
    
    GoTo Finally

Error_Handler:
    'Resume Sub_EXIT
    Err.Raise 5001
    
Finally:

End Sub

Private Function MakeValue(shet As Worksheet, row As Integer) As String
    Dim values(ELEMENT_COUNT) As String
    values(0) = GetCellValue(shet, row, COL_COLUMN_TYPE)
    values(1) = GetCellValue(shet, row, COL_COLUMN_PKEY)
    values(2) = GetCellValue(shet, row, COL_COLUMN_FKEY)
    values(3) = GetCellValue(shet, row, COL_COLUMN_NOTNULL)
    values(4) = GetCellValue(shet, row, COL_COLUMN_EXPLAIN)
    values(5) = GetCellValue(shet, row, COL_COLUMN_SAMPLE)
    
    MakeValue = Join(values, "@")
    
End Function

Private Function GetCellValue(shet As Worksheet, row As Integer, col As Integer)
    GetCellValue = ""
    If Not IsError(shet.Cells(row, col).Value) Then
        GetCellValue = shet.Cells(row, col)
    End If
End Function

Private Function MakeErrorValue(errStr As String) As String
    Dim values(ELEMENT_COUNT) As String
    values(0) = errStr
    MakeErrorValue = Join(values, "@")
    
End Function

'*******************************************************************************
' �L�[���w�肵�Ēl���擾
' �L�[ = tableName , columnName
'*******************************************************************************
Public Function GetValue(tableName As String, column As String) As String
    Dim dic As Object, key As String
    Set dic = GetDictionary()
    key = GenerateKey(tableName, column)
    If dic.Exists(key) Then
        GetValue = dic(key)
    Else
        GetValue = ""
    End If
End Function

'*******************************************************************************
' �L�[�𐶐�
'*******************************************************************************
Private Function GenerateKey(tableName As String, column As String) As String
    If tableName = "" Then
        GenerateKey = column
    Else
        GenerateKey = tableName & "@" & column
    End If
    
End Function