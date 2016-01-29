VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************
' �萔��`
'*******************************************************************************

'�R�[�h��`���̃t�@�C���p�X
Private Const CODE_DEFINE_FILE_PATH = "\\192.168.10.250\Common\AP�`�[�����L�b��\90_���̑�\99_cijnext\����Adam\�̔�API\�݌v�Ώ�DB��`\03_�R�[�h�ꗗ\�R�[�h��`��.xlsx"
Private Const CODE_DEFINE_SHEET_NAME = "�����R�[�h��`"

'�R�[�h��`��_���`
Private Const COL_CODEDEF_CODE = 3              ' �R�[�h
Private Const COL_CODEDEF_CODE_VALUE = 4        ' �R�[�h�l
Private Const COL_CODEDEF_CODE_NAME = 5         ' �R�[�h����
Private Const COL_CODEDEF_CODE_TABLES = 6       ' �g�p�e�[�u����

'*******************************************************************************
' �t�B�[���h��`
'*******************************************************************************

' �R�[�h��`����
Private m_codeDefDic As Object


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
    If m_codeDefDic Is Nothing Then
        Call LoadDictionary
    End If
    Set GetDictionary = m_codeDefDic
End Function
'*******************************************************************************
' �R�[�h��`�����玫���쐬
'*******************************************************************************
Private Sub LoadDictionary()
 
    Dim xlAPP As Application        ' Excel.Application
    Dim dic As Object               ' Scripting.Dictionary
    Dim CodeDefineBook As Workbook
    Dim objWBK As Workbook          ' ���[�N�u�b�NObject
    
    Set dic = CreateObject("Scripting.Dictionary")
    Set xlAPP = Application

    Application.ScreenUpdating = False
    ' �X�e�[�^�X�o�[�ɏ�������\��
    xlAPP.StatusBar = "�R�[�h��`���[�h���D�D�D�D"
    
    ' ���[�N�u�b�N���J��
    Set CodeDefineBook = Workbooks.Open(fileName:=CODE_DEFINE_FILE_PATH, _
                                UpdateLinks:=False, _
                                ReadOnly:=True)
    CodeDefineBook.Activate
    
    
    Dim i As Integer
    Dim targetSheet As Worksheet
    Set targetSheet = CodeDefineBook.Sheets(CODE_DEFINE_SHEET_NAME)
    For i = 2 To targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).row
        Dim keys As Collection
        Dim key As Variant
        
        ' �R�[�h��̒�`���Ȃ��Ȃ�܂Ń��[�v
        If targetSheet.Cells(i, COL_CODEDEF_CODE) = "" Then
            Exit For
        End If
        ' �g�p�e�[�u���͕�����`����Ă��邱�Ƃ�����ׁA�R���N�V�����Ŏ擾
        Set keys = CreateKeys(CodeDefineBook, i)
        For Each key In keys
            dic(key) = dic(key) & MakeDefineText( _
                    targetSheet.Cells(i, COL_CODEDEF_CODE_VALUE), _
                    targetSheet.Cells(i, COL_CODEDEF_CODE_NAME))
        Next key
    Next i
    
    CodeDefineBook.Close
    xlAPP.StatusBar = False
    Application.ScreenUpdating = True
    Set m_codeDefDic = dic
 
End Sub

'*******************************************************************************
' �R�[�h��`������݌v���ɋL�ڂ�����l�̕�����𐶐�
'*******************************************************************************
Private Function MakeDefineText(val As String, codeName As String) As String
    MakeDefineText = "�u" & val & "�v" & "(" & codeName & ")"     '�u0�v�i���L�����Z���j
End Function


'*******************************************************************************
' �R�[�h��`���̎g�p�e�[�u������e�[�u�������擾
'*******************************************************************************
Private Function CreateKeys(codeDifineBook As Workbook, rowIdx As Integer) As Collection
    Dim code As String
    Dim table() As String
    Dim tableName As Variant
    Dim list As Collection
    
    code = Cells(rowIdx, COL_CODEDEF_CODE)     '�R�[�h
    table = Split(Cells(rowIdx, COL_CODEDEF_CODE_TABLES), ",")
  
    Set list = New Collection
    
    For Each tableName In table()
        Call list.Add(tableName & "." & code)
    Next tableName
    
    Set CreateKeys = list
End Function

'*******************************************************************************
' �L�[���w�肵�Ēl���擾
' �L�[ = tableName , columnName
'*******************************************************************************
Public Function GetValue(tableName As String, code As String) As String
    Dim dic As Object, key As String
    Set dic = GetDictionary()
    key = GenerateKey(tableName, code)
    If dic.Exists(key) Then
        GetValue = dic(key)
    Else
        GetValue = ""
    End If
End Function

'*******************************************************************************
' �L�[�𐶐�
'*******************************************************************************
Private Function GenerateKey(tableName As String, code As String) As String
    GenerateKey = tableName & "." & code
End Function
