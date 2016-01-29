VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmdCodeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' �R�[�h��`���̓��e�����ɁA�݌v���̓���̃Z���ɒl��ݒ肷��
'*******************************************************************************

Option Explicit

'*******************************************************************************
' �萔��`
'*******************************************************************************

'�C���^�t�F�[�X�d�l��_�^��`�V�[�g�̗��`
'Private Const COL_SPEC_EXTRA = 17 ' �R�[�h��`�̋L�q��
Private Const COL_SPEC_NO = 1 ' NO��
'Private Const COL_SPEC_FORMAT = 12 ' ������E������

Private Const COL_SPEC_FORMAT = 14 ' ������E������

'-------------------------------------------------
' �R�[�h��`�����Q�Ƃ���L�[���L�q�����ԍ�
' �u�e�[�u����.�R�[�h���v�̌`���ŋL�q
'-------------------------------------------------
Private Const COL_SPEC_EXTRA = 15  ' �R�[�h��`�̋L�q��

'�C���^�t�F�[�X�d�l��_�^��`�V�[�g�̃V�[�g���ړ���
Private Const TARGET_SHEET_PREFIX = "�^��`_"

' �R�[�h��`����
Private m_codeDefDic As CodeDictionary

'*******************************************************************************
' �R���X�g���N�^
'*******************************************************************************
Private Sub Class_Initialize()
    Set m_codeDefDic = New CodeDictionary
    Debug.Print ("CmdCodeDef Constructor is called.")
End Sub

'*******************************************************************************
' �f�X�g���N�^
'*******************************************************************************
Private Sub Class_Terminate()
    Debug.Print ("CmdCodeDef Destructor is called.")
End Sub


'*******************************************************************************
' ������
'*******************************************************************************
Public Sub ExecCommand(targetSheet As Worksheet)

    Dim defineCellValue As String
    Dim defineKey As String
    Dim coronPos As Integer
    
    Dim sheetName As String
    
    sheetName = targetSheet.Name
    
    ' �������擾
    Dim codeDefDic As Object
    Set codeDefDic = m_codeDefDic.GetDictionary()
    
    If InStr(sheetName, TARGET_SHEET_PREFIX) = 1 Then
        Dim i As Integer
        For i = 2 To targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).row
            
            ' No��̒�`���Ȃ��Ȃ�܂Ń��[�v
            If targetSheet.Cells(i, COL_SPEC_NO) = "" Then
                Exit For
            End If
        
            defineCellValue = targetSheet.Cells(i, COL_SPEC_EXTRA)
            If defineCellValue <> "" Then
            
                'Dim crPos As Integer
                'crPos = WorksheetFunction.Min(InStr(defineCellValue, vbCr), _
                '                                InStr(defineCellValue, vbLf), _
                '                                InStr(defineCellValue, vbCrLf))
                ' ���s�R�[�h������ꍇ�ɂ͉��s�R�[�h�ȍ~�̕������ޔ�
                'If 0 < crPos Then
                '    defineCellValue = Mid(defineCellValue, 0, crPos)
                'End If
                
                'coronPos = InStr(defineCellValue, ":")
                coronPos = 0
                If 0 < coronPos Then
                    defineKey = Mid(defineCellValue, coronPos)
                Else
                    defineKey = defineCellValue
                End If
                
                '---------------------------------------------------------------------------
                '�ݒ肷��l�𐶐�����
                '---------------------------------------------------------------------------
                If codeDefDic.Exists(defineKey) Then
                    targetSheet.Cells(i, COL_SPEC_FORMAT).Value = codeDefDic(defineKey) & "�̂����ꂩ"
                End If
                
            End If

        Next
    End If
    
End Sub

