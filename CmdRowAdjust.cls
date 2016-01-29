VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmdRowAdjust"
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

'�C���^�t�F�[�X�d�l��_�^��`�V�[�g�̃V�[�g���ړ���
Private Const TARGET_SHEET_PREFIX = "�^��`_"

'*******************************************************************************
' �R���X�g���N�^
'*******************************************************************************
Private Sub Class_Initialize()
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
    
    If InStr(sheetName, TARGET_SHEET_PREFIX) = 1 Then
    'If True Then
        For i = 2 To targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).row
            
            ' No��̒�`���Ȃ��Ȃ�܂Ń��[�v
            If targetSheet.Cells(i, COL_SPEC_NO) = "" Then
                Exit For
            End If
            
            Call AdgustRowHeight(targetSheet, i)

        Next
    End If
    
End Sub

'*******************************************************************************
' �s�̍����𒲐�����B
'*******************************************************************************
Public Sub AdgustRowHeight(adjustShee As Worksheet, _
                            rowIdx As Integer)
    
    adjustShee.Rows(rowIdx).AutoFit
    adjustShee.Rows(rowIdx).rowHeight = adjustShee.Rows(rowIdx).Height + 5
                                                        
End Sub


