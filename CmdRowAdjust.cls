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
' コード定義書の内容を元に、設計書の特定のセルに値を設定する
'*******************************************************************************

Option Explicit

'*******************************************************************************
' 定数定義
'*******************************************************************************

'インタフェース仕様書_型定義シートの列定義
'Private Const COL_SPEC_EXTRA = 17 ' コード定義の記述列
Private Const COL_SPEC_NO = 1 ' NO列

'インタフェース仕様書_型定義シートのシート名接頭語
Private Const TARGET_SHEET_PREFIX = "型定義_"

'*******************************************************************************
' コンストラクタ
'*******************************************************************************
Private Sub Class_Initialize()
    Debug.Print ("CmdRowAdjust Constructor is called.")
End Sub

'*******************************************************************************
' デストラクタ
'*******************************************************************************
Private Sub Class_Terminate()
    Debug.Print ("CmdRowAdjust Destructor is called.")
End Sub

'*******************************************************************************
' 実処理
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
            
            ' No列の定義がなくなるまでループ
            If targetSheet.Cells(i, COL_SPEC_NO) = "" Then
                Exit For
            End If
            
            Call AdgustRowHeight(targetSheet, i)

        Next
    End If
    
End Sub

'*******************************************************************************
' 行の高さを調整する。
'*******************************************************************************
Public Sub AdgustRowHeight(adjustShee As Worksheet, _
                            rowIdx As Integer)
    
    adjustShee.Rows(rowIdx).AutoFit
    adjustShee.Rows(rowIdx).rowHeight = adjustShee.Rows(rowIdx).Height + 5
                                                        
End Sub



