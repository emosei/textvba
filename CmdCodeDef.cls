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
' コード定義書の内容を元に、設計書の特定のセルに値を設定する
'*******************************************************************************

Option Explicit

'*******************************************************************************
' 定数定義
'*******************************************************************************

'インタフェース仕様書_型定義シートの列定義
'Private Const COL_SPEC_EXTRA = 17 ' コード定義の記述列
Private Const COL_SPEC_NO = 1 ' NO列
'Private Const COL_SPEC_FORMAT = 12 ' 文字種・書式列

Private Const COL_SPEC_FORMAT = 14 ' 文字種・書式列

'-------------------------------------------------
' コード定義書を参照するキーを記述する列番号
' 「テーブル名.コード名」の形式で記述
'-------------------------------------------------
Private Const COL_SPEC_EXTRA = 15  ' コード定義の記述列

'インタフェース仕様書_型定義シートのシート名接頭語
Private Const TARGET_SHEET_PREFIX = "型定義_"

' コード定義辞書
Private m_codeDefDic As CodeDictionary

'*******************************************************************************
' コンストラクタ
'*******************************************************************************
Private Sub Class_Initialize()
    Set m_codeDefDic = New CodeDictionary
    Debug.Print ("CmdCodeDef Constructor is called.")
End Sub

'*******************************************************************************
' デストラクタ
'*******************************************************************************
Private Sub Class_Terminate()
    Debug.Print ("CmdCodeDef Destructor is called.")
End Sub


'*******************************************************************************
' 実処理
'*******************************************************************************
Public Sub ExecCommand(targetSheet As Worksheet)

    Dim defineCellValue As String
    Dim defineKey As String
    Dim coronPos As Integer
    
    Dim sheetName As String
    
    sheetName = targetSheet.Name
    
    ' 辞書を取得
    Dim codeDefDic As Object
    Set codeDefDic = m_codeDefDic.GetDictionary()
    
    If InStr(sheetName, TARGET_SHEET_PREFIX) = 1 Then
        Dim i As Integer
        For i = 2 To targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).row
            
            ' No列の定義がなくなるまでループ
            If targetSheet.Cells(i, COL_SPEC_NO) = "" Then
                Exit For
            End If
        
            defineCellValue = targetSheet.Cells(i, COL_SPEC_EXTRA)
            If defineCellValue <> "" Then
            
                'Dim crPos As Integer
                'crPos = WorksheetFunction.Min(InStr(defineCellValue, vbCr), _
                '                                InStr(defineCellValue, vbLf), _
                '                                InStr(defineCellValue, vbCrLf))
                ' 改行コードがある場合には改行コード以降の文字列を退避
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
                '設定する値を生成する
                '---------------------------------------------------------------------------
                If codeDefDic.Exists(defineKey) Then
                    targetSheet.Cells(i, COL_SPEC_FORMAT).Value = codeDefDic(defineKey) & "のいずれか"
                End If
                
            End If

        Next
    End If
    
End Sub


