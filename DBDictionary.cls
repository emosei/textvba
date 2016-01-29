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
' 定数定義
'*******************************************************************************

'項目定義書のファイルパス
Private Const CODE_DEFINE_FILE_PATH = "\\192.168.10.250\Common\APチーム共有暫定\90_その他\99_cijnext\次期Adam\販売API\設計対象DB定義\02_項目定義書\項目定義書.xlsx"
Private Const CODE_DEFINE_SHEET_NAME = "次期項目定義"

'項目定義書_列定義
Private Const COL_TABLE_NAME = 1              ' テーブル名（論理）
Private Const COL_TABLE_NAME_P = 2              ' テーブル名（物理）
Private Const COL_COLUMN_NAME = 3               ' 項目名
Private Const COL_COLUMN_TYPE = 5               ' 項目タイプ
Private Const COL_COLUMN_PKEY = 6               ' PKEY
Private Const COL_COLUMN_FKEY = 7               ' FKEY
Private Const COL_COLUMN_NOTNULL = 8            ' NOTNULL
Private Const COL_COLUMN_EXPLAIN = 9            ' 説明
Private Const COL_COLUMN_SAMPLE = 17

Private Const ELEMENT_COUNT = 5

'*******************************************************************************
' フィールド定義
'*******************************************************************************

' コード定義辞書
Private m_defDic As Object


'*******************************************************************************
' コンストラクタ
'*******************************************************************************
Private Sub Class_Initialize()
End Sub

'*******************************************************************************
' デストラクタ
'*******************************************************************************
Private Sub Class_Terminate()
    Debug.Print ("Destructor is called.")
End Sub

'*******************************************************************************
' コード定義書から辞書作成
'*******************************************************************************
Public Function GetDictionary() As Object
    If m_defDic Is Nothing Then
        Call LoadDictionary
    End If
    Set GetDictionary = m_defDic
End Function
'*******************************************************************************
' コード定義書から辞書作成
'*******************************************************************************
Private Sub LoadDictionary()
 
    Dim xlAPP As Application        ' Excel.Application
    Dim dic As Object               ' Scripting.Dictionary
    Dim CodeDefineBook As Workbook
    Dim objWBK As Workbook          ' ワークブックObject

    
 On Error GoTo Error_Handler
 
    Set dic = CreateObject("Scripting.Dictionary")
    Set xlAPP = Application

    Application.ScreenUpdating = False
    ' ステータスバーに処理中を表示
    xlAPP.StatusBar = "項目定義ロード中．．．．"
    
    ' ワークブックを開く
    Set CodeDefineBook = Workbooks.Open(fileName:=CODE_DEFINE_FILE_PATH, _
                                UpdateLinks:=False, _
                                ReadOnly:=True)
    CodeDefineBook.Activate
    
    
    Dim i As Integer
    Dim targetSheet As Worksheet
    Set targetSheet = CodeDefineBook.Sheets(CODE_DEFINE_SHEET_NAME)
    For i = 4 To targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).row
        Dim key_table_and_colmun As Variant 'テーブル名とカラム名からなるキー
        Dim key_colmun As Variant  'カラム名からなるキー
        Dim val As String
        
        ' コード列の定義がなくなるまでループ
        If targetSheet.Cells(i, COL_TABLE_NAME) = "" Then
            Exit For
        End If
        If targetSheet.Cells(i, COL_TABLE_NAME_P) <> "" Then
            ' 使用テーブルは複数定義されていることがある為、コレクションで取得
            key_table_and_colmun = targetSheet.Cells(i, COL_TABLE_NAME) & "@" & targetSheet.Cells(i, COL_COLUMN_NAME)
            key_colmun = targetSheet.Cells(i, COL_COLUMN_NAME)
            
            val = MakeValue(targetSheet, i)
            
            'テーブル名とカラム名からなるキーにエントリ
            If dic.Exists(key_table_and_colmun) Then
                ' 型タイプまでチェック
                If Mid(dic(key_table_and_colmun), 1, InStr(dic(key_table_and_colmun), "@")) <> Mid(val, 1, InStr(val, "@")) Then
                    Call debugLog("ERR:" & key_table_and_colmun & "は定義が重複しています。", ERROR_LOG)
                    dic.Add(key_table_and_colmun) = MakeErrorValue("ERR:" & key_table_and_colmun & "は定義が重複しています。")
                End If
            Else
                dic(key_table_and_colmun) = val
            End If
            
            'カラム名だけでも動作するようにカラム名だけのキーもエントリする
            If dic.Exists(key_colmun) Then
                If Mid(dic(key_table_and_colmun), 1, InStr(dic(key_table_and_colmun), "@")) <> Mid(val, 1, InStr(val, "@")) Then
                    Call debugLog("ERR:" & key_colmun & "は定義が重複しています。", ERROR_LOG)
                    dic(key_colmun) = MakeErrorValue("ERR:" & key_colmun & "は定義が重複しています。")
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
' キーを指定して値を取得
' キー = tableName , columnName
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
' キーを生成
'*******************************************************************************
Private Function GenerateKey(tableName As String, column As String) As String
    If tableName = "" Then
        GenerateKey = column
    Else
        GenerateKey = tableName & "@" & column
    End If
    
End Function
