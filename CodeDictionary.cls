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
' 定数定義
'*******************************************************************************

'コード定義書のファイルパス
Private Const CODE_DEFINE_FILE_PATH = "\\192.168.10.250\Common\APチーム共有暫定\90_その他\99_cijnext\次期Adam\販売API\設計対象DB定義\03_コード一覧\コード定義書.xlsx"
Private Const CODE_DEFINE_SHEET_NAME = "次期コード定義"

'コード定義書_列定義
Private Const COL_CODEDEF_CODE = 3              ' コード
Private Const COL_CODEDEF_CODE_VALUE = 4        ' コード値
Private Const COL_CODEDEF_CODE_NAME = 5         ' コード名称
Private Const COL_CODEDEF_CODE_TABLES = 6       ' 使用テーブル名

'*******************************************************************************
' フィールド定義
'*******************************************************************************

' コード定義辞書
Private m_codeDefDic As Object


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
    If m_codeDefDic Is Nothing Then
        Call LoadDictionary
    End If
    Set GetDictionary = m_codeDefDic
End Function
'*******************************************************************************
' コード定義書から辞書作成
'*******************************************************************************
Private Sub LoadDictionary()
 
    Dim xlAPP As Application        ' Excel.Application
    Dim dic As Object               ' Scripting.Dictionary
    Dim CodeDefineBook As Workbook
    Dim objWBK As Workbook          ' ワークブックObject
    
    Set dic = CreateObject("Scripting.Dictionary")
    Set xlAPP = Application

    Application.ScreenUpdating = False
    ' ステータスバーに処理中を表示
    xlAPP.StatusBar = "コード定義ロード中．．．．"
    
    ' ワークブックを開く
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
        
        ' コード列の定義がなくなるまでループ
        If targetSheet.Cells(i, COL_CODEDEF_CODE) = "" Then
            Exit For
        End If
        ' 使用テーブルは複数定義されていることがある為、コレクションで取得
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
' コード定義書から設計書に記載する候補値の文字列を生成
'*******************************************************************************
Private Function MakeDefineText(val As String, codeName As String) As String
    MakeDefineText = "「" & val & "」" & "(" & codeName & ")"     '「0」（未キャンセル）
End Function


'*******************************************************************************
' コード定義書の使用テーブルからテーブル名を取得
'*******************************************************************************
Private Function CreateKeys(codeDifineBook As Workbook, rowIdx As Integer) As Collection
    Dim code As String
    Dim table() As String
    Dim tableName As Variant
    Dim list As Collection
    
    code = Cells(rowIdx, COL_CODEDEF_CODE)     'コード
    table = Split(Cells(rowIdx, COL_CODEDEF_CODE_TABLES), ",")
  
    Set list = New Collection
    
    For Each tableName In table()
        Call list.Add(tableName & "." & code)
    Next tableName
    
    Set CreateKeys = list
End Function

'*******************************************************************************
' キーを指定して値を取得
' キー = tableName , columnName
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
' キーを生成
'*******************************************************************************
Private Function GenerateKey(tableName As String, code As String) As String
    GenerateKey = tableName & "." & code
End Function

