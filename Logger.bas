Attribute VB_Name = "Logger"
Public Const LOG_SW = True      'ログ出力の有効・無効
Public Const LOG_SHEET = "Log"  'ログの出力先シート名
Public Const MAX_LOG = 2         'ログの本数（何本でも）
Public Const LOG_START_ROW = 2     'ログ出力開始行
Public Const INIT_LOG = "INIT"    '初期化指定
Public Const ERROR_LOG = 1         'ERROR 用ログ
Public Const OUTPUT_LOG = 2         'データ吐き出しログ

'以下は、ログ吐き出しまくりマクロ本体。（どのモジュールに書いても可。）
Public Sub debugLog(log As String, opt As Integer)
  Static last_line(MAX_LOG) As Integer    '次にログを書く行
  If Not LOG_SW Then Exit Sub
  If log = INIT_LOG Then
    For n = 1 To MAX_LOG
      last_line(n) = LOG_START_ROW
      'ログ領域のクリア
      With ThisWorkbook.Worksheets(LOG_SHEET)
        .Range(.Cells(LOG_START_ROW, n), _
                   .Cells(Rows.Count, n)).Value = ""
      End With
    Next
    Exit Sub
  End If
  ThisWorkbook.Worksheets(LOG_SHEET).Cells(last_line(opt), opt).Value = log
  last_line(opt) = last_line(opt) + 1
End Sub

