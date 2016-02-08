Attribute VB_Name = "WorkUtil"
'*******************************************************************************
'  ワークブック順次処理
'*******************************************************************************
Option Explicit

Private Const cnsYEN = "\"

'*******************************************************************************
'  旧シート名
'*******************************************************************************
Private Const SHEET_NAME_COVER_OLD = "表紙"
Private Const SHEET_NAME_HISTORY_OLD = "改訂履歴"
Private Const SHEET_NAME_INOUT_OLD = "1.入出力定義"
Private Const SHEET_NAME_TYPE_LIST_OLD = "2.要素一覧"

'*******************************************************************************
'  新シート名
'*******************************************************************************
Private Const SHEET_NAME_COVER_NEW = "表紙"
Private Const SHEET_NAME_HISTORY_NEW = "改訂履歴"
Private Const SHEET_NAME_REF_LIST_NEW = "1.参照ファイル一覧"
Private Const SHEET_NAME_COMMON_NEW = "2.共通情報"
Private Const SHEET_NAME_TYPE_LIST_NEW = "3.型一覧"
Private Const SHEET_NAME_INOUT_NEW = "4.入出力定義"
Private Const TEMPLATE_SHEET_NAME = "型定義_カート素材予約応答"

'*******************************************************************************
'  型シートプレフィックス
'*******************************************************************************
Private Const TARGET_SHEET_PREFIX_OLD = "4."
Private Const TARGET_SHEET_PREFIX_NEW = "型定義_"
Private Const TEMPLATE_TYPE_DEF_SHEET_MAX_ROW = 50
Private Const TEMPLATE_REF_LIST_SHEET_MAX_ROW = 50
Private Const TEMPLATE_TYPE_LIST_SHEET_MAX_ROW = 50
'テンプレートファイルのファイルパス
Private Const TEMPLATE_FILE_PATH = "\\192.168.10.250\common\APチーム共有暫定\03_外部内部\02_規約類\01_設計書規約\06_他システム関連\他システムインターフェース仕様書_IFAWS000I_カート素材予約API.xlsx"

'*******************************************************************************
'  破棄モード
' 1 -- シート削除
' 2 -- シート名変更
'*******************************************************************************
Private Const DESTRUCTION_MODE_DELETE = 1
Private Const DESTRUCTION_MODE_RENAME = 2
Private Const DESTRUCTION_MODE = DESTRUCTION_MODE_DELETE


'*******************************************************************************
'他システムインターフェース仕様書_旧型定義シート（要素一覧シート）_列定義
'*******************************************************************************
Private Const COL_OLD_TYPEDEF_NO = 1 ' NO列
Private Const COL_OLD_TYPEDEF_NAME = 2 ' 項目名
Private Const COL_OLD_TYPEDEF_ID = 3 ' 項目ID
Private Const COL_OLD_TYPEDEF_EXPLAIN = 4 ' 説明
Private Const COL_OLD_TYPEDEF_REQIRED = 5 ' 必須
Private Const COL_OLD_TYPEDEF_COND_REQ = 6 ' 必須条件
Private Const COL_OLD_TYPEDEF_MIN_LOOP = 7 ' 最小繰り返し
Private Const COL_OLD_TYPEDEF_MAX_LOOP = 8 ' 最大繰り返し
Private Const COL_OLD_TYPEDEF_TYPE = 9 ' 型
Private Const COL_OLD_TYPEDEF_TYPE_DETAIL = 10 ' 型詳細
Private Const COL_OLD_TYPEDEF_MINLENGTH = 11 ' 最小桁数
Private Const COL_OLD_TYPEDEF_MAXLENGTH = 12 ' 最大桁数
Private Const COL_OLD_TYPEDEF_DATA_SAMPLE = 13 ' データ例
Private Const COL_OLD_TYPEDEF_BIKO = 14 ' 備考

'欄外
Private Const COL_OLD_TYPEDEF_DB_TYPE = COL_OLD_TYPEDEF_BIKO + 1 ' DB型
Private Const COL_OLD_TYPEDEF_NOT_NULL = COL_OLD_TYPEDEF_BIKO + 4 ' DB_Notnull


'*******************************************************************************
'他システムインターフェース仕様書_新型定義シート_列定義
'*******************************************************************************
Private Const COL_NEW_TYPEDEF_NO = 1 ' NO列
Private Const COL_NEW_TYPEDEF_NAME = 2 ' 項目名
Private Const COL_NEW_TYPEDEF_ID = 3 ' 項目ID
Private Const COL_NEW_TYPEDEF_EXPLAIN = 4 ' 説明
Private Const COL_NEW_TYPEDEF_MIN_LOOP = 5 ' 最小繰り返し
Private Const COL_NEW_TYPEDEF_MAX_LOOP = 6 ' 最大繰り返し
Private Const COL_NEW_TYPEDEF_COND_REQ = 7 ' 必須条件
Private Const COL_NEW_TYPEDEF_LIMIT_NUM = 8 ' 繰り返し回数制限
Private Const COL_NEW_TYPEDEF_TYPE = 9 ' 型
Private Const COL_NEW_TYPEDEF_TYPE_NAME = 10 ' 型名
Private Const COL_NEW_TYPEDEF_FILE_ID = 11 ' ファイルID
Private Const COL_NEW_TYPEDEF_CHAR_FORMAT = 12 ' 文字種・書式
Private Const COL_NEW_TYPEDEF_MINLENGTH = 13 ' 最小桁数
Private Const COL_NEW_TYPEDEF_MAXLENGTH = 14 ' 最大桁数
Private Const COL_NEW_TYPEDEF_DATA_SAMPLE = 15 ' データ例
Private Const COL_NEW_TYPEDEF_BIKO = 16 ' 備考
Private Const COL_NEW_TYPEDEF_MIN_NO = COL_NEW_TYPEDEF_NO ' 最左の列番号
Private Const COL_NEW_TYPEDEF_MAX_NO = COL_NEW_TYPEDEF_BIKO ' 最右の列番号

'*******************************************************************************
'他システムインターフェース仕様書_型一覧シート_列定義
'*******************************************************************************
Private Const COL_TYPELIST_NO = 1 ' NO列
Private Const COL_TYPELIST_LNAME = 2 ' 型論理名
Private Const COL_TYPELIST_PNAME = 3 ' 型物理名
Private Const COL_TYPELIST_BIKO = 4 ' 備考
Private Const COL_TYPELIST_MIN_NO = COL_TYPELIST_NO ' 最左の列番号
Private Const COL_TYPELIST_MAX_NO = COL_TYPELIST_BIKO ' 最右の列番号

'*******************************************************************************
'他システムインターフェース仕様書_参照ファイル一覧シート_列定義
'*******************************************************************************
Private Const COL_REFLIST_NO = 1 ' 項番
Private Const COL_REFLIST_FILE_NAME = 2 ' ファイル名
Private Const COL_REFLIST_FILE_ID = 3 ' ファイルID
Private Const COL_REFLIST_BIKO = 4 ' 備考
Private Const COL_REFLIST_MIN_NO = COL_REFLIST_NO ' 最左の列番号
Private Const COL_REFLIST_MAX_NO = COL_REFLIST_BIKO ' 最右の列番号

'*******************************************************************************
'他システムインターフェース仕様書_旧入出力定義シート_列定義
'*******************************************************************************
Private Const COL_OLD_IO_NO = 1 ' 項番
Private Const COL_OLD_IO_IO = 2 ' ファイル名
Private Const COL_OLD_IO_ROOT_ELEMENT_ID = 4 ' 最上位要素ID
Private Const COL_OLD_IO_ELEMENT_NAME = 5 ' 最上位要素名
Private Const COL_OLD_IO_BIKO = 6 ' 備考

'*******************************************************************************
'他システムインターフェース仕様書_新入出力定義シート_列定義
'*******************************************************************************
Private Const COL_NEW_IO_NO = 1 ' 項番
Private Const COL_NEW_IO_IO = 2 ' 入出力区分
Private Const COL_NEW_IO_ROOT_ELEMENT_ID = 3 ' 最上位要素ID
Private Const COL_NEW_IO_ROOT_ELEMENT_NAME = 4 ' 最上位要素名
Private Const COL_NEW_IO_TYPE_NAME = 5 ' 型名
Private Const COL_NEW_IO_FILE_ID = 6 ' ファイルID
Private Const COL_NEW_IO_BIKO = 7 ' 備考
Private Const COL_NEW_IO_MIN_NO = COL_NEW_IO_NO ' 最左の列番号
Private Const COL_NEW_IO_MAX_NO = COL_NEW_IO_BIKO ' 最右の列番号

'*******************************************************************************
'共通型一覧定義シート
'*******************************************************************************
Private Const SHEET_NAME_TYPEDEF = "共通型定義"


'*******************************************************************************
' 旧シート名から要素一覧作成
'*******************************************************************************
Public Sub 旧シート名から要素一覧作成()
    Call Doシート名から要素一覧作成(TARGET_SHEET_PREFIX_OLD)
End Sub

'*******************************************************************************
' 関数開始ログ
'*******************************************************************************
Public Sub LogStart(subName As String, symbol As String)
    Call debugLog("Start:" & subName & " # " & symbol, OUTPUT_LOG)
End Sub

'*******************************************************************************
' 関数終了ログ
'*******************************************************************************
Public Sub LogEnd(subName As String, symbol As String)
    Call debugLog("End:" & subName & " # " & symbol, OUTPUT_LOG)
End Sub

'*******************************************************************************
' エラーハンドリングログ
'*******************************************************************************
Public Sub LogErrorHandle(desc As String, subName As String, file As String)
    Call debugLog("Error:" & desc & " in " & subName & "@" & file, OUTPUT_LOG)
End Sub

'*******************************************************************************
' エラーハンドリングログ
'*******************************************************************************
Public Sub LogOutput(logtext As String, file As String)
    Call debugLog(logtext & "@" & file, OUTPUT_LOG)
End Sub

'*******************************************************************************
' 新シート名から要素一覧作成
'*******************************************************************************
Public Sub 新シート名から要素一覧作成()
    
    Call Doシート名から要素一覧作成(TARGET_SHEET_PREFIX_NEW)
End Sub
'*******************************************************************************
' シートから要素一覧作成処理関数
'*******************************************************************************
Public Sub Doシート名から要素一覧作成(targetPrefix As String)
    Call debugLog(INIT_LOG, 0)
    Dim xlAPP As Application        ' Excel.Application
    Dim strPathName As String       ' 指定フォルダ名
    Dim strFileName As String       ' 検出したファイル名
    Dim swESC As Boolean            ' Escキー判定
    
    Call LogStart("Doシート名から要素一覧作成", "")
    ' ｢フォルダの参照｣よりフォルダ名の取得(modFolderPicker1に収容)
    strPathName = FolderDialog("フォルダを指定して下さい", True)
    If strPathName = "" Then Exit Sub
    
    ' 指定フォルダ内のExcelワークブックのファイル名を参照する(1件目)
    strFileName = Dir(strPathName & "\*.xls", vbNormal)
    If strFileName = "" Then
        MsgBox "このフォルダにはExcelワークブックは存在しません。"
        Exit Sub
    End If
    
    Dim files As Collection
    Set files = New Collection
    Dim vFileName As Variant
    Do While strFileName <> ""
        Call files.Add(strFileName)
        ' 次のファイル名を参照
        strFileName = Dir
    Loop
    'For Each vFileName In files
    
    Set xlAPP = Application
    With xlAPP
        .ScreenUpdating = False             ' 画面描画停止
        .EnableEvents = False               ' イベント動作停止
        .EnableCancelKey = xlErrorHandler   ' Escキーでエラートラップする
        .Cursor = xlWait                    ' カーソルを砂時計にする
    End With
    On Error GoTo Error_Handler
        
    ' 指定フォルダの全Excelワークブックについて繰り返す
    For Each vFileName In files
        ' Escキー打鍵判定
        DoEvents
        If swESC = True Then
            ' 中断するのかをメッセージで確認
            If MsgBox("中断キーが押されました。ここで終了しますか？", _
                vbInformation + vbYesNo) = vbYes Then
                GoTo Error_Handler
            Else
                swESC = False
            End If
        End If

        '-----------------------------------------------------------------------
        ' 検索した１ファイル単位の処理
        Call WB要素一覧作成処理(xlAPP, strPathName, CStr(vFileName), targetPrefix)
        
        '-----------------------------------------------------------------------
    Next
    
    GoTo Sub_EXIT
    
'----------------
' Escキー脱出用行ラベル
Error_Handler:
    
    If Err.Number = 18 Then
        Call LogErrorHandle("EscキーでのエラーRaise", "Doシート名から要素一覧作成", "")
        ' EscキーでのエラーRaise
        swESC = True
        Resume
    ElseIf Err.Number = 1004 Then
        ' 隠しシートや印刷対象なしの実行時エラーは無視
        Resume Next
    Else
        Call LogErrorHandle(Err.Description, "Doシート名から要素一覧作成", "")
        ' その他のエラーはメッセージ表示後終了
        MsgBox Err.Description
    End If

    Resume Sub_EXIT
'----------------
' 処理終了
Sub_EXIT:
    With xlAPP
        .StatusBar = False                  ' ステータスバーを復帰
        .EnableEvents = True                ' イベント動作再開
        .EnableCancelKey = xlInterrupt      ' Escキー動作を戻す
        .Cursor = xlDefault                 ' カーソルをﾃﾞﾌｫﾙﾄにする
        .ScreenUpdating = True              ' 画面描画再開
    End With
    Set xlAPP = Nothing
    Call LogEnd("Doシート名から要素一覧作成", "")
    MsgBox "終了しました"
End Sub


'*******************************************************************************
' １つのワークブックの処理
'*******************************************************************************
Private Sub WB要素一覧作成処理(xlAPP As Application, _
                            strPathName As String, _
                            strFileName As String, _
                            targetPrefix As String)
    On Error GoTo Error_Handler
    '---------------------------------------------------------------------------
    Dim objWBK As Workbook          ' ワークブックObject
    
    Call LogStart("WB要素一覧作成処理", "")
    ' ステータスバーに処理ファイル名を表示
    xlAPP.StatusBar = strFileName & " 処理中..."
    ' ワークブックを開く
    Set objWBK = OpenWorkBook(strPathName & cnsYEN & strFileName, False, False)
    If Not objWBK Is Nothing Then
    
        Call Do要素一覧作成処理(objWBK, targetPrefix)
        objWBK.Close SaveChanges:=True

    End If
    
    GoTo Sub_EXIT

Error_Handler:
    'Resume Sub_EXIT
    Err.Raise 5001
   
Sub_EXIT:
    Call LogEnd("WB要素一覧作成処理", "")
    xlAPP.StatusBar = False
    Set objWBK = Nothing

End Sub

'*******************************************************************************
' 要素一覧(型一覧)作成処理
'*******************************************************************************
Private Sub Do要素一覧作成処理(objWBK As Workbook, targetPrefix As String)
    '---------------------------------------------------------------------------
    Dim aSheet As Worksheet         ' ワークブックObject
    Dim elmListSheet As Worksheet
    
    Call LogStart("Do要素一覧作成処理", "")
    
    '要素一覧（型一覧）シート取得
    Set elmListSheet = GetTypeListSheet(objWBK)
    
    If elmListSheet Is Nothing Then
        Exit Sub
    End If
    '---------------------------------------------------------------------------
    ' 要素リストのエントリーを削除
    '---------------------------------------------------------------------------
    Dim i As Integer
    Dim maxNum As Integer: maxNum = elmListSheet.Cells(elmListSheet.Rows.count, 1).End(xlUp).row
    
    If maxNum < TEMPLATE_TYPE_LIST_SHEET_MAX_ROW Then
        maxNum = TEMPLATE_TYPE_LIST_SHEET_MAX_ROW
    End If
    For i = maxNum To 2 Step -1
        elmListSheet.Rows(i).Delete
    Next
    
    '---------------------------------------------------------------------------
    ' シート毎の処理
    '---------------------------------------------------------------------------
    i = 2
    For Each aSheet In objWBK.Sheets
        If InStr(aSheet.Name, targetPrefix) = 1 Then
            elmListSheet.Cells(i, COL_TYPELIST_LNAME).Value = GetTypeNameFromSheetName(aSheet, targetPrefix)
            elmListSheet.Cells(i, COL_TYPELIST_NO).Value = "=Row()-1"
            With elmListSheet
              .Range(.Cells(i, COL_TYPELIST_NO), .Cells(i, COL_TYPELIST_MAX_NO)).Borders.LineStyle = xlContinuous
            End With
            i = i + 1
        End If
    Next
    
    Call LogEnd("Do要素一覧作成処理", "")
End Sub

'*******************************************************************************
' 参照ファイル一覧作成処理
'*******************************************************************************
Private Sub Do参照ファイル一覧作成処理(objWBK As Workbook, list As Collection)
    '---------------------------------------------------------------------------
    Dim aSheet As Worksheet         ' ワークブックObject
    Dim refSheet As Worksheet
    
    Call LogStart("Do参照ファイル一覧作成処理", "")
    
    '参照ファイル一覧取得
    Set refSheet = GetSheet(objWBK, SHEET_NAME_REF_LIST_NEW)
    If refSheet Is Nothing Then
        Exit Sub
    End If
    
    '---------------------------------------------------------------------------
    ' 参照ファイル一覧のエントリーを削除
    '---------------------------------------------------------------------------
    Dim j As Integer
    For j = TEMPLATE_REF_LIST_SHEET_MAX_ROW To 2 Step -1
        refSheet.Rows(j).Delete
    Next
    
    '---------------------------------------------------------------------------
    ' シート毎の処理
    '---------------------------------------------------------------------------
    Dim rowIdx As Integer: rowIdx = 2
    Dim item As Variant
    For Each item In list
        refSheet.Cells(rowIdx, COL_REFLIST_NO).Value = "=Row()-1"
        refSheet.Cells(rowIdx, COL_REFLIST_FILE_NAME).Value = GenCommonTypeFileName(CStr(item))
        refSheet.Cells(rowIdx, COL_REFLIST_FILE_ID).Value = Mid(item, InStr(item, ":") + 1)
        With refSheet
          .Range(.Cells(rowIdx, COL_REFLIST_MIN_NO), .Cells(rowIdx, COL_REFLIST_MAX_NO)).Borders.LineStyle = xlContinuous
        End With
        rowIdx = rowIdx + 1
    Next
    Call LogEnd("Do参照ファイル一覧作成処理", "")
End Sub

'*******************************************************************************
' 共通型・設計書ファイル名の取得
'*******************************************************************************
Private Function GenCommonTypeFileName(item As String)
    GenCommonTypeFileName = "他システムインターフェース仕様書_IFASW" & _
                            Mid(item, InStr(item, ":") + 1) & _
                            "I_" & _
                            Mid(item, 1, InStr(item, ":") - 1) & ".xlsx"

End Function
'*******************************************************************************
' 要素一覧からシート作成
'*******************************************************************************
Public Sub 要素一覧からシート作成()
    Call LogStart("要素一覧からシート作成", "")
    
    Call debugLog(INIT_LOG, 0)
    Dim xlAPP As Application        ' Excel.Application
    Dim strPathName As String       ' 指定フォルダ名
    Dim strFileName As String       ' 検出したファイル名
    Dim swESC As Boolean            ' Escキー判定
    Dim templateWBK As Workbook     ' Excel.Workbook(テンプレートファイル)
    
    ' ｢フォルダの参照｣よりフォルダ名の取得(modFolderPicker1に収容)
    strPathName = FolderDialog("フォルダを指定して下さい", True)
    If strPathName = "" Then Exit Sub
    
    ' 指定フォルダ内のExcelワークブックのファイル名を参照する(1件目)
    strFileName = Dir(strPathName & "\*.xls", vbNormal)
    If strFileName = "" Then
        MsgBox "このフォルダにはExcelワークブックは存在しません。"
        Exit Sub
    End If
    
    Dim files As Collection
    Set files = New Collection
    Dim vFileName As Variant
    Do While strFileName <> ""
        Call files.Add(strFileName)
        ' 次のファイル名を参照
        strFileName = Dir
    Loop
    'For Each vFileName In files
    
    Set xlAPP = Application
    With xlAPP
        .ScreenUpdating = False             ' 画面描画停止
        .EnableEvents = False               ' イベント動作停止
        .EnableCancelKey = xlErrorHandler   ' Escキーでエラートラップする
        .Cursor = xlWait                    ' カーソルを砂時計にする
    End With
    On Error GoTo Error_Handler
    
    ' テンプレートワークブックを開く
    Set templateWBK = OpenWorkBook(TEMPLATE_FILE_PATH, False, True)

    If Not templateWBK Is Nothing Then
    
        ' 指定フォルダの全Excelワークブックについて繰り返す
        For Each vFileName In files
            ' Escキー打鍵判定
            DoEvents
            If swESC = True Then
                ' 中断するのかをメッセージで確認
                If MsgBox("中断キーが押されました。ここで終了しますか？", _
                    vbInformation + vbYesNo) = vbYes Then
                    GoTo Error_Handler
                Else
                    swESC = False
                End If
            End If
    
            '-----------------------------------------------------------------------
            ' 検索した１ファイル単位の処理
            Call WB型一覧シート作成処理(xlAPP, templateWBK, strPathName, CStr(vFileName))
            
            '-----------------------------------------------------------------------
        Next
        templateWBK.Close SaveChanges:=False
    
    End If
    
    GoTo Sub_EXIT
    
'----------------
' Escキー脱出用行ラベル
Error_Handler:
    If Err.Number = 18 Then
        ' EscキーでのエラーRaise
        swESC = True
        Resume
    ElseIf Err.Number = 1004 Then
        ' 隠しシートや印刷対象なしの実行時エラーは無視
        Resume Next
    Else
        Call LogErrorHandle(Err.Description, "要素一覧からシート作成", "")
        ' その他のエラーはメッセージ表示後終了
        MsgBox Err.Description
    End If

    Resume Sub_EXIT
'----------------
' 処理終了
Sub_EXIT:

    With xlAPP
        .StatusBar = False                  ' ステータスバーを復帰
        .EnableEvents = True                ' イベント動作再開
        .EnableCancelKey = xlInterrupt      ' Escキー動作を戻す
        .Cursor = xlDefault                 ' カーソルをﾃﾞﾌｫﾙﾄにする
        .ScreenUpdating = True              ' 画面描画再開
    End With
    Set xlAPP = Nothing
    Call LogEnd("要素一覧からシート作成", "")
    MsgBox "終了しました"
End Sub


'*******************************************************************************
' 型一覧シート作成_１ワークブックの処理
'*******************************************************************************
Private Sub WB型一覧シート作成処理(xlAPP As Application, _
                            templateWBK As Workbook, _
                            strPathName As String, _
                            strFileName As String)
    On Error GoTo Error_Handler
    '---------------------------------------------------------------------------
    Call LogStart("WB型一覧シート作成処理", strFileName)
    Dim objWBK As Workbook          ' ワークブックObject

    ' ステータスバーに処理ファイル名を表示
    xlAPP.StatusBar = strFileName & " 型一覧反映中"
    ' ワークブックを開く
    Set objWBK = OpenWorkBook(strPathName & cnsYEN & strFileName, False, False)
    
    If Not objWBK Is Nothing Then
    
        '要素一覧（型一覧）シート取得
        Dim elmListSheet As Worksheet
        Set elmListSheet = GetTypeListSheet(objWBK)
    
        Dim i As Integer, _
            typeName As String, _
            targetSheetName As String, _
            targetSheet As Worksheet
    
        For i = 2 To elmListSheet.Cells(elmListSheet.Rows.count, 1).End(xlUp).row
            '型名
            typeName = elmListSheet.Cells(i, COL_NEW_TYPEDEF_NAME)
            'シート名
            targetSheetName = TARGET_SHEET_PREFIX_NEW & typeName
    
            ' 存在しない場合にテンプレートから作成する。
            If Not ContainsSheet(objWBK, targetSheetName) Then
                ' コピー実行
                templateWBK.Worksheets(TEMPLATE_SHEET_NAME).Copy After:=objWBK.Worksheets(objWBK.Worksheets.count)
                objWBK.Worksheets(objWBK.Worksheets.count).Name = targetSheetName
                Set targetSheet = objWBK.Sheets(targetSheetName)
                If Not targetSheet Is Nothing Then
                    '---------------------------------------------------------------------------
                    ' 要素リストのエントリーを削除
                    '---------------------------------------------------------------------------
                    Dim j As Integer
                    For j = targetSheet.Cells(targetSheet.Rows.count, 1).End(xlUp).row To 2 Step -1
                        targetSheet.Rows(j).Delete
                    Next
                End If
            End If
        Next
    
        objWBK.Close SaveChanges:=True
    
    End If
    
    GoTo Sub_EXIT
    
Error_Handler:
    Call LogErrorHandle("5002", "WB型一覧シート作成処理", strFileName)
    'Resume Sub_EXIT
    Err.Raise 5002
   
Sub_EXIT:
    Call LogEnd("WB型一覧シート作成処理", strFileName)
    xlAPP.StatusBar = False
    Set objWBK = Nothing
    Set elmListSheet = Nothing
End Sub

'*******************************************************************************
' 型一覧シート作成_１ワークブックの処理
'*******************************************************************************
Private Sub WBデータベース項目(xlAPP As Application, _
                            templateWBK As Workbook, _
                            strPathName As String, _
                            strFileName As String)
    On Error GoTo Error_Handler
    '---------------------------------------------------------------------------
    Call LogStart("WB型一覧シート作成処理", strFileName)
    Dim objWBK As Workbook          ' ワークブックObject

    ' ステータスバーに処理ファイル名を表示
    xlAPP.StatusBar = strFileName & " 型一覧反映中"
    ' ワークブックを開く
    Set objWBK = OpenWorkBook(strPathName & cnsYEN & strFileName, False, False)
    
    If Not objWBK Is Nothing Then
    
        '要素一覧（型一覧）シート取得
        Dim elmListSheet As Worksheet
        Set elmListSheet = GetTypeListSheet(objWBK)
    
        Dim i As Integer, _
            typeName As String, _
            targetSheetName As String, _
            targetSheet As Worksheet
    
        For i = 2 To elmListSheet.Cells(elmListSheet.Rows.count, 1).End(xlUp).row
            '型名
            typeName = elmListSheet.Cells(i, COL_NEW_TYPEDEF_NAME)
            'シート名
            targetSheetName = TARGET_SHEET_PREFIX_NEW & typeName
    
            ' 存在しない場合にテンプレートから作成する。
            If Not ContainsSheet(objWBK, targetSheetName) Then
                ' コピー実行
                templateWBK.Worksheets(TEMPLATE_SHEET_NAME).Copy After:=objWBK.Worksheets(objWBK.Worksheets.count)
                objWBK.Worksheets(objWBK.Worksheets.count).Name = targetSheetName
                Set targetSheet = objWBK.Sheets(targetSheetName)
                If Not targetSheet Is Nothing Then
                    '---------------------------------------------------------------------------
                    ' 要素リストのエントリーを削除
                    '---------------------------------------------------------------------------
                    Dim j As Integer
                    For j = targetSheet.Cells(targetSheet.Rows.count, 1).End(xlUp).row To 2 Step -1
                        targetSheet.Rows(j).Delete
                    Next
                End If
            End If
        Next
    
        objWBK.Close SaveChanges:=True
    
    End If
    
    GoTo Sub_EXIT
    
Error_Handler:
    Call LogErrorHandle("5002", "WB型一覧シート作成処理", strFileName)
    'Resume Sub_EXIT
    Err.Raise 5002
   
Sub_EXIT:
    Call LogEnd("WB型一覧シート作成処理", strFileName)
    xlAPP.StatusBar = False
    Set objWBK = Nothing
    Set elmListSheet = Nothing
End Sub


'*******************************************************************************
' １つのワークブック中にシートが存在するかどうかの処理
'*******************************************************************************
Private Function ContainsSheet(targetWb As Workbook, sname As String) As Boolean
    Dim ws As Worksheet, flag As Boolean
    flag = False
    For Each ws In targetWb.Worksheets
        If ws.Name = sname Then
            flag = True
            Exit For
        End If
    Next ws
    ContainsSheet = flag
End Function

'*******************************************************************************
' 要素一覧シートの取得
' 新フォーマットの名称があればそちらを優先
'*******************************************************************************
Private Function GetTypeListSheet(targetWb As Workbook) As Worksheet
    Dim ws As Worksheet, retWs As Worksheet
    If ContainsSheet(targetWb, SHEET_NAME_TYPE_LIST_NEW) Then
        Set retWs = GetSheet(targetWb, SHEET_NAME_TYPE_LIST_NEW)
    Else
        Set retWs = GetSheet(targetWb, SHEET_NAME_TYPE_LIST_OLD)
    End If
    Set GetTypeListSheet = retWs
End Function

'*******************************************************************************
' ワークブックから指定の名前のワークシートを探して返却
'*******************************************************************************
Private Function GetSheet(targetWb As Workbook, targetSheet As String) As Worksheet
    Dim ws As Worksheet, retWs As Worksheet
    For Each ws In targetWb.Worksheets
        If ws.Name = targetSheet Then
            Set retWs = ws
            Exit For
        End If
    Next ws
    Set GetSheet = retWs
End Function

'*******************************************************************************
' ワークブックから集約シートを探す
'*******************************************************************************
Private Function GetSummarySheet(targetWb As Workbook) As Worksheet
    Dim ws As Worksheet, retWs As Worksheet
    For Each ws In targetWb.Worksheets
        If InStr(ws.Name, "3.1") = 1 Then
            Set retWs = ws
            Exit For
        End If
    Next ws
    Set GetSummarySheet = retWs
End Function

'*******************************************************************************
' 旧シート名から要素名を取得
'*******************************************************************************
Private Function GetTypeNameFromSheetName(targetSheet As Worksheet, targetPrefix As String) As String
    Dim retVal As String, pos As Long
    If InStr(targetSheet.Name, targetPrefix) = 1 Then
        '旧シート名
        If targetPrefix = TARGET_SHEET_PREFIX_OLD Then
            '最後の.(ドットの位置)
            pos = InStrRev(targetSheet.Name, ".")
            If pos > 0 Then
                ' 要素一覧シートに追記
                retVal = Mid(targetSheet.Name, pos + 1)
            End If
        End If

        '新シート名
        If targetPrefix = TARGET_SHEET_PREFIX_NEW Then
            '最後の.(ドットの位置)
            pos = InStrRev(targetSheet.Name, TARGET_SHEET_PREFIX_NEW)
            If pos = 1 Then
                ' 要素一覧シートに追記
                retVal = Mid(targetSheet.Name, Len(TARGET_SHEET_PREFIX_NEW) + 1)
            End If
        End If
    End If
    GetTypeNameFromSheetName = retVal
End Function

'*******************************************************************************
' 共通型抽出処理
'*******************************************************************************
Public Sub 共通型抽出処理()
    Call debugLog(INIT_LOG, 0)
    Dim xlAPP As Application        ' Excel.Application
    Dim strPathName As String       ' 指定フォルダ名
    Dim strFileName As String       ' 検出したファイル名
    Dim swESC As Boolean            ' Escキー判定
    Dim templateWBK As Workbook     ' Excel.Workbook(テンプレートファイル)
    Dim targetFiles As Collection   ' 対象ファイル一覧
    
    ' ｢フォルダの参照｣よりフォルダ名の取得(modFolderPicker1に収容)
    strPathName = FolderDialog("フォルダを指定して下さい", True)
    If strPathName = "" Then Exit Sub
    
    ' 指定フォルダ内のExcelワークブックのファイル名を参照する(1件目)
    strFileName = Dir(strPathName & "\*.xlsx", vbNormal)
    If strFileName = "" Then
        MsgBox "このフォルダにはExcelワークブックは存在しません。"
        Exit Sub
    End If
      
    Set xlAPP = Application
    With xlAPP
        .ScreenUpdating = False             ' 画面描画停止
        .EnableEvents = False               ' イベント動作停止
        .EnableCancelKey = xlErrorHandler   ' Escキーでエラートラップする
        .Cursor = xlWait                    ' カーソルを砂時計にする
    End With
    On Error GoTo Error_Handler
    
    Call LogStart("共通型抽出処理", "")
    ' テンプレートワークブックを開く
    Set templateWBK = OpenWorkBook(TEMPLATE_FILE_PATH, False, True)
    
    If Not templateWBK Is Nothing Then

        Set targetFiles = New Collection
        targetFiles.Add (strFileName)
        ' 指定フォルダの全Excelワークブックについて繰り返す
        Do While strFileName <> ""
            targetFiles.Add (strFileName)
            ' 次のファイル名を参照
            strFileName = Dir
        Loop
        
        ' 指定フォルダの全Excelワークブックについて繰り返す
        Dim aFile As Variant
        For Each aFile In targetFiles
            ' Escキー打鍵判定
            DoEvents
            If swESC = True Then
                ' 中断するのかをメッセージで確認
                If MsgBox("中断キーが押されました。ここで終了しますか？", _
                    vbInformation + vbYesNo) = vbYes Then
                    GoTo Error_Handler
                Else
                    swESC = False
                End If
            End If
            '-----------------------------------------------------------------------
            ' 検索した１ファイル単位の処理
            Call WB共通型抽出処理(xlAPP, templateWBK, strPathName, CStr(aFile))
            '-----------------------------------------------------------------------
    
        Next
    
        templateWBK.Close SaveChanges:=False
    
    End If

    GoTo Sub_EXIT
    
'----------------
' Escキー脱出用行ラベル
Error_Handler:
    If Err.Number = 18 Then
        ' EscキーでのエラーRaise
        swESC = True
        Resume
    ElseIf Err.Number = 1004 Then
        ' 隠しシートや印刷対象なしの実行時エラーは無視
        Resume Next
    Else
        Call LogErrorHandle(Err.Description, "共通型抽出処理", "")
        ' その他のエラーはメッセージ表示後終了
        MsgBox Err.Description
    End If
    Resume Sub_EXIT
'----------------
' 処理終了
Sub_EXIT:
    Call LogEnd("共通型抽出処理", "")
    With xlAPP
        .StatusBar = False                  ' ステータスバーを復帰
        .EnableEvents = True                ' イベント動作再開
        .EnableCancelKey = xlInterrupt      ' Escキー動作を戻す
        .Cursor = xlDefault                 ' カーソルをﾃﾞﾌｫﾙﾄにする
        .ScreenUpdating = True              ' 画面描画再開
    End With
    Set xlAPP = Nothing
    MsgBox "終了しました"
End Sub

'*******************************************************************************
' １つのワークブックの移行処理
'*******************************************************************************
Private Sub WB共通型抽出処理(xlAPP As Application, _
                            templateWBK As Workbook, _
                            strPathName As String, _
                            strFileName As String)
    
    On Error GoTo Error_Handler
    Call LogStart("WB共通型抽出処理", strFileName)
    '---------------------------------------------------------------------------
    Dim objWBK As Workbook          ' ワークブックObject

    ' ステータスバーに処理ファイル名を表示
    xlAPP.StatusBar = strFileName & " 処理中．．．"
    ' ワークブックを開く
    Set objWBK = OpenWorkBook(strPathName & cnsYEN & strFileName, False, False)
    
    If Not objWBK Is Nothing Then
    
        If Not ContainsSheet(objWBK, SHEET_NAME_COVER_NEW) Then
            objWBK.Close SaveChanges:=False
            GoTo Sub_EXIT
            Exit Sub
        End If
        
        '---------------------------------------------------------------------------
        ' シート毎に処理
        '---------------------------------------------------------------------------
        Call 移行_共通型定義(objWBK, strPathName)
        
        objWBK.Close SaveChanges:=True

    End If
    
    GoTo Sub_EXIT
    
Error_Handler:
    Call LogErrorHandle(Err.Description, "WB共通型抽出処理", strFileName)
    If Not objWBK Is Nothing Then
        objWBK.Close SaveChanges:=False
    End If

    '-- エラーが発生した場合
    'Resume Sub_EXIT
    Err.Raise 5003
   
Sub_EXIT:
    Call LogEnd("WB共通型抽出処理", strFileName)
    xlAPP.StatusBar = False
    Set objWBK = Nothing

End Sub

'*******************************************************************************
' 旧フォーマットから新フォーマットへ移行処理
'*******************************************************************************
Public Sub 移行()
    Call debugLog(INIT_LOG, 0)
    Dim xlAPP As Application        ' Excel.Application
    Dim strPathName As String       ' 指定フォルダ名
    Dim strFileName As String       ' 検出したファイル名
    Dim swESC As Boolean            ' Escキー判定
    Dim templateWBK As Workbook     ' Excel.Workbook(テンプレートファイル)
    Dim targetFiles As Collection   ' 対象ファイル一覧
    
    ' ｢フォルダの参照｣よりフォルダ名の取得(modFolderPicker1に収容)
    strPathName = FolderDialog("フォルダを指定して下さい", True)
    If strPathName = "" Then Exit Sub
    
    ' 指定フォルダ内のExcelワークブックのファイル名を参照する(1件目)
    strFileName = Dir(strPathName & "\*.xlsx", vbNormal)
    If strFileName = "" Then
        MsgBox "このフォルダにはExcelワークブックは存在しません。"
        Exit Sub
    End If

    Set xlAPP = Application
    With xlAPP
        .ScreenUpdating = False             ' 画面描画停止
        .EnableEvents = False               ' イベント動作停止
        .EnableCancelKey = xlErrorHandler   ' Escキーでエラートラップする
'        .Cursor = xlWait                    ' カーソルを砂時計にする
    End With
    On Error GoTo Error_Handler
 
    Call LogStart("移行", strFileName)
    
    ' テンプレートワークブックを開く
    Set templateWBK = OpenWorkBook(TEMPLATE_FILE_PATH, False, True)
    
    If Not templateWBK Is Nothing Then

        Set targetFiles = New Collection
        ' 指定フォルダの全Excelワークブックについて繰り返す
        Do While strFileName <> ""
            targetFiles.Add (strFileName)
            ' 次のファイル名を参照
            strFileName = Dir
        Loop
        
        ' 指定フォルダの全Excelワークブックについて繰り返す
        Dim aFile As Variant
        For Each aFile In targetFiles
            ' Escキー打鍵判定
            DoEvents
            If swESC = True Then
                ' 中断するのかをメッセージで確認
                If MsgBox("中断キーが押されました。ここで終了しますか？", _
                    vbInformation + vbYesNo) = vbYes Then
                    GoTo Error_Handler
                Else
                    swESC = False
                End If
            End If
            '-----------------------------------------------------------------------
            ' 検索した１ファイル単位の処理
            Call WB移行(xlAPP, templateWBK, strPathName, CStr(aFile))
            '-----------------------------------------------------------------------
    
        Next
    
        templateWBK.Close SaveChanges:=False
    
    End If
    
    GoTo Sub_EXIT
    
'----------------
' Escキー脱出用行ラベル
Error_Handler:
    If Err.Number = 18 Then
        ' EscキーでのエラーRaise
        swESC = True
        Resume
    ElseIf Err.Number = 1004 Then
        ' 隠しシートや印刷対象なしの実行時エラーは無視
        Resume Next
    Else
        Call LogErrorHandle(Err.Description, "移行", strFileName)
        ' その他のエラーはメッセージ表示後終了
        MsgBox Err.Description
    End If
    Resume Sub_EXIT
'----------------
' 処理終了
Sub_EXIT:
    Call LogEnd("移行", strFileName)
    With xlAPP
        .StatusBar = False                  ' ステータスバーを復帰
        .EnableEvents = True                ' イベント動作再開
        .EnableCancelKey = xlInterrupt      ' Escキー動作を戻す
        .Cursor = xlDefault                 ' カーソルをﾃﾞﾌｫﾙﾄにする
        .ScreenUpdating = True              ' 画面描画再開
    End With
    Set xlAPP = Nothing
    MsgBox "終了しました"
End Sub

'*******************************************************************************
' １つのワークブックの移行処理
'*******************************************************************************
Private Sub WB移行(xlAPP As Application, _
                            templateWBK As Workbook, _
                            strPathName As String, _
                            strFileName As String)
    
    On Error GoTo Error_Handler
    '---------------------------------------------------------------------------
    Dim objWBK As Workbook          ' ワークブックObject

    Call LogStart("WB移行", strFileName)
    
    ' ステータスバーに処理ファイル名を表示
    xlAPP.StatusBar = strFileName & " 処理中．．．"
    ' ワークブックを開く
    Set objWBK = OpenWorkBook(strPathName & cnsYEN & strFileName, False, False)
    
    If Not objWBK Is Nothing Then
        
        If Not ContainsSheet(objWBK, SHEET_NAME_INOUT_OLD) Then
            objWBK.Close SaveChanges:=False
            GoTo Sub_EXIT
            Exit Sub
        End If
        
        '---------------------------------------------------------------------------
        ' シート毎に移行
        '---------------------------------------------------------------------------
        Call 移行_表紙(objWBK, templateWBK)
        Call 移行_改訂履歴(objWBK, templateWBK)
        Call 移行_参照ファイル一覧(objWBK, templateWBK)
        Call 移行_共通情報(objWBK, templateWBK)
        Call 移行_型一覧(objWBK, templateWBK)
        Call 移行_入出力定義(objWBK, templateWBK)
        Call 移行_型シート(objWBK, templateWBK)
        
        '---------------------------------------------------------------------------
        ' 型一覧の生成
        '---------------------------------------------------------------------------
        Call Do要素一覧作成処理(objWBK, TARGET_SHEET_PREFIX_NEW)
            
        Call 移行_共通型定義(objWBK, strPathName)
        '---------------------------------------------------------------------------
        ' 集約シートの削除
        '---------------------------------------------------------------------------
        Call DiscardSheet(GetSummarySheet(objWBK))
    
        
        objWBK.Close SaveChanges:=True

    End If
    
    GoTo Sub_EXIT
    
Error_Handler:
    Call LogErrorHandle(Err.Description, "WB移行", strFileName)
    
    If Not objWBK Is Nothing Then
        objWBK.Close SaveChanges:=False
    End If

    '-- エラーが発生した場合
    'Resume Sub_EXIT
    Err.Raise 5003
   
Sub_EXIT:
    Call LogEnd("WB移行", strFileName)
    xlAPP.StatusBar = False
    Set objWBK = Nothing

End Sub

'*******************************************************************************
' 「表紙」の移行処理
'*******************************************************************************
Private Sub 移行_表紙(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("移行_表紙", targetWBK.Name)
    
    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_COVER_NEW
    Dim OLD_SHEET_NAME As String: OLD_SHEET_NAME = SHEET_NAME_COVER_OLD
    
    '-------------------
    '▼移行要否チェック
    '-------------------
    'シートが存在しているか
    'B21のセルに値があるかどうか
    Dim targetWS As Worksheet
    Set targetWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    If (Not targetWS Is Nothing) Then
        If (targetWS.Range("B21").Value <> "") Then
            Exit Sub
        End If
    End If
    
    '-------------------
    '▼移行実行
    '-------------------
    Dim templateWS As Worksheet, newWS As Worksheet
    Set templateWS = GetSheet(templateWBK, NEW_SHEET_NAME)
    Set targetWS = GetSheet(targetWBK, OLD_SHEET_NAME)
    '対象シートがない場合はoldシートを探す
    If targetWS Is Nothing Then
        Set targetWS = GetSheet(targetWBK, "old_" & OLD_SHEET_NAME)
        If targetWS Is Nothing Then
            MsgBox "移行元シートがありません"
            Err.Raise 6001
        End If
    End If
    
    'リネーム
    If targetWS.Name = OLD_SHEET_NAME Then
        targetWS.Name = "old_" & targetWS.Name
    End If
    
    '新シートが存在していたら削除
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    If Not newWS Is Nothing Then
        Application.DisplayAlerts = False
        newWS.Delete
        Application.DisplayAlerts = True
    End If

    'テンプレートからコピー
    templateWS.Copy Before:=Sheets(1)
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    '値の移行
    newWS.Range("B21").Value = "=CONCATENATE(""第"",TEXT(MAX(改訂履歴!A2:A1048576),""0.00""),""版 "")"
    newWS.Range("B23").Value = "=CONCATENATE(YEAR(MAX(改訂履歴!B2:B1048576)),""."",MONTH(MAX(改訂履歴!B2:B1048576)),""."",DAY(MAX(改訂履歴!B2:B1048576)))"
    newWS.Range("B16").Value = targetWS.Range("B16").Value
    newWS.Range("E16").Value = targetWS.Range("E16").Value

    '---------------------------------------------------------------------------
    ' シート破棄
    '---------------------------------------------------------------------------
    Call DiscardSheet(targetWS)
    
    Call LogEnd("移行_表紙", targetWBK.Name)

End Sub

'*******************************************************************************
' 「改訂履歴」の移行処理
'*******************************************************************************
Private Sub 移行_改訂履歴(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("移行_改訂履歴", targetWBK.Name)

    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_HISTORY_NEW
    Dim OLD_SHEET_NAME As String: OLD_SHEET_NAME = SHEET_NAME_HISTORY_OLD
    
    '-------------------
    '▼移行要否チェック
    '-------------------
    '条件なし
    
    '-------------------
    '▼移行実行
    '-------------------
    Dim targetWS As Worksheet
    Set targetWS = GetSheet(targetWBK, OLD_SHEET_NAME)
    '対象シートがない場合はエラー
    If targetWS Is Nothing Then
        MsgBox "移行元シートがありません"
        Err.Raise 6002
    End If
    
    '書式変更のみ
    targetWS.Range("A2:A30").NumberFormatLocal = "G/標準"
    
    Call LogEnd("移行_改訂履歴", targetWBK.Name)

       
End Sub

'*******************************************************************************
' 「参照ファイル一覧」の移行処理
'*******************************************************************************
Private Sub 移行_参照ファイル一覧(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("移行_参照ファイル一覧", targetWBK.Name)


    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_REF_LIST_NEW '"1.参照ファイル一覧"
    Dim newWS As Worksheet
    
    '-------------------
    '▼移行要否チェック
    '-------------------
    '「1.参照ファイル一覧」のシートがあれば実行しない
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    '対象シートがない場合にコピー
    If Not newWS Is Nothing Then
        Exit Sub
    End If

    '-------------------
    '▼移行実行
    '-------------------
    'テンプレートからコピー
    templateWBK.Worksheets(NEW_SHEET_NAME).Copy After:=GetSheet(targetWBK, SHEET_NAME_HISTORY_NEW)
    Set newWS = targetWBK.Worksheets(NEW_SHEET_NAME)
    '---------------------------------------------------------------------------
    ' 不要コメント削除
    '---------------------------------------------------------------------------
    newWS.Range("A1:D30").ClearComments
    '---------------------------------------------------------------------------
    ' サンプル削除
    '---------------------------------------------------------------------------
    newWS.Range("A2:D2").Value = ""

    '---------------------------------------------------------------------------
    ' 行の高さを調整する。
    '---------------------------------------------------------------------------
    Call AdjustRowHeight(newWS)

    Call LogEnd("移行_参照ファイル一覧", targetWBK.Name)

End Sub

'*******************************************************************************
' 「共通情報」の移行処理
'*******************************************************************************
Private Sub 移行_共通情報(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("移行_共通情報", targetWBK.Name)

    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_COMMON_NEW '"2.共通情報"
    Dim newWS As Worksheet
    
    '-------------------
    '▼移行要否チェック
    '-------------------
    '「2.共通情報」のシートがあれば実行しない
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    '対象シートがない場合にコピー
    If Not newWS Is Nothing Then
        Exit Sub
    End If

    '-------------------
    '▼移行実行
    '-------------------
    'テンプレートからコピー
    templateWBK.Worksheets(NEW_SHEET_NAME).Copy After:=targetWBK.Worksheets(SHEET_NAME_REF_LIST_NEW) '「1.参照ファイル一覧」の後ろ
    Set newWS = targetWBK.Worksheets(NEW_SHEET_NAME)
    '削除
    newWS.Range("C2").Value = ""
    
    Call LogEnd("移行_共通情報", targetWBK.Name)

End Sub

'*******************************************************************************
' 「型一覧」の移行処理
'*******************************************************************************
Private Sub 移行_型一覧(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("移行_型一覧", targetWBK.Name)
    
    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_TYPE_LIST_NEW '"3.型一覧"
    Dim OLD_SHEET_NAME As String: OLD_SHEET_NAME = SHEET_NAME_TYPE_LIST_OLD '"2.要素一覧"

    Dim newWS As Worksheet, oldWS As Worksheet
    
    '-------------------
    '▼移行要否チェック
    '-------------------
    '「3.型一覧」のシートがあれば実行しない
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    '対象シートがない場合にコピー
    If Not newWS Is Nothing Then
        Exit Sub
    End If
    
    '-------------------
    '▼移行実行
    '-------------------
    'テンプレートからコピー
    templateWBK.Worksheets(NEW_SHEET_NAME).Copy After:=targetWBK.Worksheets(SHEET_NAME_COMMON_NEW) '「2.共通情報」の後ろ
    Set newWS = targetWBK.Worksheets(NEW_SHEET_NAME)

    '---------------------------------------------------------------------------
    ' 要素リストのエントリーを削除
    '---------------------------------------------------------------------------
    Dim i As Integer
    For i = newWS.Cells(newWS.Rows.count, 1).End(xlUp).row To 2 Step -1
        newWS.Rows(i).Delete
    Next
    
    '---------------------------------------------------------------------------
    ' 不要コメント削除
    '---------------------------------------------------------------------------
    newWS.Range("B1:F6").ClearComments
                    
    '---------------------------------------------------------------------------
    ' 行の高さを調整する。
    '---------------------------------------------------------------------------
    Call AdjustRowHeight(newWS)
                    
    '---------------------------------------------------------------------------
    ' シート破棄
    '---------------------------------------------------------------------------
    Call DiscardSheet(GetSheet(targetWBK, OLD_SHEET_NAME))
    
    Call LogEnd("移行_型一覧", targetWBK.Name)
    
End Sub

'*******************************************************************************
' 対象シートの行の高さを調整する。
'*******************************************************************************
Private Sub AdjustRowHeight(aSheet As Worksheet)
    Dim cmd As CmdRowAdjust
    Set cmd = New CmdRowAdjust
    Call cmd.ExecCommand(aSheet)
End Sub

'*******************************************************************************
' 「入出力定義」の移行処理
'*******************************************************************************
Private Sub 移行_入出力定義(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("移行_入出力定義", targetWBK.Name)
    
    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_INOUT_NEW  '"4.入出力定義"
    Dim OLD_SHEET_NAME As String: OLD_SHEET_NAME = SHEET_NAME_INOUT_OLD  '"1.入出力定義"
    Dim newWS As Worksheet, oldWS As Worksheet
    
    '-------------------
    '▼移行要否チェック
    '-------------------
    '対象シートが既にある場合は終了
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)
    If Not newWS Is Nothing Then
        Exit Sub
    End If
    
    '元となるシートがない場合は終了
    Set oldWS = GetSheet(targetWBK, OLD_SHEET_NAME)
    If oldWS Is Nothing Then
        Exit Sub
    End If
    
    '-------------------
    '▼移行実行
    '-------------------
    'テンプレートからコピー
    templateWBK.Worksheets(NEW_SHEET_NAME).Copy After:=targetWBK.Worksheets(SHEET_NAME_TYPE_LIST_NEW) '「3.型一覧」の後ろ
    '新シートを取得
    Set newWS = GetSheet(targetWBK, NEW_SHEET_NAME)

    '---------------------------------------------------------------------------
    ' 新シートの入出力定義のエントリーを削除
    '---------------------------------------------------------------------------
    Dim i As Integer
    For i = newWS.Cells(newWS.Rows.count, 1).End(xlUp).row To 2 Step -1
        newWS.Rows(i).Delete
    Next
    
    Dim j As Integer
    For j = 2 To oldWS.Cells(oldWS.Rows.count, 1).End(xlUp).row
        newWS.Cells(j, COL_NEW_IO_NO).Value = "=ROW()-1"
        newWS.Cells(j, COL_NEW_IO_IO).Value = oldWS.Cells(j, COL_OLD_IO_IO).Value
        newWS.Cells(j, COL_NEW_IO_ROOT_ELEMENT_ID).Value = oldWS.Cells(j, COL_OLD_IO_ROOT_ELEMENT_ID).Value
        newWS.Cells(j, COL_NEW_IO_ROOT_ELEMENT_NAME).Value = oldWS.Cells(j, COL_OLD_IO_ELEMENT_NAME).Value
        newWS.Cells(j, COL_NEW_IO_TYPE_NAME).Value = oldWS.Cells(j, COL_OLD_IO_ELEMENT_NAME).Value
        With newWS
          .Range(.Cells(i, COL_NEW_IO_MIN_NO), .Cells(i, COL_NEW_IO_MAX_NO)).Borders.LineStyle = xlContinuous
        End With
    Next

    '---------------------------------------------------------------------------
    ' シート破棄
    '---------------------------------------------------------------------------
    Call DiscardSheet(oldWS)
    
    Call LogEnd("移行_入出力定義", targetWBK.Name)

End Sub

'*******************************************************************************
' 「型シート」の移行処理
'*******************************************************************************
Private Sub 移行_型シート(targetWBK As Workbook, templateWBK As Workbook)

    Call LogStart("移行_型シート", targetWBK.Name)
    
    Dim NEW_SHEET_NAME As String: NEW_SHEET_NAME = SHEET_NAME_REF_LIST_NEW '"1.参照ファイル一覧"
    
    '-------------------
    '▼移行要否チェック
    '-------------------
    '旧の名称のシートが存在する場合に実施する。
    Dim exists As Boolean: exists = False
    Dim aSheet As Worksheet
    For Each aSheet In targetWBK.Sheets
        If InStr(aSheet.Name, TARGET_SHEET_PREFIX_OLD) = 1 Then
            exists = True
            Exit For
        End If
    Next
    
    If Not exists Then
        Exit Sub
    End If
    
    '-------------------
    '▼移行実行
    '-------------------
    Dim newWS As Worksheet, oldWS As Worksheet
    
    Dim pos As Integer, typeName As String, newSheetName As String
    Dim discardSheets As Collection
    Set discardSheets = New Collection
    
    '旧の名称のシートを順次処理
    For Each oldWS In targetWBK.Sheets
        If InStr(oldWS.Name, TARGET_SHEET_PREFIX_OLD) = 1 And oldWS.Name <> SHEET_NAME_INOUT_NEW Then
            '最後の.(ドットの位置)
            pos = InStrRev(oldWS.Name, ".")
            If pos > 0 Then
                ' 要素一覧シートに追記
                typeName = Mid(oldWS.Name, pos + 1)
                newSheetName = TARGET_SHEET_PREFIX_NEW & typeName
                ' 存在しない場合にテンプレートから作成する。
                If Not ContainsSheet(targetWBK, newSheetName) Then
                    ' コピー実行
                    templateWBK.Worksheets(TEMPLATE_SHEET_NAME).Copy After:=targetWBK.Worksheets(targetWBK.Worksheets.count)
                    targetWBK.Worksheets(targetWBK.Worksheets.count).Name = newSheetName
                    Set newWS = targetWBK.Sheets(newSheetName)

                    '---------------------------------------------------------------------------
                    ' サンプル削除
                    '---------------------------------------------------------------------------
                    With newWS.Range("A2:P30")
                        .Value = ""             '値削除
                        .ClearComments          'コメント削除
                    End With
  
                    '---------------------------------------------------------------------------
                    ' データ移行
                    '---------------------------------------------------------------------------
                    Dim j As Integer
                    For j = 2 To oldWS.Cells(oldWS.Rows.count, 1).End(xlUp).row
                        
                        newWS.Cells(j, COL_NEW_TYPEDEF_NO).Value = "=ROW()-1"    '項番
                        newWS.Cells(j, COL_NEW_TYPEDEF_NAME).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_NAME).Value   '項目名
                        newWS.Cells(j, COL_NEW_TYPEDEF_ID).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_ID).Value   '項目ID
                        newWS.Cells(j, COL_NEW_TYPEDEF_EXPLAIN).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_EXPLAIN).Value   '説明
                        newWS.Cells(j, COL_NEW_TYPEDEF_MIN_LOOP).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_MIN_LOOP).Value   '最小繰り返し
                        newWS.Cells(j, COL_NEW_TYPEDEF_MIN_LOOP).HorizontalAlignment = xlRight
                        '最大繰り返し
                        If oldWS.Cells(j, COL_OLD_TYPEDEF_MAX_LOOP).Value = "n" Or oldWS.Cells(j, COL_OLD_TYPEDEF_MAX_LOOP).Value = "ｎ" Then
                            newWS.Cells(j, COL_NEW_TYPEDEF_MAX_LOOP).Value = "*"
                        Else
                            newWS.Cells(j, COL_NEW_TYPEDEF_MAX_LOOP).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_MAX_LOOP).Value
                        End If
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAX_LOOP).HorizontalAlignment = xlRight
                        
                        
                        newWS.Cells(j, COL_NEW_TYPEDEF_COND_REQ).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_COND_REQ).Value   '必須条件
                        
                        '型
                        If oldWS.Cells(j, COL_OLD_TYPEDEF_TYPE).Value = "数値" Then
                            newWS.Cells(j, COL_NEW_TYPEDEF_TYPE).Value = "整数"
                        Else
                            newWS.Cells(j, COL_NEW_TYPEDEF_TYPE).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_TYPE).Value
                        End If
                        
                        '型名
                        If oldWS.Cells(j, COL_OLD_TYPEDEF_TYPE).Value = "要素" Then
                            newWS.Cells(j, COL_NEW_TYPEDEF_TYPE_NAME).Value = oldWS.Cells(j, 2).Value
                        Else
                            newWS.Cells(j, COL_NEW_TYPEDEF_TYPE_NAME).Value = ""
                        End If
                        
                        newWS.Cells(j, COL_NEW_TYPEDEF_CHAR_FORMAT).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_TYPE_DETAIL).Value  '文字種・書式
                        newWS.Cells(j, COL_NEW_TYPEDEF_MINLENGTH).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_MINLENGTH).Value  '最小桁数
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAXLENGTH).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_MAXLENGTH).Value  '最大桁数
                        newWS.Cells(j, COL_NEW_TYPEDEF_DATA_SAMPLE).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_DATA_SAMPLE).Value  'データ例
                        newWS.Cells(j, COL_NEW_TYPEDEF_BIKO).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_BIKO).Value  '備考
                        '欄外も4列程度コピー
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAX_NO + 1).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_BIKO + 1).Value
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAX_NO + 2).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_BIKO + 2).Value
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAX_NO + 3).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_BIKO + 3).Value
                        newWS.Cells(j, COL_NEW_TYPEDEF_MAX_NO + 4).Value = oldWS.Cells(j, COL_OLD_TYPEDEF_BIKO + 4).Value
                        
                        '罫線
                        With newWS
                            .Range(.Cells(j, COL_NEW_TYPEDEF_MIN_NO), .Cells(j, COL_NEW_TYPEDEF_MAX_NO)).Borders.LineStyle = xlContinuous
                        End With
                    Next
                    '---------------------------------------------------------------------------
                    ' 不要な行の削除
                    '---------------------------------------------------------------------------
                    If newWS.Cells(newWS.Rows.count, 1).End(xlUp).row < TEMPLATE_TYPE_DEF_SHEET_MAX_ROW Then
                        For j = TEMPLATE_TYPE_LIST_SHEET_MAX_ROW To newWS.Cells(newWS.Rows.count, 1).End(xlUp).row + 1 Step -1
                            newWS.Rows(j).Delete
                        Next
                    End If
                    
                    Dim cmd As CmdRowAdjust
                    Set cmd = New CmdRowAdjust
                    Call cmd.ExecCommand(newWS)
                    Call discardSheets.Add(oldWS)
                    
                End If
            End If
        End If
    Next
    '
    '---------------------------------------------------------------------------
    ' 旧シートの破棄
    '---------------------------------------------------------------------------
    For Each oldWS In discardSheets
        Call DiscardSheet(oldWS)
    Next
    
    Call LogEnd("移行_型シート", targetWBK.Name)
    
End Sub

'*******************************************************************************
' 旧シートの破棄処理
'*******************************************************************************
Private Sub DiscardSheet(oldWS As Worksheet)
    If Not oldWS Is Nothing Then
        If DESTRUCTION_MODE = DESTRUCTION_MODE_DELETE Then
            'シート削除
            Application.DisplayAlerts = False
            oldWS.Delete
            Application.DisplayAlerts = True
            
        ElseIf DESTRUCTION_MODE = DESTRUCTION_MODE_RENAME Then
            'シート名を変更
            oldWS.Name = "old_" & oldWS.Name
        End If
    End If
End Sub

'*******************************************************************************
' 共通型定義の移行処理
'*******************************************************************************
Private Function GetCommonTypes() As Collection
    
    Dim defSheet As Worksheet, typeName As String, typeId As String
    Set defSheet = ThisWorkbook.Worksheets(SHEET_NAME_TYPEDEF)
    
    Set GetCommonTypes = New Collection
    
    Dim i As Integer
    For i = 3 To defSheet.Cells(defSheet.Rows.count, 3).End(xlUp).row
        typeId = defSheet.Cells(i, 2)
        typeName = defSheet.Cells(i, 3)
        If typeId <> "" And typeName <> "" Then
            GetCommonTypes.Add (typeName & ":" & typeId)
        End If
    Next
End Function

'*******************************************************************************
' 共通型定義の移行処理
'*******************************************************************************
Private Sub 移行_共通型定義(targetWBK As Workbook, strPathName As String)

    Call LogStart("移行_共通型定義", targetWBK.Name)
    
    Dim CommonTypes As Collection
    Set CommonTypes = GetCommonTypes()
    
    Dim aSheet As Worksheet
    ' 一覧作成用参照型定義リスト
    Dim refTypeNames As Collection
    Set refTypeNames = New Collection
    
    Dim changed As Boolean: changed = False

    '-------------------
    '▼移行実行
    '-------------------

    '---------------------------------------------------------------------------
    '　各型定義シート内の共通型の部分にファイルIDを設定する処理
    '---------------------------------------------------------------------------
    Call Do参照型定義抽出処理(targetWBK, refTypeNames, CommonTypes, strPathName)
    
    If refTypeNames.count > 0 Then
        '---------------------------------------------------------------------------
        '　参照ファイル一覧の作成処理
        '---------------------------------------------------------------------------
        Call Do参照ファイル一覧作成処理(targetWBK, refTypeNames)
          
        '---------------------------------------------------------------------------
        '　共通型となるシートを別名保存する処理
        '---------------------------------------------------------------------------
        Call Do共通型別名保存処理(targetWBK, refTypeNames, strPathName)
    
        '---------------------------------------------------------------------------
        '　ID違いのファイルがある場合にリネームする処理
        '---------------------------------------------------------------------------
        Call Do共通型リネーム処理(targetWBK, refTypeNames, strPathName)
        
    End If
    
    
    Call LogEnd("移行_共通型定義", targetWBK.Name)
    
        
End Sub

'*******************************************************************************
' 参照型定義抽出処理
'*******************************************************************************
Sub Do参照型定義抽出処理(targetWBK As Workbook, refTypeNames As Collection, CommonTypes As Collection, strPathName As String)
    Dim aSheet As Worksheet
    If targetWBK Is Nothing Then
        Exit Sub
    End If
    For Each aSheet In targetWBK.Sheets

        If InStr(aSheet.Name, TARGET_SHEET_PREFIX_NEW) = 1 Then
            '---------------------------------------------------------------------------
            ' 行毎の処理
            '---------------------------------------------------------------------------
            Dim j As Integer
            Dim refTypeName As String, itm As Variant, bufWb As Workbook
            For j = 2 To aSheet.Cells(aSheet.Rows.count, 1).End(xlUp).row
                refTypeName = aSheet.Cells(j, COL_NEW_TYPEDEF_TYPE_NAME)
                If Len(refTypeName) > 0 Then
                    Dim exists As Boolean: exists = False
                    
                    For Each itm In CommonTypes
                        '一致するものがあればIDを記入する。
                        If InStr(CStr(itm), refTypeName & ":") = 1 Then
                            exists = True
                            'セルの値が異なる場合のみ更新
                            Call UpdateCell(aSheet, j, COL_NEW_TYPEDEF_FILE_ID, Mid(itm, Len(refTypeName) + 2))

                            If ContainsItem(refTypeNames, CStr(itm)) = False Then
                                Call refTypeNames.Add(itm)
                                If GetSheet(targetWBK, TARGET_SHEET_PREFIX_NEW & refTypeName) Is Nothing Then
                                    Set bufWb = OpenWorkBook(strPathName & cnsYEN & GenCommonTypeFileName(CStr(itm)), False, True, False)
                                    If Not bufWb Is Nothing Then
                                        Call Do参照型定義抽出処理(bufWb, refTypeNames, CommonTypes, strPathName)
                                        bufWb.Close SaveChanges:=False
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If Not exists Then
                        'セルの値が異なる場合のみ更新
                        Call UpdateCell(aSheet, j, COL_NEW_TYPEDEF_FILE_ID, "")
                    End If

                    If aSheet.Cells(j, COL_NEW_TYPEDEF_FILE_ID) = "" Then
                        If GetSheet(targetWBK, TARGET_SHEET_PREFIX_NEW & refTypeName) Is Nothing Then
                            Call LogErrorHandle("シート「" & aSheet.Name & "」に記載の型「" & refTypeName & "」は対象のシートが存在しません。", "", "")
                        End If
                    End If
                End If
            Next
            
            Dim typeWb As Workbook
            For Each itm In CommonTypes
                If InStr(CStr(itm), Mid(aSheet.Name, 5) & ":") = 1 Then
                    If ContainsItem(refTypeNames, CStr(itm)) = False Then
                        Call refTypeNames.Add(itm)
                    End If
                End If
            Next
        End If
    
    Next

End Sub

'*******************************************************************************
' 更新処理
'*******************************************************************************
Sub UpdateCell(aSheet As Worksheet, rowId As Integer, colId As Integer, val As Variant)
    If CStr(aSheet.Cells(rowId, colId)) <> CStr(val) Then
        aSheet.Cells(rowId, colId) = val
    End If
End Sub


'*******************************************************************************
' コレクション中にアイテムが含まれるかどうか
'*******************************************************************************
Private Function ContainsItem(list As Collection, item As String) As Boolean
    Dim t As Variant, flag As Boolean: flag = False
    
    For Each t In list
        If CStr(t) = item Then
            flag = True
            Exit For
        End If
    Next
    ContainsItem = flag
End Function

'*******************************************************************************
' 共通型のシートを別名保存する
'*******************************************************************************
Private Sub Do共通型別名保存処理(objWBK As Workbook, list As Collection, strPathName As String)
    '---------------------------------------------------------------------------
    Dim item As Variant
    
    For Each item In list
        Call 別名保存処理( _
            objWBK, _
            GetSheet(objWBK, TARGET_SHEET_PREFIX_NEW & Mid(CStr(item), 1, InStr(CStr(item), ":") - 1)), _
            GenCommonTypeFileName(CStr(item)), _
            strPathName)
    Next
    
End Sub

'*******************************************************************************
' 共通型のシートを別名保存する
'*******************************************************************************
Private Sub 別名保存処理(objWBK As Workbook, tgtSheet As Worksheet, fileName As String, strPathName As String)

    Call LogStart("別名保存処理", objWBK.Name & ":" & fileName)
    
    '参照型定義一覧の型が処理中のBookに存在した場合に、別名保存処理を実施する。
    If tgtSheet Is Nothing Then
        Exit Sub
    End If
        
    'パスを取得する ※２
    Dim filePath As String
    Dim testFileName As String: testFileName = fileName
    Dim pos As Integer
    Dim counter As Integer: counter = 1
    Dim newWorkBook As String
    
    filePath = strPathName & "\" & testFileName
    
    '既にファイルが存在していたら処理を行わない。
    'ブック名を取得
    '最後の_(アンダースコアの位置)
    Dim bookName As String
    pos = InStrRev(objWBK.Name, "_")
    If pos > 0 Then
        bookName = Mid(objWBK.Name, pos + 1)
    End If
        
    If Not Dir(filePath) <> "" Then
        'Workbooks.Add
        'newWorkBook = ActiveWorkbook.Name
        'tgtSheet.Copy Before:=Workbooks(newWorkBook).Sheets(1)
        tgtSheet.Copy
        ActiveWorkbook.SaveAs fileName:=filePath        '別名を付けてブックを保存する
        ActiveWorkbook.Close                            '別名ブックを閉じる
        Call debugLog(bookName & ":共通シートコピー作成：" & filePath, OUTPUT_LOG)
    End If

    'book固有ファイルを作成
    Dim filePath2 As String
    pos = InStrRev(fileName, ".")
    filePath2 = strPathName & "\" & Mid(fileName, 1, pos - 1) & "_" & bookName
    tgtSheet.Move
    ActiveWorkbook.SaveAs fileName:=filePath2
    ActiveWorkbook.Close
    Call debugLog(bookName & ":共通シート作成：" & filePath2, OUTPUT_LOG)

    Call LogEnd("別名保存処理", objWBK.Name & ":" & fileName)
    
End Sub

'*******************************************************************************
' 共通型でID違いのファイルがあればリネームする。
'*******************************************************************************
Private Sub Do共通型リネーム処理(objWBK As Workbook, CommonTypes As Collection, strPathName As String)

    Dim buf As String, item As Variant, searchPath As String
    For Each item In CommonTypes
        searchPath = "他システムインターフェース仕様書_IFASW*" & _
                            "I_" & _
                            Mid(item, 1, InStr(item, ":") - 1) & ".xlsx"
                            
        buf = Dir(strPathName & cnsYEN & searchPath, vbNormal)
        If buf <> "" Then
            If buf <> GenCommonTypeFileName(CStr(item)) Then
                Dim buf2 As String
                buf2 = Dir(strPathName & cnsYEN & GenCommonTypeFileName(CStr(item)))
                If buf2 = "" Then
                    Call debugLog("リネーム:「" & buf & "」=>「" & GenCommonTypeFileName(CStr(item)) & "」", OUTPUT_LOG)
                    'リネーム
                    Name strPathName & cnsYEN & buf As strPathName & cnsYEN & GenCommonTypeFileName(CStr(item))
                Else
                    Call debugLog(buf & "は不要なファイルです。" & buf2 & "が存在します。", OUTPUT_LOG)
                End If
            End If
        End If
CONTINUE:
    Next
End Sub

'*******************************************************************************
' 指定のワークブックをオープンする。
'*******************************************************************************
Function OpenWorkBook(filePath As String, updateLinks_ As Boolean, readOnly_ As Boolean, Optional msgFlg As Boolean) As Workbook
    If IsMissing(msgFlg) Then
        msgFlg = True
    End If
    Dim buf As String, wb As Workbook
    ''ファイルの存在チェック
    buf = Dir(filePath)
    If buf = "" Then
        If msgFlg Then
            MsgBox filePath & vbCrLf & "は存在しません", vbExclamation
        End If
        Exit Function
    End If
    ''同名ブックのチェック
    For Each wb In Workbooks
        If wb.Name = buf Then
            If msgFlg Then
                MsgBox buf & vbCrLf & "はすでに開いています", vbExclamation
            End If
            Exit Function
        End If
    Next wb
    
    Set OpenWorkBook = Workbooks.Open(fileName:=filePath, _
                                updateLinks:=updateLinks_, _
                                readOnly:=readOnly_)
End Function
