Attribute VB_Name = "RunCommand"
'*******************************************************************************
'  ワークブック順次処理
'*******************************************************************************
Option Explicit

Private Const cnsYEN = "\"

' コマンド定義
Const CMD_CODE_DEF As Byte = &H1
Const CMD_AUT_HEIGHT As Byte = &H2
Const CMD_AUT_GEN As Byte = &H4

'コード定義反映実行フラグ
Public execCmd As Byte
    
'*******************************************************************************
' 「」ボタン押下時処理
'*******************************************************************************
Sub Button1_Click()
    execCmd = CMD_CODE_DEF
    Call ProcMain
End Sub

'*******************************************************************************
' 「一括調整」ボタン押下時処理
'*******************************************************************************
Sub Button2_Click()
    execCmd = CMD_CODE_DEF Or _
                CMD_AUT_HEIGHT
    Call ProcMain
End Sub

'*******************************************************************************
' 「DB項目との整合性合わせ」ボタン押下時処理
'*******************************************************************************
Sub AutoGen_Click()
    execCmd = CMD_AUT_GEN Or _
                CMD_AUT_HEIGHT
    Call ProcMain
End Sub

'*******************************************************************************
' メイン処理
'*******************************************************************************
Sub ProcMain()

    Dim xlAPP As Application        ' Excel.Application
    Dim strPathName As String       ' 指定フォルダ名
    Dim strFileName As String       ' 検出したファイル名
    Dim swESC As Boolean            ' Escキー判定
    
    Call debugLog(INIT_LOG, 0)
    
    ' ｢フォルダの参照｣よりフォルダ名の取得(modFolderPicker1に収容)
    strPathName = FolderDialog("フォルダを指定して下さい", True)
    If strPathName = "" Then Exit Sub
    
    ' 指定フォルダ内のExcelワークブックのファイル名を参照する(1件目)
    strFileName = Dir(strPathName & "\*.xls", vbNormal)
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
    On Error GoTo Button1_Click_ESC
    
    Dim commands As Collection
    Set commands = GetCommands()
        
    ' 指定フォルダの全Excelワークブックについて繰り返す
    Do While strFileName <> ""
        ' Escキー打鍵判定
        DoEvents
        If swESC = True Then
            ' 中断するのかをメッセージで確認
            If MsgBox("中断キーが押されました。ここで終了しますか？", _
                vbInformation + vbYesNo) = vbYes Then
                GoTo Button1_Click_EXIT
            Else
                swESC = False
            End If
        End If

        '-----------------------------------------------------------------------
        ' 検索した１ファイル単位の処理
        Call OneWorkbookProc(xlAPP, commands, strPathName, strFileName)
        
        '-----------------------------------------------------------------------
        ' 次のファイル名を参照
        strFileName = Dir
    Loop
    GoTo Button1_Click_EXIT
    
'----------------
' Escキー脱出用行ラベル
Button1_Click_ESC:
    If Err.Number = 18 Then
        ' EscキーでのエラーRaise
        swESC = True
        Resume
    ElseIf Err.Number = 1004 Then
        ' 隠しシートや印刷対象なしの実行時エラーは無視
        Resume Next
    Else
        ' その他のエラーはメッセージ表示後終了
        MsgBox Err.Description
    End If

'----------------
' 処理終了
Button1_Click_EXIT:
    MsgBox "終了しました"
    With xlAPP
        .StatusBar = False                  ' ステータスバーを復帰
        .EnableEvents = True                ' イベント動作再開
        .EnableCancelKey = xlInterrupt      ' Escキー動作を戻す
        .Cursor = xlDefault                 ' カーソルをﾃﾞﾌｫﾙﾄにする
        .ScreenUpdating = True              ' 画面描画再開
    End With
    Set xlAPP = Nothing
End Sub

'*******************************************************************************
' 実行コマンドリストの生成
'*******************************************************************************
Private Function GetCommands() As Collection
    '------------------------------------------------
    ' コマンド生成
    ' シートに対する処理コマンドのインスタンスを生成する。
    ' 各処理コマンドに割り当てられたビット値によって、該当コマンドのインスタンスを生成する。
    '------------------------------------------------
    Dim commands As Collection      ' コマンドリスト
    Set commands = New Collection

    If (execCmd And CMD_CODE_DEF) > 0 Then
        Call commands.Add(New CmdCodeDef)
    End If
    If (execCmd And CMD_AUT_GEN) > 0 Then
        Call commands.Add(New CmdAutoGen)
    End If
    If (execCmd And CMD_AUT_HEIGHT) > 0 Then
        Call commands.Add(New CmdRowAdjust)
    End If
    
    Set GetCommands = commands
End Function


'*******************************************************************************
' １つのワークブックの処理
'*******************************************************************************
Private Sub OneWorkbookProc(xlAPP As Application, _
                            commands As Collection, _
                            strPathName As String, _
                            strFileName As String)
    '---------------------------------------------------------------------------
    Dim objWBK As Workbook          ' ワークブックObject
    Dim aSheet As Worksheet          ' ワークブックObject
    ' ステータスバーに処理ファイル名を表示
    xlAPP.StatusBar = strFileName & " コード定義反映中"
    ' ワークブックを開く
    Set objWBK = Workbooks.Open(fileName:=strPathName & cnsYEN & strFileName, _
                                UpdateLinks:=False, _
                                ReadOnly:=False)
    '---------------------------------------------------------------------------
    ' シート毎の処理
    '---------------------------------------------------------------------------
    For Each aSheet In objWBK.Sheets
        Call OneSheetProc(commands, aSheet)
    Next

    objWBK.Close SaveChanges:=True
    xlAPP.StatusBar = False
    Set objWBK = Nothing
End Sub


'*******************************************************************************
' １つのシートの処理
'*******************************************************************************
Private Sub OneSheetProc(commands As Collection, _
                            targetSheet As Worksheet)
    
    Dim cmd As Variant
    For Each cmd In commands
        Call cmd.ExecCommand(targetSheet)
    Next
    
End Sub


                                                      

