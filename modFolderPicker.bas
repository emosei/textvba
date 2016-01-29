Attribute VB_Name = "modFolderPicker"
'*******************************************************************************
'   ファイル、フォルダ参照のダイアログ処理(Application.FileDialog)
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'*******************************************************************************
'変更日付 Rev  変更履歴内容---------------------------------------------------->
'14/04/14(1.00)新規作成
'*******************************************************************************
Option Explicit
Private Const MAX_PATH As Long = 260
Private Const g_cnsYen As String = "\"
Private Const g_cnsCol As String = ":"
' ■ローカルドライブからマウントされているネットワークリソース名を取得する
Private Declare Function WNetGetConnection Lib "MPR.dll" _
    Alias "WNetGetConnectionA" _
    (ByVal lpszLocalName As String, _
     ByVal lpszRemoteName As String, _
     cbRemoteName As Long) As Long

'*******************************************************************************
'   「フォルダの参照」ダイアログを表示させ、選択したフォルダ名を返す
'-------------------------------------------------------------------------------
'   引渡値 = ①ウィンドウタイトル
'   　　　   ②ネットワークドライブ→ネットワークリソースの置換区分(Option)
'   　　　   ③ルートフォルダ(Option)
'            ④ルートフォルダ固定スイッチ
'              (Option, 1=固定する, 2=ルートに戻す, 3=キャンセル時は初期化)
'   　　　   ⑤初期化用ルートフォルダ(Option)
'            ⑥ボタン表示名(Option)
'   戻り値 = フォルダ名(フルパスで右\なし、未選択時はブランク)
'-------------------------------------------------------------------------------
' 　作成日：2014年04月14日
' 　作成者：井上 治
' 　更新日：2014年04月14日
' 　更新者：井上 治
'*******************************************************************************
Public Function FolderDialog(strTitle As String, _
                             Optional blnNetGetConnection As Boolean = False, _
                             Optional strRootPath As String = "", _
                             Optional swFixRootPath As Integer = 0, _
                             Optional strDefaultRootPath As String = "", _
                             Optional strButtonName As String = "OK") As String
    '---------------------------------------------------------------------------
    Static strPrevDir As String
    Dim strPathName As String, strPathName2 As String
    ' ファイルダイアログの表示(FolderPicker)
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = strTitle
        ' ルートフォルダの設定
        Call GP_SetRootPathName( _
            .InitialFileName, strPrevDir, strRootPath, strDefaultRootPath, swFixRootPath)
        ' ボタン名の設定
        If Len(strButtonName) > 0 Then
            .ButtonName = strButtonName
        End If
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        If .Show Then
            strPathName = .SelectedItems(1)
            ' ローカルドライブ(ネットワークドライブ)から
            ' マウントされているネットワークリソース名を取得する
            If (blnNetGetConnection And (Mid$(strPathName, 2, 1) = g_cnsCol)) Then
                strPathName2 = GetResourceNameFromLocalDrive(strPathName)
                If Left(strPathName2, 2) = g_cnsYen & g_cnsYen Then
                    strPathName = strPathName2 & Mid$(strPathName, 3)
                End If
            End If
            If ((swFixRootPath = 0) Or (Len(strPrevDir) = 0)) Then
                strPrevDir = strPathName
            End If
        ElseIf swFixRootPath = 3 Then
            ' キャンセル時は初期化の指定
            strPrevDir = strDefaultRootPath
        End If
    End With
    FolderDialog = strPathName
End Function

'*******************************************************************************
'   「ファイルを開く」ダイアログを表示させ、選択したファイル名を返す
'-------------------------------------------------------------------------------
'   引渡値 = ①ウィンドウタイトル
'   　　　   ②ファイルフィルタ(Option:2次配列)
'   　　　   ③ネットワークドライブ→ネットワークリソースの置換区分(Option)
'   　　　   ④ルートフォルダ(Option)
'            ⑤ルートフォルダ固定スイッチ
'              (Option, 1=固定する, 2=ルートに戻す, 3=キャンセル時は初期化)
'   　　　   ⑥初期化用ルートフォルダ(Option)
'            ⑦ボタン表示名(Option)
'   戻り値 = ファイル名(フルパス、未選択時はブランク)
'-------------------------------------------------------------------------------
' 　作成日：2014年04月14日
' 　作成者：井上 治
' 　更新日：2014年04月14日
' 　更新者：井上 治
'*******************************************************************************
Public Function OpenDialog(strTitle As String, _
                           Optional tblFilter As Variant, _
                           Optional blnNetGetConnection As Boolean = False, _
                           Optional strRootPath As String = "", _
                           Optional swFixRootPath As Integer = 0, _
                           Optional strDefaultRootPath As String = "", _
                           Optional strButtonName As String = "開く") As String
    '---------------------------------------------------------------------------
    Static strPrevDir As String
    Dim strFileName As String, strFileName2 As String
    Dim IX As Integer
    ' ファイルダイアログの表示(FileDialogOpen)
    With Application.FileDialog(msoFileDialogOpen)
        .Title = strTitle
        ' ルートフォルダの設定
        Call GP_SetRootPathName( _
            .InitialFileName, strPrevDir, strRootPath, strDefaultRootPath, swFixRootPath)
        ' ボタン名の設定
        If Len(strButtonName) > 0 Then
            .ButtonName = strButtonName
        End If
        ' ファイルフィルタの設定
        If IsArray(tblFilter) Then
            With .Filters
                .Clear
                IX = 0
                Do While IX <= UBound(tblFilter)
                    .Add tblFilter(IX, 0), tblFilter(IX, 1), IX + 1
                    IX = IX + 1
                Loop
            End With
        End If
        .InitialFileName = strPrevDir
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        If .Show Then
            strFileName = .SelectedItems(1)
            ' ローカルドライブ(ネットワークドライブ)から
            ' マウントされているネットワークリソース名を取得する
            If (blnNetGetConnection And (Mid$(strFileName, 2, 1) = g_cnsCol)) Then
                strFileName2 = GetResourceNameFromLocalDrive(strFileName)
                If Left(strFileName2, 2) = g_cnsYen & g_cnsYen Then
                    strFileName = strFileName2 & Mid$(strFileName, 3)
                End If
            End If
            If ((swFixRootPath = 0) Or (Len(strPrevDir) = 0)) Then
                strFileName2 = Left(strFileName, InStrRev(strFileName, g_cnsYen))
                strPrevDir = strFileName2
            End If
        ElseIf swFixRootPath = 3 Then
            ' キャンセル時は初期化の指定
            strPrevDir = strDefaultRootPath
        End If
    End With
    OpenDialog = strFileName
End Function

'*******************************************************************************
'   「名前を付けて保存」ダイアログを表示させ、選択したファイル名を返す
'-------------------------------------------------------------------------------
'   引渡値 = ①ウィンドウタイトル
'   　　　   ②ネットワークドライブ→ネットワークリソースの置換区分(Option)
'   　　　   ③ルートフォルダ(Option)
'            ④ルートフォルダ固定スイッチ
'              (Option, 1=固定する, 2=ルートに戻す, 3=キャンセル時は初期化)
'   　　　   ⑤初期化用ルートフォルダ(Option)
'            ⑥ボタン表示名(Option)
'   戻り値 = ファイル名(フルパス、未選択時はブランク)
'-------------------------------------------------------------------------------
' 　作成日：2014年04月14日
' 　作成者：井上 治
' 　更新日：2014年04月14日
' 　更新者：井上 治
'*******************************************************************************
Public Function SaveDialog(strTitle As String, _
                           Optional blnNetGetConnection As Boolean = False, _
                           Optional strRootPath As String = "", _
                           Optional swFixRootPath As Integer = 0, _
                           Optional strDefaultRootPath As String = "", _
                           Optional strButtonName As String = "保存") As String
    '---------------------------------------------------------------------------
    Static strPrevDir As String
    Dim strFileName As String, strFileName2 As String
    Dim IX As Integer
    ' ファイルダイアログの表示(FileDialogSaveAs)
    ' この方法ではファイルフィルタの指定ができません
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = strTitle
        ' ルートフォルダの設定
        Call GP_SetRootPathName( _
            .InitialFileName, strPrevDir, strRootPath, strDefaultRootPath, swFixRootPath)
        ' ボタン名の設定
        If Len(strButtonName) > 0 Then
            .ButtonName = strButtonName
        End If
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        If .Show Then
            strFileName = .SelectedItems(1)
            ' ローカルドライブ(ネットワークドライブ)から
            ' マウントされているネットワークリソース名を取得する
            If (blnNetGetConnection And (Mid$(strFileName, 2, 1) = g_cnsCol)) Then
                strFileName2 = GetResourceNameFromLocalDrive(strFileName)
                If Left(strFileName2, 2) = g_cnsYen & g_cnsYen Then
                    strFileName = strFileName2 & Mid$(strFileName, 3)
                End If
            End If
            If ((swFixRootPath = 0) Or (Len(strPrevDir) = 0)) Then
                strFileName2 = Left(strFileName, InStrRev(strFileName, g_cnsYen))
                strPrevDir = strFileName2
            End If
        ElseIf swFixRootPath = 3 Then
            ' キャンセル時は初期化の指定
            strPrevDir = strDefaultRootPath
        End If
    End With
    SaveDialog = strFileName
End Function

'*******************************************************************************
'   ネットワークドライブシンボルからネットワークリソースを取得
'-------------------------------------------------------------------------------
'   引渡値 = ①パス名
'   戻り値 = パス名
'-------------------------------------------------------------------------------
' 　作成日：2014年04月14日
' 　作成者：井上 治
' 　更新日：2014年04月14日
' 　更新者：井上 治
'*******************************************************************************
Public Function GetResourceNameFromLocalDrive(strDrv As String) As String
    Dim strBuf As String
    Dim strDriveName As String
    Dim lngLen As Long
    '---------------------------------------------------------------------------
    strDriveName = Left$(strDrv, 1) & g_cnsCol
    On Error GoTo GetResourceNameFromLocalDrive_ERROR
    strBuf = String$(MAX_PATH + 1, vbNullChar)
    WNetGetConnection strDriveName, strBuf, MAX_PATH
    '取得したパス名から必要な文字列だけを抽出
    lngLen = InStr(1, strBuf, vbNullChar)
    If lngLen > 1 Then
        GetResourceNameFromLocalDrive = Left$(strBuf, lngLen - 1)
    Else
        GetResourceNameFromLocalDrive = strDriveName
    End If
    On Error GoTo 0
    Exit Function
    
'-------------------------------------------------------------------------------
GetResourceNameFromLocalDrive_ERROR:
    GetResourceNameFromLocalDrive = strDriveName
End Function

'*******************************************************************************
' ■■■ 内部共通プロシージャ ■■■
'*******************************************************************************
'   ルートフォルダの設定(Private)
'-------------------------------------------------------------------------------
'   引渡値 = ①Application.FileDialogのInitialFileName
'   　　　   ②直前使用フォルダ
'   　　　   ③ルートフォルダ
'   　　　   ④初期化用ルートフォルダ
'            ⑤ルートフォルダ固定スイッチ
'              (0=通常, 1=固定する, 2=ルートに戻す, 3=キャンセル時は初期化)
'-------------------------------------------------------------------------------
' 　作成日：2014年04月14日
' 　作成者：井上 治
' 　更新日：2014年04月14日
' 　更新者：井上 治
'*******************************************************************************
Private Sub GP_SetRootPathName(ByRef strInitialFileName As String, _
                               ByRef strPrevDir As String, _
                               ByRef strRootPath As String, _
                               ByRef strDefaultRootPath As String, _
                               ByVal swFixRootPath As Integer)
    '---------------------------------------------------------------------------
    If Len(strRootPath) = 0 Then
        strRootPath = strDefaultRootPath
    End If
    If Len(strRootPath) > 0 Then
        If swFixRootPath = 1 Then
            ' ルートフォルダ固定の指定
            strPrevDir = strRootPath
        End If
        If ((Len(strPrevDir) <= 0) Or (swFixRootPath = 2)) Then
            strPrevDir = strRootPath
        End If
    End If
    If Len(strPrevDir) > 0 Then
        strInitialFileName = strPrevDir
    End If
End Sub

'----------------------------<< End of Source >>--------------------------------

