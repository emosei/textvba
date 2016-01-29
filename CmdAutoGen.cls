VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmdAutoGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' コード定義書の内容を元に、設計書の特定のセルに値を設定する
' DB定義を参照して
'*******************************************************************************

Option Explicit

'*******************************************************************************
' 定数定義
'*******************************************************************************

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

Private Const COL_NEW_TYPEDEF_TABLE_NAME = 21 ' テーブル名
Private Const COL_NEW_TYPEDEF_COLUMN_NAME = 22 ' 項目名


'インタフェース仕様書_型定義シートのシート名接頭語
Private Const TARGET_SHEET_PREFIX = "型定義_"

' コード定義辞書
Private m_codeDefDic As CodeDictionary
' DB項目定義辞書
Private m_dbDefDic As DBDictionary
'*******************************************************************************
' コンストラクタ
'*******************************************************************************
Private Sub Class_Initialize()
    Set m_codeDefDic = New CodeDictionary
    Set m_dbDefDic = New DBDictionary
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
    
    ' 辞書を取得
    Dim defCont As String, codeCont As String, varchar2Size As Integer
    Dim colType As String, pkey As Boolean, notnull As Boolean, explain As String, sample As String
    'モードを取得
    Dim modeSheet As Worksheet, mode As String
    Set modeSheet = ThisWorkbook.Worksheets("DB項目チェック")
    
    mode = modeSheet.Range("C14").Value
    
    If InStr(sheetName, TARGET_SHEET_PREFIX) = 1 Then
        For i = 2 To targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).row

            ' No列の定義がなくなるまでループ
            If targetSheet.Cells(i, COL_NEW_TYPEDEF_NO) = "" Then
                Exit For
            End If
            
            '項目IDを空白にする。
            targetSheet.Cells(i, COL_NEW_TYPEDEF_ID).Value = ""
            
            '要素の場合に
            If targetSheet.Cells(i, COL_NEW_TYPEDEF_TYPE) = "要素" Then
                GoTo CONTINUE
            End If
            
            '導出元の記述がある場合
            defCont = ""
            codeCont = ""
            If targetSheet.Cells(i, COL_NEW_TYPEDEF_TABLE_NAME) <> "" And targetSheet.Cells(i, COL_NEW_TYPEDEF_COLUMN_NAME) <> "" Then
                defCont = m_dbDefDic.GetValue(targetSheet.Cells(i, COL_NEW_TYPEDEF_TABLE_NAME), targetSheet.Cells(i, COL_NEW_TYPEDEF_COLUMN_NAME))
                codeCont = m_codeDefDic.GetValue(targetSheet.Cells(i, COL_NEW_TYPEDEF_TABLE_NAME), targetSheet.Cells(i, COL_NEW_TYPEDEF_COLUMN_NAME))
            Else
                defCont = m_dbDefDic.GetValue("", targetSheet.Cells(i, COL_NEW_TYPEDEF_NAME))
            End If
            
            
            ' エラーじゃない場合
            If defCont <> "" And InStr(defCont, "ERR:") <> 1 Then
            
                Dim tmp As Variant
                tmp = Split(defCont, "@")
                colType = tmp(0)
                If tmp(1) = "○" Then
                    pkey = True
                Else
                    pkey = False
                End If
                
                If tmp(3) = "○" Then
                    notnull = True
                Else
                    notnull = False
                End If
                explain = tmp(4)
                sample = tmp(5)
                
                If notnull Then
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MIN_LOOP, 1)
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAX_LOOP, 1)
                Else
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MIN_LOOP, 0)
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAX_LOOP, 1)
                End If

                Dim dbType As String  'DBのデータ型
                Dim dbLength As String 'DB項目の長さ
                
                Dim sizeOpenPos As Integer: sizeOpenPos = InStr(colType, "(")
                '全角に対応
                If sizeOpenPos = 0 Then
                    sizeOpenPos = InStr(colType, "（")
                End If
                
                Dim sizeClosePos As Integer: sizeClosePos = InStr(colType, ")")
                '全角に対応
                If sizeClosePos = 0 Then
                    sizeClosePos = InStr(colType, "）")
                End If
                '-------
                ' 補足
                ' 「CHAR(8)」というような表記から、"CHAR" という型情報と 8という長さを求める処理
                '-------
                If sizeOpenPos > 0 Then
                    dbType = StrConv(Mid(colType, 1, sizeOpenPos - 1), vbUpperCase)
                    dbLength = Mid(colType, sizeOpenPos + 1, sizeClosePos - sizeOpenPos - 1)
                Else
                    dbType = colType
                    dbLength = 0
                End If

                '===================
                ' サンプル値の修正
                '===================
                sample = RebuildSampleText(sample, dbType, dbLength, codeCont)

                If dbType = "CHAR" Then
                    varchar2Size = dbLength
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_TYPE, "文字列")
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MINLENGTH, varchar2Size)
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, varchar2Size)
                    
                    'サンプルの修正
                    sample = Padding(sample, "X", varchar2Size)
                
                ElseIf dbType = "VARCHAR2" Or dbType = "VARCHAR" Or dbType = "VACHAR2" Then '変な間違いに対応（課題96）
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_TYPE, "文字列")
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MINLENGTH, "1")
                    varchar2Size = dbLength
                    If (varchar2Size Mod 2) > 0 Then
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, (varchar2Size))
                        
                    ElseIf HankakuColumn(targetSheet.Cells(i, COL_NEW_TYPEDEF_COLUMN_NAME)) Then
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, (varchar2Size))
                        
                    Else
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, (varchar2Size / 2))
                    End If
                    
                ElseIf dbType = "NUMBER" Then
                    If InStr(dbLength, ",") > 0 Then
                        varchar2Size = Mid(dbLength, 1, InStr(dbLength, ",") - 1)
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_TYPE, "数値")
                    Else
                        varchar2Size = dbLength
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_TYPE, "整数")
                    End If
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MINLENGTH, 1)
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, varchar2Size)
 
                ElseIf dbType = "DATE" Then
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_TYPE, "日時")
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MINLENGTH, 1)
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_MAXLENGTH, 12)

                Else
                    Call debugLog("シート名「" & targetSheet.Name & "」/ 行「" & i & "」/ 予期せぬ型です:" & targetSheet.Name & ":" & colType & ":" & dbType, ERROR_LOG)
                End If
                
                '説明
                Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_EXPLAIN, explain)
                
                'データ例
                Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_DATA_SAMPLE, sample)
                
                ' 文字種・書式
                If codeCont <> "" Then
                    Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, codeCont & "のいずれか")
                Else
                    If dbType = "CHAR" Then
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "半角英数")
                        
                        If dbLength = 8 And targetSheet.Cells(i, COL_NEW_TYPEDEF_NAME) Like "*日" Then
                            sample = "20160505"
                        End If
                        
                        If dbLength = 8 And Match("^(\d{4})(\d{2})(\d{2})$", sample) Then
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "yyyyMMdd")
                            'データ例
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_DATA_SAMPLE, "'20160505")
                        End If
                    
                    ElseIf dbType = "VARCHAR2" Then
                        varchar2Size = dbLength
                        If (varchar2Size Mod 2) > 0 Then
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "半角英数")
                            
                        ElseIf HankakuColumn(targetSheet.Cells(i, COL_NEW_TYPEDEF_COLUMN_NAME)) Then
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "半角英数")
                            
                        ElseIf Match("^(\d{4})(\d{2})(\d{2})$", sample) Then
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "yyyyMMdd")
                            'データ例
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_DATA_SAMPLE, "'20160505")

                        Else
                            Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "")
                        End If
                        
                    ElseIf dbType = "NUMBER" Then
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "半角数字")
                    
                    ElseIf dbType = "DATE" Then
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_CHAR_FORMAT, "")
                        'データ例
                        Call SetCellValue(mode, targetSheet, i, COL_NEW_TYPEDEF_DATA_SAMPLE, "2002-05-30T09:30:10+06:00")

                    End If
                End If
                
            End If
CONTINUE:
        Next
    End If
    
End Sub

Private Function RebuildSampleText(origSample As String, dbType As String, dbLength As String, codeDef As String)

End Function

Private Function Padding(target As String, c As String, length As Integer) As String
    Dim base As String, i As Integer
    Dim sjisStr As String
    sjisStr = StrConv(target, vbFromUnicode)
    
    If LenB(sjisStr) < length Then
        For i = 1 To length
            base = base & c
        Next
        
        Padding = target & Mid(base, 1, length - LenB(sjisStr))
    Else
        Padding = target
    End If
    
End Function



'*******************************************************************************
' セルに対して設定する。
'*******************************************************************************
Private Sub SetCellValue(mode As String, shet As Worksheet, rowIdx As Integer, colIdx As Integer, celVal As Variant)
    If mode = "1" Then
        shet.Cells(rowIdx, colIdx) = celVal
    
    ElseIf mode = "2" Then
        If shet.Cells(rowIdx, colIdx) = "" Then
            shet.Cells(rowIdx, colIdx) = celVal
        End If

    ElseIf mode = "3" Then
        If shet.Cells(rowIdx, colIdx) <> celVal Then
            Call OutputCheckLog(shet, rowIdx, colIdx, celVal)
        End If
    End If
    
End Sub

Private Function HankakuColumn(columnName As String) As Boolean
    If columnName Like "*番号" Or columnName Like "*区分" Or columnName Like "*ID" Or columnName Like "*ＩＤ" Or columnName Like "*英文字" Then
        HankakuColumn = True
    End If
    HankakuColumn = False
End Function

Private Function Match(strPattern As String, target As String) As Boolean
    Dim RegexObj, i As Long, msg As String
    Set RegexObj = CreateObject("VBScript.RegExp")

    With RegexObj
        .pattern = strPattern
        .IgnoreCase = False
        .Global = True
        Match = .Test(target)
    End With
    Set RegexObj = Nothing
End Function

'*******************************************************************************
' チェックログ出力
'*******************************************************************************
Public Sub OutputCheckLog(shet As Worksheet, rowIdx As Integer, colIdx As Integer, newVal As Variant)
    Dim oldVal As Variant
    Set oldVal = shet.Cells(rowIdx, colIdx)
    Dim colName As String
    colName = shet.Cells(rowIdx, COL_NEW_TYPEDEF_NAME)
    Call debugLog("シート名「" & shet.Name & "」/ 項目「" & colName & "」/ セル「row=" & rowIdx & "," & "col=" & colIdx & "」/ 変更前「" & oldVal & "」/変更後「" & newVal & "」", OUTPUT_LOG)
End Sub
