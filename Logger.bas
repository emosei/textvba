Attribute VB_Name = "Logger"
Public Const LOG_SW = True      '���O�o�̗͂L���E����
Public Const LOG_SHEET = "Log"  '���O�̏o�͐�V�[�g��
Public Const MAX_LOG = 2         '���O�̖{���i���{�ł��j
Public Const LOG_START_ROW = 2     '���O�o�͊J�n�s
Public Const INIT_LOG = "INIT"    '�������w��
Public Const ERROR_LOG = 1         'ERROR �p���O
Public Const OUTPUT_LOG = 2         '�f�[�^�f���o�����O

'�ȉ��́A���O�f���o���܂���}�N���{�́B�i�ǂ̃��W���[���ɏ����Ă��B�j
Public Sub debugLog(log As String, opt As Integer)
  Static last_line(MAX_LOG) As Integer    '���Ƀ��O�������s
  If Not LOG_SW Then Exit Sub
  If log = INIT_LOG Then
    For n = 1 To MAX_LOG
      last_line(n) = LOG_START_ROW
      '���O�̈�̃N���A
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

