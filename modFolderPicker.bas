Attribute VB_Name = "modFolderPicker"
'*******************************************************************************
'   �t�@�C���A�t�H���_�Q�Ƃ̃_�C�A���O����(Application.FileDialog)
'
'   �쐬��:��㎡  URL:http://www.ne.jp/asahi/excel/inoue/ [Excel�ł��d��!]
'*******************************************************************************
'�ύX���t Rev  �ύX������e---------------------------------------------------->
'14/04/14(1.00)�V�K�쐬
'*******************************************************************************
Option Explicit
Private Const MAX_PATH As Long = 260
Private Const g_cnsYen As String = "\"
Private Const g_cnsCol As String = ":"
' �����[�J���h���C�u����}�E���g����Ă���l�b�g���[�N���\�[�X�����擾����
Private Declare Function WNetGetConnection Lib "MPR.dll" _
    Alias "WNetGetConnectionA" _
    (ByVal lpszLocalName As String, _
     ByVal lpszRemoteName As String, _
     cbRemoteName As Long) As Long

'*******************************************************************************
'   �u�t�H���_�̎Q�Ɓv�_�C�A���O��\�������A�I�������t�H���_����Ԃ�
'-------------------------------------------------------------------------------
'   ���n�l = �@�E�B���h�E�^�C�g��
'   �@�@�@   �A�l�b�g���[�N�h���C�u���l�b�g���[�N���\�[�X�̒u���敪(Option)
'   �@�@�@   �B���[�g�t�H���_(Option)
'            �C���[�g�t�H���_�Œ�X�C�b�`
'              (Option, 1=�Œ肷��, 2=���[�g�ɖ߂�, 3=�L�����Z�����͏�����)
'   �@�@�@   �D�������p���[�g�t�H���_(Option)
'            �E�{�^���\����(Option)
'   �߂�l = �t�H���_��(�t���p�X�ŉE\�Ȃ��A���I�����̓u�����N)
'-------------------------------------------------------------------------------
' �@�쐬���F2014�N04��14��
' �@�쐬�ҁF��� ��
' �@�X�V���F2014�N04��14��
' �@�X�V�ҁF��� ��
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
    ' �t�@�C���_�C�A���O�̕\��(FolderPicker)
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = strTitle
        ' ���[�g�t�H���_�̐ݒ�
        Call GP_SetRootPathName( _
            .InitialFileName, strPrevDir, strRootPath, strDefaultRootPath, swFixRootPath)
        ' �{�^�����̐ݒ�
        If Len(strButtonName) > 0 Then
            .ButtonName = strButtonName
        End If
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        If .Show Then
            strPathName = .SelectedItems(1)
            ' ���[�J���h���C�u(�l�b�g���[�N�h���C�u)����
            ' �}�E���g����Ă���l�b�g���[�N���\�[�X�����擾����
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
            ' �L�����Z�����͏������̎w��
            strPrevDir = strDefaultRootPath
        End If
    End With
    FolderDialog = strPathName
End Function

'*******************************************************************************
'   �u�t�@�C�����J���v�_�C�A���O��\�������A�I�������t�@�C������Ԃ�
'-------------------------------------------------------------------------------
'   ���n�l = �@�E�B���h�E�^�C�g��
'   �@�@�@   �A�t�@�C���t�B���^(Option:2���z��)
'   �@�@�@   �B�l�b�g���[�N�h���C�u���l�b�g���[�N���\�[�X�̒u���敪(Option)
'   �@�@�@   �C���[�g�t�H���_(Option)
'            �D���[�g�t�H���_�Œ�X�C�b�`
'              (Option, 1=�Œ肷��, 2=���[�g�ɖ߂�, 3=�L�����Z�����͏�����)
'   �@�@�@   �E�������p���[�g�t�H���_(Option)
'            �F�{�^���\����(Option)
'   �߂�l = �t�@�C����(�t���p�X�A���I�����̓u�����N)
'-------------------------------------------------------------------------------
' �@�쐬���F2014�N04��14��
' �@�쐬�ҁF��� ��
' �@�X�V���F2014�N04��14��
' �@�X�V�ҁF��� ��
'*******************************************************************************
Public Function OpenDialog(strTitle As String, _
                           Optional tblFilter As Variant, _
                           Optional blnNetGetConnection As Boolean = False, _
                           Optional strRootPath As String = "", _
                           Optional swFixRootPath As Integer = 0, _
                           Optional strDefaultRootPath As String = "", _
                           Optional strButtonName As String = "�J��") As String
    '---------------------------------------------------------------------------
    Static strPrevDir As String
    Dim strFileName As String, strFileName2 As String
    Dim IX As Integer
    ' �t�@�C���_�C�A���O�̕\��(FileDialogOpen)
    With Application.FileDialog(msoFileDialogOpen)
        .Title = strTitle
        ' ���[�g�t�H���_�̐ݒ�
        Call GP_SetRootPathName( _
            .InitialFileName, strPrevDir, strRootPath, strDefaultRootPath, swFixRootPath)
        ' �{�^�����̐ݒ�
        If Len(strButtonName) > 0 Then
            .ButtonName = strButtonName
        End If
        ' �t�@�C���t�B���^�̐ݒ�
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
            ' ���[�J���h���C�u(�l�b�g���[�N�h���C�u)����
            ' �}�E���g����Ă���l�b�g���[�N���\�[�X�����擾����
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
            ' �L�����Z�����͏������̎w��
            strPrevDir = strDefaultRootPath
        End If
    End With
    OpenDialog = strFileName
End Function

'*******************************************************************************
'   �u���O��t���ĕۑ��v�_�C�A���O��\�������A�I�������t�@�C������Ԃ�
'-------------------------------------------------------------------------------
'   ���n�l = �@�E�B���h�E�^�C�g��
'   �@�@�@   �A�l�b�g���[�N�h���C�u���l�b�g���[�N���\�[�X�̒u���敪(Option)
'   �@�@�@   �B���[�g�t�H���_(Option)
'            �C���[�g�t�H���_�Œ�X�C�b�`
'              (Option, 1=�Œ肷��, 2=���[�g�ɖ߂�, 3=�L�����Z�����͏�����)
'   �@�@�@   �D�������p���[�g�t�H���_(Option)
'            �E�{�^���\����(Option)
'   �߂�l = �t�@�C����(�t���p�X�A���I�����̓u�����N)
'-------------------------------------------------------------------------------
' �@�쐬���F2014�N04��14��
' �@�쐬�ҁF��� ��
' �@�X�V���F2014�N04��14��
' �@�X�V�ҁF��� ��
'*******************************************************************************
Public Function SaveDialog(strTitle As String, _
                           Optional blnNetGetConnection As Boolean = False, _
                           Optional strRootPath As String = "", _
                           Optional swFixRootPath As Integer = 0, _
                           Optional strDefaultRootPath As String = "", _
                           Optional strButtonName As String = "�ۑ�") As String
    '---------------------------------------------------------------------------
    Static strPrevDir As String
    Dim strFileName As String, strFileName2 As String
    Dim IX As Integer
    ' �t�@�C���_�C�A���O�̕\��(FileDialogSaveAs)
    ' ���̕��@�ł̓t�@�C���t�B���^�̎w�肪�ł��܂���
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = strTitle
        ' ���[�g�t�H���_�̐ݒ�
        Call GP_SetRootPathName( _
            .InitialFileName, strPrevDir, strRootPath, strDefaultRootPath, swFixRootPath)
        ' �{�^�����̐ݒ�
        If Len(strButtonName) > 0 Then
            .ButtonName = strButtonName
        End If
        .InitialView = msoFileDialogViewDetails
        .AllowMultiSelect = False
        If .Show Then
            strFileName = .SelectedItems(1)
            ' ���[�J���h���C�u(�l�b�g���[�N�h���C�u)����
            ' �}�E���g����Ă���l�b�g���[�N���\�[�X�����擾����
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
            ' �L�����Z�����͏������̎w��
            strPrevDir = strDefaultRootPath
        End If
    End With
    SaveDialog = strFileName
End Function

'*******************************************************************************
'   �l�b�g���[�N�h���C�u�V���{������l�b�g���[�N���\�[�X���擾
'-------------------------------------------------------------------------------
'   ���n�l = �@�p�X��
'   �߂�l = �p�X��
'-------------------------------------------------------------------------------
' �@�쐬���F2014�N04��14��
' �@�쐬�ҁF��� ��
' �@�X�V���F2014�N04��14��
' �@�X�V�ҁF��� ��
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
    '�擾�����p�X������K�v�ȕ����񂾂��𒊏o
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
' ������ �������ʃv���V�[�W�� ������
'*******************************************************************************
'   ���[�g�t�H���_�̐ݒ�(Private)
'-------------------------------------------------------------------------------
'   ���n�l = �@Application.FileDialog��InitialFileName
'   �@�@�@   �A���O�g�p�t�H���_
'   �@�@�@   �B���[�g�t�H���_
'   �@�@�@   �C�������p���[�g�t�H���_
'            �D���[�g�t�H���_�Œ�X�C�b�`
'              (0=�ʏ�, 1=�Œ肷��, 2=���[�g�ɖ߂�, 3=�L�����Z�����͏�����)
'-------------------------------------------------------------------------------
' �@�쐬���F2014�N04��14��
' �@�쐬�ҁF��� ��
' �@�X�V���F2014�N04��14��
' �@�X�V�ҁF��� ��
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
            ' ���[�g�t�H���_�Œ�̎w��
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

