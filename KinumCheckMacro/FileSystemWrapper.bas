Attribute VB_Name = "FileSystemWrapper"
' @breif �Ζ��\�t�@�C������𕽈ՂɈ������߂ɗp�ӂ������C�u����
' @note �t�@�C������S�ʂ��������C�u��������Log�ɂ��Ă��l����K�v��������̂ŋΖ��\�t�@�C���Ɍ��肵����
Option Explicit

Private Const filePassword = "pass"

' @breif �t�@�C�����J��
' @note �����ŃG���[���E���̂����Ƃ����Ⴎ����ɂȂ��Ă��Ă���
Public Function OpenWorkbook(folderPath As String, FileName As String) As Boolean
    LogApiIn "OpenWorkbook()"
    
    Dim FullName As String
    FullName = GenerateFullName(folderPath, FileName)
    
    If ExistsFile(FullName) = False Then
        OpenWorkbook = False
        Exit Function
    ElseIf IsOpenedSameFile(FullName) = False Then
        Workbooks.Open FileName:=FullName, ReadOnly:=True, Password:=filePassword
    End If
    
    OpenWorkbook = True
    
    LogApiOut "OpenWorkbook()"
End Function

' @breif �t���p�X����t�@�C�������擾����
' @note working...
Public Function ConvertPathToFileName(Path As String) As String
    LogApiIn "OpenWorkbook()"
    
    ConvertPathToFileName = ""
    
    LogApiIn "OpenWorkbook()"
End Function

' @breif ��΃p�X�𐶐�����
Public Function GenerateFullName(folderPath As String, FileName As String) As String
    LogApiIn "OpenWorkbook()"
    
    GenerateFullName = folderPath & "\" & FileName
    
    LogApiIn "OpenWorkbook()"
End Function

' @breif �t�@�C�������݂��Ă��邩��Ԃ�
Public Function ExistsFile(FullName As String) As Boolean
    If Dir(FullName) = "" Then
        ExistsFile = False
    Else
        ExistsFile = True
    End If
End Function

' @breif �����̋Ζ��\�t�@�C�����J���Ă��邩��Ԃ�
' @attention �u�b�N���J���Ă���ΊJ���Ă���قǃ`�F�b�N�Ɏ��Ԃ�������
Public Function IsOpenedSameFile(FileName As String) As Boolean
    LogApiIn "IsOpenedSameFile()"

    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name = FileName Then
            IsOpenedSameFile = True
            Exit Function
        End If
    Next wb
    IsOpenedSameFile = False

    LogApiOut "IsOpenedSameFile()"
End Function

' @breif �t�@�C�����N���[�Y����
' @note �t�@�C������n�֐��Q�͕ʃN���X�ɈϏ�������
Public Function CloseWorkbook(FileName As String)
    LogApiIn "CloseWorkbook()"
    
    If IsOpenedSameFile(FileName) = True Then
        Workbooks(FileName).Close
    End If
    
    LogApiOut "CloseWorkbook()"
End Function

' @breif  �t�@�C�����o�b�N�A�b�v����
' @attention �ċA�I�Ƀt�H���_���쐬����s�ׂ͊댯�𔺂����߁A�e�t�H���_���Ȃ��ꍇ�̓|�b�v�A�b�v��\��������
Public Function BackupFile(Path As String, FileName As String, bkupPath As String, bkupFileName As String) As Boolean
    LogApiIn "SaveBackupFile()"
    
    Dim parentDir As String
    parentDir = GetParentDir(bkupPath)
    If Dir(parentDir, vbDirectory) = "" Then
        Dim pressed
        pressed = MsgBox("�w�E���ꂽ�o�b�N�A�b�v��̐e�t�H���_�[������܂���B" & vbCrLf & "�t�H���_�[���ċA�I�ɍ쐬���܂����H" & vbCrLf & parentDir, vbOKCancel)
        If pressed = vbCancel Then
            BackupFile = False
            LogApiOut "SaveBackupFile()"
            Exit Function
        End If
    End If
    
    MkDirRecursive bkupPath
    
    CopyBackupFile Path, FileName, bkupPath, bkupFileName
    BackupFile = True
    
    LogApiOut "SaveBackupFile()"
End Function

' @breif �e�̃f�B���N�g���p�X���擾����
' @note https://www.atmarkit.co.jp/ait/articles/1705/01/news019.html
Private Function GetParentDir(Path As String)
    LogApiIn "GetParentDir()"

    Dim fso As New Scripting.FileSystemObject
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(Path)
    Set fso = Nothing
    GetParentDir = parentPath
    
    LogApiOut "GetParentDir()"
End Function

' @breif �K�w�I�ȃf�B���N�g�����܂Ƃ߂č쐬����
' @note ���̊֐��͎�舵�����댯�Ȃ̂Ń|�b�v�A�b�v���o�����ق�����������
' https://www.relief.jp/docs/excel-vba-mkdir-folder-structure.html
Private Function MkDirRecursive(Path As String)
  LogApiIn "MkDirRecursive()"

  Dim arr() As String
  arr = Split(Path, "\")

  Dim i As Long
  For i = 1 To UBound(arr)
    Dim tmpPath As String
    tmpPath = tmpPath & "\" & arr(i)
    If Dir(tmpPath, vbDirectory) = "" Then
      MkDir tmpPath
    End If
  Next i

    LogApiOut "MkDirRecursive()"
End Function

' @breif �t�@�C�����R�s�[����
Private Function CopyBackupFile(Path As String, FileName As String, bkupPath As String, bkupFileName As String)
    LogApiIn "CopyBackupFile()"
    
    Dim FullName As String
    FullName = GenerateFullName(Path, FileName)
    
    If Dir(FullName) <> "" Then
        Dim bkupFullName As String
        bkupFullName = GenerateFullName(bkupPath, bkupFileName)
        FileCopy GenerateFullName(Path, FileName), bkupFullName
    End If

    LogApiOut "CopyBackupFile()"
End Function


' @breif �t�@�C���̍ŏI�X�V�������擾����
Public Function GetDateLastModified(FilePath As String, ByRef lastModified As Date) As Boolean
    LogApiIn "GetDateLastModified()"
    
    If Dir(FilePath) = "" Then
        GetDateLastModified = False
        LogApiOut "GetDateLastModified()"
        Exit Function
    End If
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject ' �C���X�^���X��
    
    Dim f As File
    Set f = fso.GetFile(FilePath) ' �t�@�C�����擾
    
    lastModified = f.dateLastModified ' �X�V�������擾
    
    GetDateLastModified = True
    LogApiOut "GetDateLastModified()"
End Function

