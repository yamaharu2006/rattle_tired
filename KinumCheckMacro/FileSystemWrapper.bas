Attribute VB_Name = "FileSystemWrapper"
' @breif �Ζ��\�t�@�C������𕽈ՂɈ������߂ɗp�ӂ������C�u����
' @note �t�@�C������S�ʂ��������C�u��������Log�ɂ��Ă��l����K�v��������̂ŋΖ��\�t�@�C���Ɍ��肵����
Option Explicit

Private Const filePassword = "pass"

' @breif �t�@�C�����J��
Public Function OpenWorkbook(folderPath As String, fileName As String) As Boolean
    LogApiIn "OpenWorkbook()"
    
    Dim fullName As String
    fullName = GenerateFullName(folderPath, fileName)
    
    If ExistsFile(fullName) = False Then
        OpenWorkbook = False
        Exit Function
    ElseIf IsOpenedSameFile(fullName) = False Then
        Workbooks.Open fileName:=fullName, ReadOnly:=True, Password:=filePassword
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
Public Function GenerateFullName(folderPath As String, fileName As String) As String
    LogApiIn "OpenWorkbook()"
    
    GenerateFullName = folderPath & "\" & fileName
    
    LogApiIn "OpenWorkbook()"
End Function

' @breif �t�@�C�������݂��Ă��邩��Ԃ�
Public Function ExistsFile(fullName As String) As Boolean
    If Dir(fullName) = "" Then
        ExistsFile = False
    Else
        ExistsFile = True
    End If
End Function

' @breif �����̋Ζ��\�t�@�C�����J���Ă��邩��Ԃ�
' @attention �u�b�N���J���Ă���ΊJ���Ă���قǃ`�F�b�N�Ɏ��Ԃ�������
Public Function IsOpenedSameFile(fileName As String) As Boolean
    LogApiIn "IsOpenedSameFile()"

    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.name = fileName Then
            IsOpenedSameFile = True
            Exit Function
        End If
    Next wb
    IsOpenedSameFile = False

    LogApiOut "IsOpenedSameFile()"
End Function

' @breif �t�@�C�����N���[�Y����
' @note �t�@�C������n�֐��Q�͕ʃN���X�ɈϏ�������
Public Function CloseWorkbook(fileName As String)
    LogApiIn "CloseWorkbook()"
    
    If IsOpenedSameFile(fileName) = True Then
        Workbooks(fileName).Close
    End If
    
    LogApiIn "CloseWorkbook()"
End Function

' @breif  �t�@�C�����o�b�N�A�b�v����
' @attention �ċA�I�Ƀt�H���_���쐬����s�ׂ͊댯�������邽�߁A�e�t�H���_���Ȃ��ꍇ�̓|�b�v�A�b�v��\��������
Public Function BackupFile(Path As String, fileName As String, bkupPath As String, bkupFileName As String) As Boolean
    LogApiIn "SaveBackupFile()"
    
    Dim parentDir As String
    parentDir = GetParentDir(bkupPath)
    If Dir(parentDir, vbDirectory) = "" Then
        Dim pressed
        pressed = MsgBox("�w�E���ꂽ�o�b�N�A�b�v��̐e�t�H���_�[������܂���B" & vbCrLf & "�t�H���_�[���ċA�I�ɍ쐬���܂����H" & vbCrLf & parentDir, vbOKCancel)
        If pressed = vbCancel Then
            BackupFile = False
            Exit Function
        End If
    End If
    
    MkDirRecursive bkupPath
    
    CopyBackupFile Path, fileName, bkupPath, bkupFileName
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
Private Function CopyBackupFile(Path As String, fileName As String, bkupPath As String, bkupFileName As String)
    LogApiIn "CopyBackupFile()"
    
    Dim fullName As String
    fullName = GenerateFullName(Path, fileName)
    
    If Dir(fullName) <> "" Then
        Dim bkupFullName As String
        bkupFullName = GenerateFullName(bkupPath, bkupFileName)
        FileCopy GenerateFullName(Path, fileName), bkupFullName
    End If

    LogApiOut "CopyBackupFile()"
End Function


' @breif �t�@�C���̍ŏI�X�V�������擾����
Public Function GetDateLastModified(FilePath As String, ByRef lastModified As Date) As Boolean
    LogApiIn "GetDateLastModified()"
    
    If Dir(FilePath) = "" Then
        GetDateLastModified = False
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

