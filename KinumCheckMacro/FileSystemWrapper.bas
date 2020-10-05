Attribute VB_Name = "FileSystemWrapper"
' @breif �t�@�C������𕽈ՂɈ������߂ɗp�ӂ������C�u����
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
Public Function ConvertPathToFileName(path As String) As String
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

' �t�@�C�����o�b�N�A�b�v����
' @note �o�O�F�o�b�N�A�b�v�̃p�X��2�K�w�ȏ�t�H���_�����K�v���������Ƃ������Ȃ�
Public Function BackupFile(path As String, fileName As String, bkupPath As String, bkupFileName As String)
    LogApiIn "SaveBackupFile()"
    
    MkDirRecursive bkupPath
    CopyBackupFile path, fileName, bkupPath, bkupFileName
    
    LogApiOut "SaveBackupFile()"
End Function


' @breif �K�w�I�ȃf�B���N�g�����܂Ƃ߂č쐬����
' @note ���̊֐��͎�舵�����댯�Ȃ̂Ń|�b�v�A�b�v���o�����ق�����������
' https://www.relief.jp/docs/excel-vba-mkdir-folder-structure.html
Private Function MkDirRecursive(path As String)
  LogApiIn "MkDirRecursive()"

  Dim arr() As String
  arr = Split(path, "\")

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
Private Function CopyBackupFile(path As String, fileName As String, bkupPath As String, bkupFileName As String)
    LogApiIn "CopyBackupFile()"
    
    Dim fullName As String
    fullName = GenerateFullName(path, fileName)
    
    If Dir(fullName) <> "" Then
        Dim bkupFullName As String
        bkupFullName = GenerateFullName(bkupPath, bkupFileName)
        FileCopy GenerateFullName(path, fileName), bkupFullName
    End If

    LogApiOut "CopyBackupFile()"
End Function


' @breif �t�@�C���̍ŏI�X�V�������擾����
Public Function GetDateLastModified(FilePath As String) As Date
    LogApiIn "GetDateLastModified()"
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject ' �C���X�^���X��
    
    Dim f As File
    Set f = fso.GetFile(FilePath) ' �t�@�C�����擾
    
    Dim lastModified As Date
    lastModified = f.dateLastModified ' �X�V�������擾
    
    GetDateLastModified = lastModified
    
    LogApiOut "GetDateLastModified()"
End Function
