Attribute VB_Name = "FileSystemWrapper"
' @breif 勤務表ファイル操作を平易に扱うために用意したライブラリ
' @note ファイル操作全般を扱うライブラリだとLogについても考える必要が生じるので勤務表ファイルに限定したい
Option Explicit

Private Const filePassword = "pass"

' @breif ファイルを開く
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

' @breif フルパスからファイル名を取得する
' @note working...
Public Function ConvertPathToFileName(Path As String) As String
    LogApiIn "OpenWorkbook()"
    
    ConvertPathToFileName = ""
    
    LogApiIn "OpenWorkbook()"
End Function

' @breif 絶対パスを生成する
Public Function GenerateFullName(folderPath As String, fileName As String) As String
    LogApiIn "OpenWorkbook()"
    
    GenerateFullName = folderPath & "\" & fileName
    
    LogApiIn "OpenWorkbook()"
End Function

' @breif ファイルが存在しているかを返す
Public Function ExistsFile(fullName As String) As Boolean
    If Dir(fullName) = "" Then
        ExistsFile = False
    Else
        ExistsFile = True
    End If
End Function

' @breif 同名の勤務表ファイルを開いているかを返す
' @attention ブックを開いていれば開いているほどチェックに時間がかかる
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

' @breif ファイルをクローズする
' @note ファイル操作系関数群は別クラスに委譲したい
Public Function CloseWorkbook(fileName As String)
    LogApiIn "CloseWorkbook()"
    
    If IsOpenedSameFile(fileName) = True Then
        Workbooks(fileName).Close
    End If
    
    LogApiIn "CloseWorkbook()"
End Function

' @breif  ファイルをバックアップする
' @attention 再帰的にフォルダを作成する行為は危険性があるため、親フォルダがない場合はポップアップを表示させる
Public Function BackupFile(Path As String, fileName As String, bkupPath As String, bkupFileName As String) As Boolean
    LogApiIn "SaveBackupFile()"
    
    Dim parentDir As String
    parentDir = GetParentDir(bkupPath)
    If Dir(parentDir, vbDirectory) = "" Then
        Dim pressed
        pressed = MsgBox("指摘されたバックアップ先の親フォルダーがありません。" & vbCrLf & "フォルダーを再帰的に作成しますか？" & vbCrLf & parentDir, vbOKCancel)
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

' @breif 親のディレクトリパスを取得する
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

' @breif 階層的なディレクトリをまとめて作成する
' @note この関数は取り扱いが危険なのでポップアップを出したほうがいいかも
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

' @breif ファイルをコピーする
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


' @breif ファイルの最終更新日時を取得する
Public Function GetDateLastModified(FilePath As String, ByRef lastModified As Date) As Boolean
    LogApiIn "GetDateLastModified()"
    
    If Dir(FilePath) = "" Then
        GetDateLastModified = False
        Exit Function
    End If
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject ' インスタンス化
    
    Dim f As File
    Set f = fso.GetFile(FilePath) ' ファイルを取得
    
    lastModified = f.dateLastModified ' 更新日時を取得
    
    GetDateLastModified = True
    LogApiOut "GetDateLastModified()"
End Function

