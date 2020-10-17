Attribute VB_Name = "FileSystemWrapper"
' @breif 勤務表ファイル操作を平易に扱うために用意したライブラリ
' @note ファイル操作全般を扱うライブラリだとLogについても考える必要が生じるので勤務表ファイルに限定したい
Option Explicit

Private Const filePassword = "pass"

' @breif ファイルを開く
' @note 下回りでエラーを拾うのか割とぐちゃぐちゃになってきている
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

' @breif フルパスからファイル名を取得する
' @note working...
Public Function ConvertPathToFileName(Path As String) As String
    LogApiIn "OpenWorkbook()"
    
    ConvertPathToFileName = ""
    
    LogApiIn "OpenWorkbook()"
End Function

' @breif 絶対パスを生成する
Public Function GenerateFullName(folderPath As String, FileName As String) As String
    LogApiIn "OpenWorkbook()"
    
    GenerateFullName = folderPath & "\" & FileName
    
    LogApiIn "OpenWorkbook()"
End Function

' @breif ファイルが存在しているかを返す
Public Function ExistsFile(FullName As String) As Boolean
    If Dir(FullName) = "" Then
        ExistsFile = False
    Else
        ExistsFile = True
    End If
End Function

' @breif 同名の勤務表ファイルを開いているかを返す
' @attention ブックを開いていれば開いているほどチェックに時間がかかる
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

' @breif ファイルをクローズする
' @note ファイル操作系関数群は別クラスに委譲したい
Public Function CloseWorkbook(FileName As String)
    LogApiIn "CloseWorkbook()"
    
    If IsOpenedSameFile(FileName) = True Then
        Workbooks(FileName).Close
    End If
    
    LogApiOut "CloseWorkbook()"
End Function

' @breif  ファイルをバックアップする
' @attention 再帰的にフォルダを作成する行為は危険を伴うため、親フォルダがない場合はポップアップを表示させる
Public Function BackupFile(Path As String, FileName As String, bkupPath As String, bkupFileName As String) As Boolean
    LogApiIn "SaveBackupFile()"
    
    Dim parentDir As String
    parentDir = GetParentDir(bkupPath)
    If Dir(parentDir, vbDirectory) = "" Then
        Dim pressed
        pressed = MsgBox("指摘されたバックアップ先の親フォルダーがありません。" & vbCrLf & "フォルダーを再帰的に作成しますか？" & vbCrLf & parentDir, vbOKCancel)
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


' @breif ファイルの最終更新日時を取得する
Public Function GetDateLastModified(FilePath As String, ByRef lastModified As Date) As Boolean
    LogApiIn "GetDateLastModified()"
    
    If Dir(FilePath) = "" Then
        GetDateLastModified = False
        LogApiOut "GetDateLastModified()"
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

