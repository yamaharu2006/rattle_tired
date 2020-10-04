Attribute VB_Name = "FileSystemWrapper"
' @breif ファイル操作を平易に扱うために用意したライブラリ
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
Public Function ConvertPathToFileName(path As String) As String
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

' ファイルをバックアップする
' @note バグ：バックアップのパスが2階層以上フォルダを作る必要があったとき動かない
Public Function BackupFile(path As String, fileName As String, bkupPath As String, bkupFileName As String)
    LogApiIn "SaveBackupFile()"
    
    MkDirRecursive bkupPath
    CopyBackupFile path, fileName, bkupPath, bkupFileName
    
    LogApiOut "SaveBackupFile()"
End Function


' @breif 階層的なディレクトリをまとめて作成する
' @note この関数は取り扱いが危険なのでポップアップを出したほうがいいかも
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

' @breif ファイルをコピーする
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


' @breif ファイルの最終更新日時を取得する
Public Function GetDateLastModified(FilePath As String) As Date
    LogApiIn "GetDateLastModified()"
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject ' インスタンス化
    
    Dim f As File
    Set f = fso.GetFile(FilePath) ' ファイルを取得
    
    Dim lastModified As Date
    lastModified = f.dateLastModified ' 更新日時を取得
    
    GetDateLastModified = lastModified
    
    LogApiOut "GetDateLastModified()"
End Function
