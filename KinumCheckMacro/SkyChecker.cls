VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SkyChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Password As String = "pass"

Private path_ As String
Private backupPath_ As String
Private fileName_ As String
Private fullName_ As String
Private name_ As String
Private post_ As String
Private employeeId_ As String

' @breif コンストラクタ
Private Sub Class_Initialize()
    LogApiIn "SkyChecker:Class_Initialize()"
    
    LogApiOut "SkyChecker:Class_Initialize()"
End Sub

' @breif デストラクタ
Private Sub Class_Terminate()
    LogApiIn "SkyChecker:Class_Terminate()"

    LogApiOut "SkyChecker:Class_Terminate()"
End Sub

Public Function Initialize(Path As String, BackupPath As String, fileName As String, name As String, post As String, employeeId As String)

    path_ = Path
    backupPath_ = BackupPath
    fileName_ = fileName
    fullName_ = GenerateFullName(Path, fileName)
    name_ = name
    post_ = post
    employeeId_ = employeeId

End Function

' @breif 勤務表をチェックする
Public Function Check()
    LogApiIn "Check()"
    
    ' ファイルをバックアップする
    BackupFile path_, fileName_, backupPath_, fileName_
    
    ' ファイルを開く
    Dim retOpened As Boolean
    retOpened = OpenWorkbook(backupPath_, fileName_)
    
    If retOpened = False Then
        AddResult "Error", name_, "ファイルを開けませんでした。", fullName_
        Exit Function
    End If
    
    Dim dateLastModified As Date
    dateLastModified = GetDateLastModified(fullName_, dateLastModified)
    AddResult "Info", name_, "最終更新日時(" & format(dateLastModified, "yyyy/mm/dd hh:nn") & ")のファイルをチェックしました。", fullName_
        
    ' チェックに必要なデータを取り込む
    ImportWorkbook
    
    ' ファイルを閉じる
    CloseWorkbook (fileName_)
    
    ' 勤務表チェック
    Check_sample
    
    ' エラー件数と警告件数を出したいけどループ時間がエグそう
    Dim countError As Long
    Dim countWarning As Long
    countError = GetCountReuslt("Error", name_)
    countWarning = GetCountReuslt("Warning", name_)
    AddResult "Info", name_, "ファイルをチェックしました。(エラー:" & countError & "件, 警告:" & countWarning & "件)", fullName_
    
    LogApiOut "Check()"
End Function

Private Function ImportWorkbook()
    LogApiIn "ImportWorkbook()"
    LogApiOut "ImportWorkbook()"
End Function

Private Function Check_sample()
    LogApiIn "Check_sample()"

    AddResult "Error", name_, "何かしらのエラーが発生しました。", fullName_

    LogApiOut "Check_sample()"
End Function

