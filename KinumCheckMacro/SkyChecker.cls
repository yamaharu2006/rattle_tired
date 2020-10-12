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

Private year_ As Long
Private month_ As Long
Private path_ As String
Private backupPath_ As String
Private fileName_ As String
Private fullName_ As String
Private name_ As String
Private post_ As String
Private employeeId_ As String

Private commonChecker As commonChecker
Private data_ As SkyWorkScheduleBook

' Setter/Getter
Property Let Year(ByVal argYear As Long)
    year_ = argYear
    commonChecker.Year = argYear
    data_.Year = argYear
End Property

Property Let Month(ByVal argMonth As Long)
    month_ = argMonth
    commonChecker.Month = argMonth
    data_.Month = argMonth
End Property

Property Let Name(ByRef argName As String)
    name_ = argName
    commonChecker.Name = argName
    data_.氏名 = argName
End Property

Property Let FullName(ByRef argFullName As String)
    fullName_ = argFullName
    commonChecker.FullName = argFullName
End Property

' @breif コンストラクタ
Private Sub Class_Initialize()
    LogApiIn "SkyChecker:Class_Initialize()"
    
    Set data_ = New SkyWorkScheduleBook
    Set commonChecker = New commonChecker
    
    LogApiOut "SkyChecker:Class_Initialize()"
End Sub

' @breif デストラクタ
Private Sub Class_Terminate()
    LogApiIn "SkyChecker:Class_Terminate()"
    
    Set data_ = Nothing
    Set commonChecker = Nothing

    LogApiOut "SkyChecker:Class_Terminate()"
End Sub

' @breif インスタンス初期化
' @note 引数付きコンストラクタの代わり
Public Function Initialize(Path As String, BackupPath As String, fileName As String, Name As String, post As String, employeeId As String)

    'year_ = Year
    'month_ = Month
    path_ = Path
    backupPath_ = BackupPath
    fileName_ = fileName
    fullName_ = GenerateFullName(Path, fileName)
    name_ = Name
    post_ = post
    employeeId_ = employeeId
    
    data_.WorkbookName = fileName_

End Function

' @breif 勤務表をチェックする
Public Function Check()
    LogApiIn "Check()"
    
    ' ファイルがなければチェックしない
    If Not ExistsFile(fullName_) Then
        AddResult ErrorLog, name_, "ファイルを開けませんでした。", fullName_
        LogApiOut "Check()"
        Exit Function
    End If
    
    SetUpChecking
    
    ' 勤務表チェック
    With commonChecker
        .Check_稼働日数 data_.稼働日数
    End With
    
    Check_sample
    Check_DaysData
    
    TearDownChecking
    
    LogApiOut "Check()"
End Function

' @breif チェックの準備を行う
' @note 先に勤務表のデータをインポートすることでシートにアクセスする回数をへらす
Private Function SetUpChecking()
    LogApiIn "SetUpChecking()"
    
    ' ファイルをバックアップする
    BackupFile path_, fileName_, backupPath_, fileName_
    
    ' ファイルを開く
    OpenWorkbook backupPath_, fileName_
    
    ' チェックに必要なデータを取り込む
    data_.ImportData
    
    ' ファイルを閉じる
    CloseWorkbook (fileName_)
    
    ' チェック結果に最終更新日時を出力
    Dim dateLastModified As Date
    GetDateLastModified fullName_, dateLastModified
    AddResult InfoLog, name_, "最終更新日時(" & format(dateLastModified, "yyyy/mm/dd hh:nn") & ")のファイルをチェックしました。", fullName_

    LogApiOut "SetUpChecking()"
End Function

' @breif チェックの後処理を行う
Private Function TearDownChecking()
    LogApiIn "TearDownChecking()"
    
    ' エラー件数と警告件数を出す。ループ時間がエグいので、出力有無をSwitchしたい
    Dim countError As Long
    Dim countWarning As Long
    countError = GetCountReuslt(ErrorLog, name_)
    countWarning = GetCountReuslt(WarningLog, name_)
    AddResult InfoLog, name_, "ファイルをチェックしました。(エラー:" & countError & "件, 警告:" & countWarning & "件)", fullName_
    
    LogApiOut "TearDownChecking()"
End Function

'////////// 勤務表チェック関数 //////////
' @breif 一日単位でチェックする
' @note Let/Set を使うとFor Eachできないのか... ﾊｧ
Private Function Check_DaysData()
    LogApiIn "Check_sample()"
    
    Dim i As Long
    Dim dayData As WorkScheduleDayData
    
    ' 一日分に対してチェックする
    Const DaysDataArrayMax As Long = 32
    For i = 0 To DaysDataArrayMax
        Set dayData = data_.勤務データ(i)
        
        ' >>>>> 勤務表チェック(一日単位)ここから
        Check_備考欄通勤費記載 dayData.備考
        
        ' <<<<< 勤務表チェック(一日単位)ここまで
        
    Next i
    
    LogApiOut "Check_sample()"
End Function

Private Function Check_sample()
    LogApiIn "Check_sample()"

    AddResult ErrorLog, name_, "何かしらのエラーが発生しました。", fullName_

    LogApiOut "Check_sample()"
End Function


' @note 現状簡素なチェック
Private Function Check_備考欄通勤費記載(actual As String)
    LogApiIn "Check_備考欄通勤費記載()"
    
    Select Case True
    Case actual Like "【通勤費:往復】"
    Case actual Like "【通勤費:往のみ】"
    Case actual Like "【通勤費:復のみ】"
    Case actual Like "【通勤費:なし】"
    Case actual Like "【通勤費:なし】"
    Case actual Like "【通勤費:なし】【テレワーク】"
    Case Else
        AddResult ErrorLog, name_, "以下のいずれかを勤務表の備考欄に記載してください", fullName_
    End Select
    
    LogApiOut "Check_備考欄通勤費記載()"
End Function

