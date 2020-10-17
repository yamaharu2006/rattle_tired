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

Private post_ As String

Private commonChecker As commonChecker
Private data_ As SkyWorkScheduleBook

' Setter/Getter
Property Let Year(ByVal argYear As Long)
    commonChecker.Year = argYear
End Property

Private Property Get Year() As Long
    Year = commonChecker.Year
End Property

Property Let Month(ByVal argMonth As Long)
    commonChecker.Month = argMonth
End Property

Private Property Get Month() As Long
    Month = commonChecker.Month
End Property

Property Let EmploeeName(ByRef argName As String)
    commonChecker.EmploeeName = argName
End Property

Private Property Get EmploeeName() As String
    EmploeeName = commonChecker.EmploeeName
End Property

Property Let FullName(ByRef argFullName As String)
    commonChecker.FullName = argFullName
End Property

Private Property Get FullName() As String
    FullName = commonChecker.FullName
End Property

Property Let Path(ByRef arg As String)
    commonChecker.Path = arg
End Property

Private Property Get Path() As String
    Path = commonChecker.Path
End Property

Property Let BackupPath(ByRef arg As String)
    commonChecker.BackupPath = arg
End Property

Private Property Get BackupPath() As String
    BackupPath = commonChecker.BackupPath
End Property

Property Let FileName(ByRef arg As String)
    commonChecker.FileName = arg
    data_.WorkbookName = FileName
End Property

Private Property Get FileName() As String
    FileName = commonChecker.FileName
End Property

Property Let Post(ByRef arg As String)
    post_ = arg
End Property

Property Let EmploeeId(ByRef argId As String)
    commonChecker.EmploeeId = argId
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

' @breif 勤務表をチェックする
Public Function Check()
    LogApiIn "Check()"
    
    ' ファイルがなければチェックしない
    If Not ExistsFile(FullName) Then
        AddResult ErrorLog, EmploeeName, "ファイルを開けませんでした。", FullName
        LogApiOut "Check()"
        Exit Function
    End If
    
    SetUpChecking
    
    ' 勤務表チェック
    Check_Sheets
    Check_DaysData
    
    TearDownChecking
    
    LogApiOut "Check()"
End Function

' @breif チェックの準備を行う
' @note 先に勤務表のデータをインポートすることでシートにアクセスする回数をへらす
Private Function SetUpChecking()
    LogApiIn "SetUpChecking()"
    
    ' ファイルをバックアップする
    BackupFile Path, FileName, BackupPath, FileName
    
    ' ファイルを開く
    OpenWorkbook BackupPath, FileName
    
    ' チェックに必要なデータを取り込む
    data_.ImportData
    
    ' ファイルを閉じる
    CloseWorkbook (FileName)
    
    ' チェック結果に最終更新日時を出力
    Dim dateLastModified As Date
    GetDateLastModified FullName, dateLastModified
    AddResult InfoLog, EmploeeName, "最終更新日時(" & Format(dateLastModified, "yyyy/mm/dd hh:nn") & ")のファイルをチェックしました。", FullName

    LogApiOut "SetUpChecking()"
End Function

' @breif チェックの後処理を行う
Private Function TearDownChecking()
    LogApiIn "TearDownChecking()"
    
    ' エラー件数と警告件数を出す。ループ時間がエグいので、出力有無をSwitchしたい
    Dim countError As Long
    Dim countWarning As Long
    countError = GetCountReuslt(ErrorLog, EmploeeName)
    countWarning = GetCountReuslt(WarningLog, EmploeeName)
    AddResult InfoLog, EmploeeName, "ファイルをチェックしました。(エラー:" & countError & "件, 警告:" & countWarning & "件)", FullName
    
    LogApiOut "TearDownChecking()"
End Function

'////////// 勤務表チェック関数 //////////
' @breif シート単位でチェックをする
Private Function Check_Sheets()
    LogApiIn "Check_Sheets()"
    
    commonChecker.Check_年 data_.Year
    commonChecker.Check_月 data_.Month
    commonChecker.Check_氏名 data_.氏名
    commonChecker.Check_社員ID data_.社員ID
    commonChecker.Check_稼働日数 data_.稼働日数
    
    Check_役職 post_, data_.役職
    
    LogApiOut "Check_Sheets()"
End Function


' @breif 一日単位でチェックする
' @note Let/Set を使うとFor Eachできないのか... ﾊｧ... ﾌﾞｲﾋﾞｰｴｰｻﾝ...
Private Function Check_DaysData()
    LogApiIn "Check_DaysData()"
    
    Dim i As Long
    
    ' 一日分に対してチェックする
    For i = 1 To GetLastDayOfMonth(Year, Month)
        Dim dayData As WorkScheduleDayData
        Set dayData = data_.勤務データ(i)
        
        ' >>>>> 勤務表チェック(一日単位)ここから
        commonChecker.Check_労働基準法 dayData
        Check_備考欄通勤費記載 dayData
        
        ' <<<<< 勤務表チェック(一日単位)ここまで
        
    Next i
    
    LogApiOut "Check_DaysData()"
End Function

Private Function Check_役職(expected As String, actual As String)
    LogApiIn "Check_備考欄通勤費記載()"
    If expected <> actual Then
        AddResult ErrorLog, EmploeeName, "役職が間違っています。(期待値:" & expected & ", 入力:" & actual & ")", FullName
    End If
    LogApiOut "Check_備考欄通勤費記載()"
End Function


' @note 現状は簡素なチェック。ちゃんとしようとするならテレワークの実態と照らし合わせが必要。Mendy
Private Function Check_備考欄通勤費記載(ByRef actual As WorkScheduleDayData)
    LogApiIn "Check_備考欄通勤費記載()"

    If actual.総勤務時間 <> 0 Then
        Select Case True
        Case actual.備考 Like "【通勤費:往復】"
        Case actual.備考 Like "【通勤費:往のみ】"
        Case actual.備考 Like "【通勤費:復のみ】"
        Case actual.備考 Like "【通勤費:なし】"
        Case actual.備考 Like "【通勤費:なし】【テレワーク】"
        Case Else
            AddResult ErrorLog, EmploeeName, "以下のいずれかを勤務表の備考欄に記載してください(" & actual.日 & "日)" & vbCrLf & _
            "【通勤費:往復】/【通勤費:往のみ】/【通勤費:復のみ】/【通勤費:なし】/【通勤費:なし】【テレワーク】", FullName
        End Select
    End If
    
    LogApiOut "Check_備考欄通勤費記載()"
End Function

