Attribute VB_Name = "Main"
Option Explicit

' @breif メンバ一覧_Skyの列情報
Private Enum ColMemberSky
    IsChecked = 1
    EmplyoeeId
    MemberName
    post
    MainJob
    SubJob
    Empty1
    Empty2
    Empty3
    Prohibited
    fileName
End Enum

' @breif メンバ一覧_Skyの行情報
Private Enum RowMemberSky
    heading = 1
    First
End Enum

' @breif メンバ一覧_BPの列情報
Private Enum ColMemberPartner
    IsChecked = 1
    EmplyoeeId
    MemberName
    Corporate
    MainJob
    SubJob
    Empty1
    Empty2
    Empty3
    Prohibited
    fileName
End Enum

' @breif メンバ一覧_BPの行情報
Private Enum RowMemberBP
    heading = 1
    First
End Enum

Const FooterSizeMemberBP As Integer = 1

Public Sub ButtonCheckAll()
    SetUp
    LogInfo "Pressed Button(CheckAll)"
    CheckSky
    CheckPartner
    TearDown
End Sub

Public Sub ButtonCheckSky()
    SetUp
    LogInfo "Pressed Button(CheckSky)"
    CheckSky
    TearDown
End Sub

Public Sub ButtonCheckPartner()
    SetUp
    LogInfo "Pressed Button(Partner)"
    CheckPartner
    TearDown
End Sub

' @breif Skyの勤務表メンバをチェックする
Private Function CheckSky()
    LogApiIn "CheckSky()"
    LogInfo "Start Check(Sky)"
    
    ' 勤務表チェック
    Dim memberList As Variant
    memberList = Worksheets("チェック").Range("メンバ一覧_Sky")
    
    ' チェック対象(IsChecked = True)なら勤務表をチェックする
    Const FooterSizeMemberSky As Integer = 1
    Dim i As Long
    For i = RowMemberSky.First To UBound(memberList, 1) - FooterSizeMemberSky
        If Not IsEmpty(memberList(i, ColMemberSky.IsChecked)) And memberList(i, ColMemberSky.IsChecked) Then
            Dim checker As SkyChecker
            Set checker = New SkyChecker
            
            Dim fileName As String
            Dim Name As String
            Dim post As String
            Dim employeeId As String
            fileName = memberList(i, ColMemberSky.fileName)
            Name = memberList(i, ColMemberSky.MemberName)
            post = memberList(i, ColMemberSky.post)
            employeeId = memberList(i, ColMemberSky.EmplyoeeId)
            checker.Initialize GetTargetPath, GetBackupPath, fileName, Name, post, employeeId
            
            checker.Name = memberList(i, ColMemberSky.MemberName)
            checker.FullName = GenerateFullName(GetTargetPath, CStr(memberList(i, ColMemberSky.fileName)))
            checker.Year = GetTargetYear
            checker.Month = GetTargetMonth
            
            checker.Check
        End If
    Next i

    GetWorkDayCount 2020, 9, True

    ' ログ出力
    OutputResult
    
    LogInfo "End Check(Sky)"
    LogApiOut "CheckSky()"
End Function


' @breif BPの勤務表をチェックする
Private Function CheckPartner()
    LogApiIn "CheckPartner()"
    LogInfo "Start Check(Partner)"
    
    ' 勤務表チェック
    Dim memberList As Variant
    memberList = Worksheets("チェック").Range("メンバ一覧_BP")
    
    ' チェック対象(IsChecked = True)なら勤務表をチェックする
    Dim i As Long
    For i = RowMemberBP.First To UBound(memberList, 1) - FooterSizeMemberBP
        If Not IsEmpty(memberList(i, ColMemberBP.IsChecked)) And memberList(i, ColMemberBP.IsChecked) Then
            Dim checker As PartnerChecker
            Set checker = New PartnerChecker
            
            Dim fileName As String
            Dim Name As String
            Dim position As String
            Dim employeeId As String
            fileName = memberList(i, ColMemberSky.fileName)
            Name = memberList(i, ColMemberSky.MemberName)
            position = memberList(i, ColMemberSky.post)
            employeeId = memberList(i, ColMemberSky.EmplyoeeId)
            checker.Initialize GetTargetPath, GetBackupPath, fileName, Name, position, employeeId
            
            checker.Check
        End If
    Next i

    
    LogInfo "End Check(Partner)"
    LogApiIn "CheckPartner()"
End Function

' @breif マクロ開始準備
Private Function SetUp()
    Logger_Initialize
    LogApiIn "SetUp()"

    ' VBA高速化 Setup
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    ' 各種初期化処理呼び出し
    CheckedResult_Initialize
    ScheduleProvider_Initialize
    SettingProvider_Initialize
    
    LogApiOut "SetUp()"
End Function

Private Function TearDown()
    LogApiIn "TearDown()"

    ' 各種終了処理呼び出し
    CheckedResult_Terminate
    ScheduleProvider_Terminate
    SettingProvider_Terminate

    ' VBA高速化 Teardown
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    LogApiOut "TearDown()"
    Logger_Terminate
End Function
