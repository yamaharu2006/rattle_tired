Attribute VB_Name = "Main"
Option Explicit

Sub CheckAll()
    LogApiIn "CheckAll()"

    
    
    LogApiOut "CheckAll()"
End Sub

Sub CheckSky()
    LogApiIn "CheckSky()"
    
    ' 事前準備
    SetUp
    
    ' 勤務表チェック
    Dim checker As SkyChecker
    Set checker = New SkyChecker
    checker.Check

    GetWorkDayCount 2020, 9, True

    ' ログ出力
    OutputResult

    ' 事後処理
    TearDown
    
    LogApiOut "CheckSky()"
End Sub

Sub CheckPartner()
    LogApiIn "CheckPartner()"
    
    StartUp

    TearDown
    
    LogApiIn "CheckPartner()"
End Sub

' @note 二度読んでも実害はない(と思う)のでブロック処理は作らない
Private Function SetUp()
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
End Function
