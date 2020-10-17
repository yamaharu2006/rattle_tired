Attribute VB_Name = "ScheduleProvider"
Option Explicit

' @class カレンダー提供クラス
' @breif カレンダー情報を提供する
' @note 速度を考えるなら年月を指定させるべき
' @note 一つだけ存在すればいいんだから標準モジュールでもよかったのかな

' @breif Range取得したVariant型配列にアクセスするためのEnum
Private Enum SceduleRow
    Heading1 = 1
    Heading2
    First
End Enum

Private Enum SceduleColomun
    Pading1 = 1
    Date
    SceduleName
    SceduleType
    Pading2
End Enum

' シート情報
Private Const SceduleSheetName As String = "カレンダー"
Private Const SceduleRangeName As String = "スケジュール"

' define
Private daysInfo As Dictionary
Private workDayCount As Dictionary

' @breif 初期化関数
Public Function ScheduleProvider_Initialize()
    LogApiIn "ScheduleProvider_Initialize()"
    
    Set daysInfo = New Dictionary
    Set workDayCount = New Dictionary
    
    SetDayInfo
    
    LogApiOut "ScheduleProvider_Initialize()"
End Function

' @breif 終了関数
Public Function ScheduleProvider_Terminate()
    LogApiIn "ScheduleProvider_Terminate()"

    Set daysInfo = Nothing
    Set workDayCount = Nothing
    
    LogApiOut "ScheduleProvider_Terminate()"
End Function

' @breif 稼働日を取得する
' @note 稼働日が未計算であれば計算する。計算回数は抑えたい
Public Function GetWorkDayCount(Year As Long, Month As Long, IsSky As Boolean) As Long
    LogApiIn "GetWorkDayCount()"
    
    Dim key  As Date
    key = DateSerial(Year, Month, 1)
    
    If workDayCount.Item(key) = "" Then
        GetWorkDayCount = CalculateWorkDayCount(Year, Month, IsSky)
    Else
        GetWorkDayCount = workDayCount.Item(key)
    End If
    
    LogApiOut "GetWorkDayCount()"
End Function

' @brief 出勤日を計算する
Private Function CalculateWorkDayCount(Year As Long, Month As Long, IsSky As Boolean) As Long
    LogApiIn "CalculateWorkDayCount()"

    Dim count As Long
    count = 0

    Dim lastDayOfMonth As Long
    lastDayOfMonth = GetLastDayOfMonth(Year, Month)
    
    Dim i As Long
    For i = 1 To lastDayOfMonth
        Dim d As Date
        d = DateSerial(Year, Month, i)
        If IsWorkDay(d, IsSky) Then
            count = count + 1
        End If
    Next i
    
    Dim key  As Date
    key = DateSerial(Year, Month, 1)
    workDayCount(key) = count
    
    CalculateWorkDayCount = count
    
    LogApiOut "CalculateWorkDayCount()"
End Function

' @brief 月の最終日を取得する
Public Function GetLastDayOfMonth(argYear As Long, argMonth As Long) As Long
    LogApiIn "CalculateWorkDayCount()"

    Dim lastDay As Date
    lastDay = DateSerial(argYear, argMonth + 1, 0)
    GetLastDayOfMonth = day(lastDay)
    
    LogApiOut "CalculateWorkDayCount()"
End Function

' @breif シートからスケジュール情報を取得する
' @attention Workbookにアクセスする関数のため呼び出し回数に注意
Private Function SetDayInfo()
    LogApiIn "SetDayInfo()"

    Dim rangeDayInfo As Variant
    rangeDayInfo = Worksheets(SceduleSheetName).Range(SceduleRangeName)
    
    Dim i As Long
    For i = SceduleRow.First To UBound(rangeDayInfo, 1) - 1
        
        Dim KeyDate As Date
        Dim DateType As String
        KeyDate = rangeDayInfo(i, SceduleColomun.Date)
        DateType = rangeDayInfo(i, SceduleColomun.SceduleType)
        daysInfo.Add KeyDate, DateType
        
    Next i
    
    LogApiOut "SetDayInfo()"
End Function

' @breif 出勤日かどうかを取得する
' @note バグを埋め込む可能性が一番高い関数。ロジックも汚い
Public Function IsWorkDay(ArgDate As Date, IsSky As Boolean) As Boolean
    LogApiIn "IsWorkDay()"

    Dim dayType As String
    dayType = GetDayType(ArgDate)
        
    IsWorkDay = True
    
    Select Case dayType
    Case "国民の祝日"
        IsWorkDay = False
    Case "Sky式典日"
        If IsSky = False Then
            IsWorkDay = False
        End If
    Case Else
        If (Weekday(ArgDate) = vbSunday) Or (Weekday(ArgDate) = vbSaturday) Then
            IsWorkDay = False
        End If
    End Select
    
    
    LogApiOut "IsWorkDay()"
End Function

' @breif 指定日付情報を取得する
Private Function GetDayType(ArgDate As Date) As String
    LogApiIn "GetDayType()"

    GetDayType = daysInfo.Item(ArgDate)

    LogApiOut "GetDayType()"
End Function



