VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SkyWorkScheduleBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @breif Sky社員向け勤務表データ
Option Explicit

Private Enum 出勤情報_列情報
    日列 = 1
    曜日列
    出勤印列
    勤務時間_出社列
    勤務時間_退社列
    勤怠列1
    勤怠列2
    備考列
    空白列1
    定時時間_普休憩列
    定時時間_深休憩列
    実働列
    定時後休憩列
    時間外勤務_開始列
    時間外勤務_終了列
    時間外勤務_普休憩列
    時間外勤務_深休憩列
    時間外勤務_実施列
    時間外勤務_普通列
    時間外勤務_休日列
    時間外勤務_深夜列
    総勤務時間列
    法定超過列
    所属長印列
End Enum

Const Range年 As String = "D1"
Const Range月 As String = "B2"
Const Range稼働日数 As String = "D3"
Const Rangeチーム As String = "G2"
Const Range役職 As String = "O1"
Const Rangeグループ As String = "O2"
Const Range社員ID As String = "U1"
Const Range氏名 As String = "U2"

Const Range勤務データ As String = "A7:X68"

Private WorkbookName_ As String

' [総務課提出]シート
'Const SheetFiling As String = "総務課提出"
Private 年_ As Long
Private 月_ As Long
Private 稼働日数_ As Long
Public チーム As String
Public 役職 As String
Public グループ As String
Public 所属 As String
Public 社員ID As String
Public 氏名 As String

Private Const DaysDataArrayMax As Long = 32 ' 勤務データ配列の最大要素数(31+1)
Private 勤務データ_(1 To DaysDataArrayMax) As WorkScheduleDayData ' Index = 日になるように配列を設定する

' Setter/Getter
Property Let WorkbookName(Name As String)
    If Name <> "" Then
        WorkbookName_ = Name
    End If
End Property

Property Let Year(arg As Variant)
    If Not IsEmpty(arg) Then
        年_ = arg
    End If
End Property

Property Get Year() As Variant
    Year = 年_
End Property

Property Let Month(arg As Variant)
    If Not IsEmpty(arg) Then
        月_ = arg
    End If
End Property

Property Get Month() As Variant
    Month = 月_
End Property

Property Let 稼働日数(arg As Long)
    If Not IsEmpty(arg) Then
        稼働日数_ = arg
    End If
End Property

Property Get 稼働日数() As Long
    稼働日数 = 稼働日数_
End Property

Property Get 勤務データ(i As Long) As WorkScheduleDayData
    Set 勤務データ = 勤務データ_(i)
End Property

' @breif コンストラクタ
Private Sub Class_Initialize()
    LogApiIn "SkyWorkScheduleBook:Class_Initialize()"
    
    Dim i As Long
    For i = 1 To UBound(勤務データ_)
        Set 勤務データ_(i) = New WorkScheduleDayData
    Next i
    
    LogApiOut "SkyWorkScheduleBook:Class_Initialize()"
End Sub

' @breif デストラクタ
Private Sub Class_Terminate()
    LogApiIn "SkyChecker:Class_Terminate()"
    
    Dim i As Long
    For i = 1 To UBound(勤務データ_)
        Set 勤務データ_(i) = Nothing
    Next i

    LogApiOut "SkyChecker:Class_Terminate()"
End Sub

' @breif データをインポートする
Public Function ImportData()
    LogApiIn "KinmuDataSky:ImportData()"
    
    With Workbooks(WorkbookName_).Sheets("総務課提出")
        年_ = .Range(Range年)
        月_ = .Range(Range月)
        稼働日数_ = .Range(Range稼働日数)
        チーム = .Range(Rangeチーム)
        役職 = .Range(Range役職)
        グループ = .Range(Rangeグループ)
        社員ID = .Range(Range社員ID)
        氏名 = .Range(Range氏名)
        
        Dim var As Variant
        var = .Range(Range勤務データ)
        Dim i As Long
        For i = 1 To UBound(var) Step 2
            Dim index As Long
            index = Int(i / 2 + 1)
            Dim upper As Long ' 上段
            upper = i
            Dim lower As Long ' 下段
            lower = i + 1
            
            勤務データ_(index).年 = 年_
            勤務データ_(index).月 = 月_
            勤務データ_(index).日 = var(lower, 日列)
            勤務データ_(index).出社 = var(lower, 勤務時間_出社列)
            勤務データ_(index).退社 = var(lower, 勤務時間_退社列)
            勤務データ_(index).勤怠(0) = var(upper, 勤怠列1)
            勤務データ_(index).勤怠(1) = var(lower, 勤怠列1)
            勤務データ_(index).勤怠(2) = var(upper, 勤怠列2)
            勤務データ_(index).勤怠(3) = var(lower, 勤怠列2)
            勤務データ_(index).備考 = var(upper, 備考列)
            勤務データ_(index).普休憩 = var(lower, 定時時間_普休憩列)
            勤務データ_(index).深休憩 = var(lower, 定時時間_深休憩列)
            勤務データ_(index).時間外勤務朝残業_開始 = var(upper, 時間外勤務_開始列)
            勤務データ_(index).時間外勤務朝残業_終了 = var(upper, 時間外勤務_終了列)
            勤務データ_(index).時間外勤務朝残業_普休憩 = var(upper, 時間外勤務_普休憩列)
            勤務データ_(index).時間外勤務_開始 = var(lower, 時間外勤務_開始列)
            勤務データ_(index).時間外勤務_普休憩 = var(lower, 時間外勤務_普休憩列)
            勤務データ_(index).時間外勤務_深休憩 = var(lower, 時間外勤務_終了列)
            勤務データ_(index).総勤務時間 = var(lower, 総勤務時間列)
        Next i
    End With
    
    LogApiOut "KinmuDataSky:ImportData()"
End Function

