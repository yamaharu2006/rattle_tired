VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @breif 「Sky勤務表」「作業時間報告書」のいずれでもチェックする項目をまとめて定義する
' @note 未練がましいけど抽象クラスを作れるならな... ぐちゃぐちゃもじゅーる構成気色悪い
Option Explicit

Private year_ As Long
Private month_ As Long
Private fullName_ As String
Public Path As String
Public BackupPath As String
Public FileName As String
Public EmploeeName As String
Public EmploeeId As String

Private isSky_ As Boolean ' 割と屈辱のプロパティである

'Private data_ As SkyWorkScheduleBook ' プログラム的にはこうするべきなんだけど、メモリ消費が気になる(Bookのインタフェースも必要なのであとでやる)

' @breif コンストラクタ
Private Sub Class_Initialize()
    LogApiIn "CommonChecker:Class_Initialize()"
    LogApiOut "CommonChecker:Class_Initialize()"
End Sub

' @breif デストラクタ
Private Sub Class_Terminate()
    LogApiIn "CommonChecker:Class_Terminate()"
    LogApiOut "CommonChecker:Class_Terminate()"
End Sub

' Setter/Getter
Property Let Year(argYear As Long)
    year_ = argYear
End Property

Property Get Year() As Long
    Year = year_
End Property

Property Let Month(argMonth As Long)
    month_ = argMonth
End Property

Property Get Month() As Long
    Month = month_
End Property

Property Get FullName() As String
    If fullName_ = "" Then
        fullName_ = GenerateFullName(Path, FileName)
    End If
    FullName = fullName_
End Property

' 基本入力チェック
Public Function Check_氏名(actual As String)
    LogApiIn "Check_氏名()"
    If EmploeeName <> actual Then
        AddResult ErrorLog, EmploeeName, "名前が間違っています(期待値:" & year_ & ", 入力:" & actual & ")", fullName_
    End If
    LogApiOut "Check_氏名()"
End Function

Public Function Check_社員ID(actual As String)
    LogApiIn "Check_社員ID()"
    If EmploeeId <> actual Then
        AddResult ErrorLog, EmploeeName, "社員IDが間違っています(期待値:" & EmploeeName & ", 入力:" & actual & ")", fullName_
    End If
    LogApiOut "Check_社員ID()"
End Function

Public Function Check_年(ByVal actual As Long)
    LogApiIn "Check年()"
    If year_ <> actual Then
        AddResult ErrorLog, EmploeeName, "入力年が間違っています(期待値:" & year_ & ", 入力:" & actual & ")", fullName_
    End If
    LogApiOut "Check_年()"
End Function

Public Function Check_月(ByVal actual As Long)
    LogApiIn "Check_月()"
    If month_ <> actual Then
        AddResult ErrorLog, EmploeeName, "入力月が間違っています(期待値:" & month_ & ", 入力:" & actual & ")", fullName_
    End If
    LogApiOut "Check_月()"
End Function

Public Function Check_稼働日数(ByVal actual As Long)
    LogApiIn "Check_稼働日数()"
    Dim expected As Long
    expected = GetWorkDayCount(year_, month_, isSky_)
    If expected <> actual Then
        AddResult ErrorLog, EmploeeName, "稼働日数が間違っています(入力:" & actual & ")" & vbCrLf _
        & year_ & "年" & month_ & "月の稼働日数は" & expected & "日です。", fullName_
    End If
    LogApiOut "Check_稼働日数()"
End Function

' 一日の勤怠データに対するチェック
' @breif 労働基準法の休憩時間に従っているかどうかをチェックする
'        ・8時間以上の勤務 -> 最低1時間の休憩時間
'        ・6時間以上8時間未満の勤務 -> 最低45分の休憩時間
' @note ハードコーディングがいやでテーブルにしようかと思ったけど逆にわかりづらくなる
Public Function Check_労働基準法(ByRef actual As WorkScheduleDayData)
    LogApiIn "Check_労働基準法()"
    
    Dim actualBreakTime As Date
    With actual
        actualBreakTime = .普休憩 + .時間外勤務_普休憩 + .時間外勤務_深休憩 + .時間外勤務朝残業_普休憩
    End With
    
    If actual.総勤務時間 >= #8:00:00 AM# And actualBreakTime < #1:00:00 AM# Then
        AddResult ErrorLog, EmploeeName, "労働基準法に違反しています。総勤務時間が1時間を超える場合は最低1時間の休憩を取得してください。(" & actual.日 & "日)", fullName_
    ElseIf actual.総勤務時間 >= #6:00:00 AM# And actualBreakTime < #12:45:00 AM# Then
        AddResult ErrorLog, EmploeeName, "労働基準法に違反しています。総勤務時間が6時間を超える場合は最低45分の休憩を取得してください。(" & actual.日 & "日)", fullName_
    End If

    LogApiOut "Check_労働基準法()"
End Function
