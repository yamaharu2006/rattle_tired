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

Private year_ As Integer
Private month_ As Integer
Private fullName_ As String
Private name_ As String

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

Property Let Month(argMonth As Long)
    month_ = argMonth
End Property

Property Let FullName(argFullName As String)
    fullName_ = argFullName
End Property

Property Let Name(argName As String)
    name_ = argName
End Property

' 基本入力
Public Function Check_氏名(ByRef expected As String, ByRef actual As String)
    LogApiIn "Check_氏名()"
    If expected <> actual Then
        AddResult ErrorLog, name_, "名前が間違っています(入力:" & actual & ")", fullName_
    End If
    LogApiOut "Check_氏名()"
End Function

Public Function Check_社員ID(ByRef expected As String, ByRef actual As String)
    LogApiIn "Check_社員ID()"
    If expected <> actual Then
        AddResult ErrorLog, name_, "社員IDが間違っています(入力:" & actual & ")", fullName_
    End If
    LogApiOut "Check_社員ID()"
End Function

Public Function Check_年(ByVal expected As Integer, ByVal actual As Integer)
    LogApiIn "Check年()"
    If expected <> actual Then
        AddResult ErrorLog, name_, "入力年が間違っています(入力:" & actual & ")", fullName_
    End If
    LogApiOut "Check_年()"
End Function

Public Function Check_月(ByVal expected As Integer, ByVal actual As Integer)
    LogApiIn "Check_月()"
    If expected <> actual Then
        AddResult ErrorLog, name_, "入力月が間違っています(入力:" & actual & ")", fullName_
    End If
    LogApiOut "Check_月()"
End Function

Public Function Check_稼働日数(ByVal actual As Integer)
    LogApiIn "Check_稼働日数()"
    Dim expected As Long
    expected = GetWorkDayCount(year_, month_, isSky_)
    If expected <> actual Then
        AddResult ErrorLog, name_, "稼働日数が間違っています(入力:" & actual & ")" & vbCrLf _
        & year_ & "年" & month_ & "月の稼働日数は" & expected & "日です。", fullName_
    End If
    LogApiOut "Check_稼働日数()"
End Function

' 一日の勤怠データ
Public Function Check_労働基準法(ByRef actual As String)
    LogApiIn "Check_労働基準法()"
    If expected <> actual Then
        AddResult WarningLog, name_, "労働基準法を違反しています。休憩時間を確認してください。(入力:" & actual & ")" & vbCrLf _
        , fullName_
    End If
    LogApiOut "Check_労働基準法()"
End Function
