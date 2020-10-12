VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkScheduleDayData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Message
' It is obvious that Japanese(kanji) is better than English in the point of number of character.
' And, You won't understand the creepy English you were forced to translate.
' How in the world do you expect me to translate "teiji-gai kinmu asa-zangyou jikan futsu kyuukei"?
Option Explicit

'Private ”N_ As Long
'Private Œ_ As Long
'Private “ú_ As Long
Private “ú•t_ As Date
Private oĞ_ As Date
Private ‘ŞĞ_ As Date
Private ‹Î‘Ó_(4) As String
Public ”õl As String
Private •‹xŒe_ As Date
Private [‹xŒe_ As Date
Private ’èŒã‹xŒe_ As Date
Private ŠÔŠO‹Î–±’©c‹Æ_ŠJn_ As Date
Private ŠÔŠO‹Î–±’©c‹Æ_I—¹_ As Date
Private ŠÔŠO‹Î–±’©c‹Æ_•‹xŒe_ As Date
Private ŠÔŠO‹Î–±_ŠJn_ As Date
Private ŠÔŠO‹Î–±_I—¹_ As Date
Private ŠÔŠO‹Î–±_•‹xŒe_ As Date
Private ŠÔŠO‹Î–±_[‹xŒe_ As Date
Private ‘‹Î–±ŠÔ_ As Date

Public Function SetDate(argYear As Long, argMonth As Long, argDay As Long)
    Dim strDate As String
    strDate = argYear & "/" & argMonth & "/" & argDay
    If IsDate(strDate) Then
        “ú•t_ = CDate(strDate)
    End If
End Function

Property Let oĞ(time As Variant)
    If Not IsEmpty(time) Then
        “ú•t_ = time
    End If
End Property

Property Let ‘ŞĞ(time As Variant)
    If Not IsEmpty(time) Then
        ‘ŞĞ_ = time
    End If
End Property

Property Let ‹Î‘Ó(var() As Variant)
    Dim i As Long
    For i = 0 To UBound(‹Î‘Ó_) - 1
        ‹Î‘Ó_(i) = var(i)
    Next i
End Property

Property Let •‹xŒe(time As Variant)
    If Not IsEmpty(time) Then
        •‹xŒe_ = time
    End If
End Property

Property Let [‹xŒe(time As Variant)
    If Not IsEmpty(time) Then
        [‹xŒe_ = time
    End If
End Property

Property Let ’èŒã‹xŒe(time As Variant)
    If Not IsEmpty(time) Then
        ’èŒã‹xŒe_ = time
    End If
End Property

Property Let ŠÔŠO‹Î–±’©c‹Æ_ŠJn(time As Variant)
    If Not IsEmpty(time) Then
        ŠÔŠO‹Î–±’©c‹Æ_ŠJn_ = time
    End If
End Property

Property Let ŠÔŠO‹Î–±’©c‹Æ_I—¹(time As Variant)
    If Not IsEmpty(time) Then
        ŠÔŠO‹Î–±’©c‹Æ_I—¹_ = time
    End If
End Property

Property Let ŠÔŠO‹Î–±’©c‹Æ_•‹xŒe(time As Variant)
    If Not IsEmpty(time) Then
        ŠÔŠO‹Î–±’©c‹Æ_•‹xŒe_ = time
    End If
End Property

Property Let ŠÔŠO‹Î–±_ŠJn(time As Variant)
    If Not IsEmpty(time) Then
        ŠÔŠO‹Î–±_ŠJn_ = time
    End If
End Property

Property Let ŠÔŠO‹Î–±_I—¹(time As Variant)
    If Not IsEmpty(time) Then
        ŠÔŠO‹Î–±_I—¹_ = time
    End If
End Property

Property Let ŠÔŠO‹Î–±_•‹xŒe(time As Variant)
    If Not IsEmpty(time) Then
        ŠÔŠO‹Î–±_•‹xŒe_ = time
    End If
End Property

Property Let ŠÔŠO‹Î–±_[‹xŒe(time As Variant)
    If Not IsEmpty(time) Then
        ŠÔŠO‹Î–±_[‹xŒe_ = time
    End If
End Property

Property Let ‘‹Î–±ŠÔ(time As Variant)
    If Not IsEmpty(time) Then
        ‘‹Î–±ŠÔ_ = time
    End If
End Property



