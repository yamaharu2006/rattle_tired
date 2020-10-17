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

Private ”N_ As Long
Private Œ_ As Long
Private “ú_ As Long
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

Property Get “ú•t()
    Dim strDate As String
    strDate = ”N_ & "/" & Œ_ & "/" & “ú_
    If IsDate(strDate) Then
        “ú•t = CDate(strDate)
    End If
End Property

Property Let ”N(arg As Long)
    ”N_ = arg
End Property

Property Let Œ(arg As Long)
    Œ_ = arg
End Property

Property Let “ú(arg As Long)
    “ú_ = arg
End Property

Property Get “ú() As Long
    “ú = “ú_
End Property

Property Let oĞ(time As Variant)
    If Not IsEmpty(time) Then
        oĞ_ = time
    End If
End Property

Property Let ‘ŞĞ(time As Variant)
    If Not IsEmpty(time) Then
        ‘ŞĞ_ = time
    End If
End Property

Property Let ‹Î‘Ó(i As Long, val As String)
    ‹Î‘Ó_(i) = val
End Property

Property Get ‹Î‘Ó(i As Long) As String
    ‹Î‘Ó = ‹Î‘Ó_(i)
End Property

Property Let •‹xŒe(time As Variant)
    If Not IsEmpty(time) Then
        •‹xŒe_ = time
    End If
End Property

Property Get •‹xŒe() As Variant
    •‹xŒe = •‹xŒe_
End Property

Property Let [‹xŒe(time As Variant)
    If Not IsEmpty(time) Then
        [‹xŒe_ = time
    End If
End Property

Property Get [‹xŒe() As Variant
    [‹xŒe = [‹xŒe_
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

Property Get ŠÔŠO‹Î–±’©c‹Æ_•‹xŒe() As Variant
    ŠÔŠO‹Î–±’©c‹Æ_•‹xŒe = ŠÔŠO‹Î–±’©c‹Æ_•‹xŒe_
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

Property Get ŠÔŠO‹Î–±_•‹xŒe() As Variant
    ŠÔŠO‹Î–±_•‹xŒe = ŠÔŠO‹Î–±_•‹xŒe_
End Property

Property Let ŠÔŠO‹Î–±_[‹xŒe(time As Variant)
    If Not IsEmpty(time) Then
        ŠÔŠO‹Î–±_[‹xŒe_ = time
    End If
End Property

Property Get ŠÔŠO‹Î–±_[‹xŒe() As Variant
    ŠÔŠO‹Î–±_[‹xŒe = ŠÔŠO‹Î–±_[‹xŒe_
End Property

Property Let ‘‹Î–±ŠÔ(time As Variant)
    If Not IsEmpty(time) Then
        ‘‹Î–±ŠÔ_ = time
    End If
End Property

Property Get ‘‹Î–±ŠÔ() As Variant
    ‘‹Î–±ŠÔ = ‘‹Î–±ŠÔ_
End Property


