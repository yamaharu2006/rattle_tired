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

Private �N_ As Long
Private ��_ As Long
Private ��_ As Long
Private �o��_ As Date
Private �ގ�_ As Date
Private �Α�_(4) As String
Public ���l As String
Private ���x�e_ As Date
Private �[�x�e_ As Date
Private �莞��x�e_ As Date
Private ���ԊO�Ζ����c��_�J�n_ As Date
Private ���ԊO�Ζ����c��_�I��_ As Date
Private ���ԊO�Ζ����c��_���x�e_ As Date
Private ���ԊO�Ζ�_�J�n_ As Date
Private ���ԊO�Ζ�_�I��_ As Date
Private ���ԊO�Ζ�_���x�e_ As Date
Private ���ԊO�Ζ�_�[�x�e_ As Date
Private ���Ζ�����_ As Date

Property Get ���t()
    Dim strDate As String
    strDate = �N_ & "/" & ��_ & "/" & ��_
    If IsDate(strDate) Then
        ���t = CDate(strDate)
    End If
End Property

Property Let �N(arg As Long)
    �N_ = arg
End Property

Property Let ��(arg As Long)
    ��_ = arg
End Property

Property Let ��(arg As Long)
    ��_ = arg
End Property

Property Get ��() As Long
    �� = ��_
End Property

Property Let �o��(time As Variant)
    If Not IsEmpty(time) Then
        �o��_ = time
    End If
End Property

Property Let �ގ�(time As Variant)
    If Not IsEmpty(time) Then
        �ގ�_ = time
    End If
End Property

Property Let �Α�(i As Long, val As String)
    �Α�_(i) = val
End Property

Property Get �Α�(i As Long) As String
    �Α� = �Α�_(i)
End Property

Property Let ���x�e(time As Variant)
    If Not IsEmpty(time) Then
        ���x�e_ = time
    End If
End Property

Property Get ���x�e() As Variant
    ���x�e = ���x�e_
End Property

Property Let �[�x�e(time As Variant)
    If Not IsEmpty(time) Then
        �[�x�e_ = time
    End If
End Property

Property Get �[�x�e() As Variant
    �[�x�e = �[�x�e_
End Property

Property Let �莞��x�e(time As Variant)
    If Not IsEmpty(time) Then
        �莞��x�e_ = time
    End If
End Property

Property Let ���ԊO�Ζ����c��_�J�n(time As Variant)
    If Not IsEmpty(time) Then
        ���ԊO�Ζ����c��_�J�n_ = time
    End If
End Property

Property Let ���ԊO�Ζ����c��_�I��(time As Variant)
    If Not IsEmpty(time) Then
        ���ԊO�Ζ����c��_�I��_ = time
    End If
End Property

Property Let ���ԊO�Ζ����c��_���x�e(time As Variant)
    If Not IsEmpty(time) Then
        ���ԊO�Ζ����c��_���x�e_ = time
    End If
End Property

Property Get ���ԊO�Ζ����c��_���x�e() As Variant
    ���ԊO�Ζ����c��_���x�e = ���ԊO�Ζ����c��_���x�e_
End Property

Property Let ���ԊO�Ζ�_�J�n(time As Variant)
    If Not IsEmpty(time) Then
        ���ԊO�Ζ�_�J�n_ = time
    End If
End Property

Property Let ���ԊO�Ζ�_�I��(time As Variant)
    If Not IsEmpty(time) Then
        ���ԊO�Ζ�_�I��_ = time
    End If
End Property

Property Let ���ԊO�Ζ�_���x�e(time As Variant)
    If Not IsEmpty(time) Then
        ���ԊO�Ζ�_���x�e_ = time
    End If
End Property

Property Get ���ԊO�Ζ�_���x�e() As Variant
    ���ԊO�Ζ�_���x�e = ���ԊO�Ζ�_���x�e_
End Property

Property Let ���ԊO�Ζ�_�[�x�e(time As Variant)
    If Not IsEmpty(time) Then
        ���ԊO�Ζ�_�[�x�e_ = time
    End If
End Property

Property Get ���ԊO�Ζ�_�[�x�e() As Variant
    ���ԊO�Ζ�_�[�x�e = ���ԊO�Ζ�_�[�x�e_
End Property

Property Let ���Ζ�����(time As Variant)
    If Not IsEmpty(time) Then
        ���Ζ�����_ = time
    End If
End Property

Property Get ���Ζ�����() As Variant
    ���Ζ����� = ���Ζ�����_
End Property


