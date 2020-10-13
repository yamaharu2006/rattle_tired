VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @breif �uSky�Ζ��\�v�u��Ǝ��ԕ񍐏��v�̂�����ł��`�F�b�N���鍀�ڂ��܂Ƃ߂Ē�`����
' @note �������܂������ǒ��ۃN���X������Ȃ��... �����Ⴎ���������[��\���C�F����
Option Explicit

Private year_ As Integer
Private month_ As Integer
Private fullName_ As String
Private name_ As String

Private isSky_ As Boolean ' ���Ƌ��J�̃v���p�e�B�ł���

'Private data_ As SkyWorkScheduleBook ' �v���O�����I�ɂ͂�������ׂ��Ȃ񂾂��ǁA����������C�ɂȂ�(Book�̃C���^�t�F�[�X���K�v�Ȃ̂ł��Ƃł��)

' @breif �R���X�g���N�^
Private Sub Class_Initialize()
    LogApiIn "CommonChecker:Class_Initialize()"
    LogApiOut "CommonChecker:Class_Initialize()"
End Sub

' @breif �f�X�g���N�^
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

' ��{����
Public Function Check_����(ByRef expected As String, ByRef actual As String)
    LogApiIn "Check_����()"
    If expected <> actual Then
        AddResult ErrorLog, name_, "���O���Ԉ���Ă��܂�(����:" & actual & ")", fullName_
    End If
    LogApiOut "Check_����()"
End Function

Public Function Check_�Ј�ID(ByRef expected As String, ByRef actual As String)
    LogApiIn "Check_�Ј�ID()"
    If expected <> actual Then
        AddResult ErrorLog, name_, "�Ј�ID���Ԉ���Ă��܂�(����:" & actual & ")", fullName_
    End If
    LogApiOut "Check_�Ј�ID()"
End Function

Public Function Check_�N(ByVal expected As Integer, ByVal actual As Integer)
    LogApiIn "Check�N()"
    If expected <> actual Then
        AddResult ErrorLog, name_, "���͔N���Ԉ���Ă��܂�(����:" & actual & ")", fullName_
    End If
    LogApiOut "Check_�N()"
End Function

Public Function Check_��(ByVal expected As Integer, ByVal actual As Integer)
    LogApiIn "Check_��()"
    If expected <> actual Then
        AddResult ErrorLog, name_, "���͌����Ԉ���Ă��܂�(����:" & actual & ")", fullName_
    End If
    LogApiOut "Check_��()"
End Function

Public Function Check_�ғ�����(ByVal actual As Integer)
    LogApiIn "Check_�ғ�����()"
    Dim expected As Long
    expected = GetWorkDayCount(year_, month_, isSky_)
    If expected <> actual Then
        AddResult ErrorLog, name_, "�ғ��������Ԉ���Ă��܂�(����:" & actual & ")" & vbCrLf _
        & year_ & "�N" & month_ & "���̉ғ�������" & expected & "���ł��B", fullName_
    End If
    LogApiOut "Check_�ғ�����()"
End Function

' ����̋ΑӃf�[�^
Public Function Check_�J����@(ByRef actual As String)
    LogApiIn "Check_�J����@()"
    If expected <> actual Then
        AddResult WarningLog, name_, "�J����@���ᔽ���Ă��܂��B�x�e���Ԃ��m�F���Ă��������B(����:" & actual & ")" & vbCrLf _
        , fullName_
    End If
    LogApiOut "Check_�J����@()"
End Function