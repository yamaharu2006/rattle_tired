VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SkyWorkScheduleBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @breif Sky�Ј������Ζ��\�f�[�^
Option Explicit

Private Enum �o�Ώ��_����
    ���� = 1
    �j����
    �o�Έ��
    �Ζ�����_�o�З�
    �Ζ�����_�ގЗ�
    �Αӗ�1
    �Αӗ�2
    ���l��
    �󔒗�1
    �莞����_���x�e��
    �莞����_�[�x�e��
    ������
    �莞��x�e��
    ���ԊO�Ζ�_�J�n��
    ���ԊO�Ζ�_�I����
    ���ԊO�Ζ�_���x�e��
    ���ԊO�Ζ�_�[�x�e��
    ���ԊO�Ζ�_���{��
    ���ԊO�Ζ�_���ʗ�
    ���ԊO�Ζ�_�x����
    ���ԊO�Ζ�_�[���
    ���Ζ����ԗ�
    �@�蒴�ߗ�
    ���������
End Enum

Const Range�N As String = "D1"
Const Range�� As String = "B2"
Const Range�ғ����� As String = "D3"
Const Range�`�[�� As String = "G2"
Const Range��E As String = "O1"
Const Range�O���[�v As String = "O2"
Const Range�Ј�ID As String = "U1"
Const Range���� As String = "U2"

Const Range�Ζ��f�[�^ As String = "A7:X68"

Private WorkbookName_ As String

' [�����ے�o]�V�[�g
'Const SheetFiling As String = "�����ے�o"
Private �N_ As Long
Private ��_ As Long
Private �ғ�����_ As Long
Private �`�[��_ As String
Private ��E_ As String
Private �O���[�v_ As String
Private ����_ As String
Private �Ј�ID_ As String
Private ����_ As String

Private Const DaysDataArrayMax As Long = 32 ' �Ζ��f�[�^�z��̍ő�v�f��(31+1)
Private �Ζ��f�[�^_(DaysDataArrayMax) As WorkScheduleDayData

' Setter/Getter
Property Let WorkbookName(Name As String)
    If Name <> "" Then
        WorkbookName_ = Name
    End If
End Property

Property Let Year(arg As Variant)
    If Not IsEmpty(arg) Then
        �N_ = arg
    End If
End Property

Property Let Month(arg As Variant)
    If Not IsEmpty(arg) Then
        ��_ = arg
    End If
End Property

Property Let �ғ�����(arg As Long)
    If Not IsEmpty(arg) Then
        �ғ�����_ = arg
    End If
End Property

Property Get �ғ�����() As Long
    �ғ����� = �ғ�����_
End Property

Property Let �`�[��(arg As String)
    �`�[��_ = arg
End Property

Property Let ��E(arg As String)
    ��E_ = arg
End Property

Property Let �O���[�v(arg As String)
    �O���[�v_ = arg
End Property

Property Let ����(arg As String)
    ����_ = arg
End Property

Property Let �Ј�ID(arg As String)
    �Ј�ID_ = arg
End Property

Property Let ����(arg As String)
    ����_ = arg
End Property

Property Get �Ζ��f�[�^(i As Long) As WorkScheduleDayData
    Set �Ζ��f�[�^ = �Ζ��f�[�^_(i)
End Property

' @breif �R���X�g���N�^
Private Sub Class_Initialize()
    LogApiIn "SkyWorkScheduleBook:Class_Initialize()"
    
    Dim i As Long
    For i = 0 To UBound(�Ζ��f�[�^_)
        Set �Ζ��f�[�^_(i) = New WorkScheduleDayData
    Next i
    
    LogApiOut "SkyWorkScheduleBook:Class_Initialize()"
End Sub

' @breif �f�X�g���N�^
Private Sub Class_Terminate()
    LogApiIn "SkyChecker:Class_Terminate()"
    
    Dim i As Long
    For i = 0 To UBound(�Ζ��f�[�^_)
        Set �Ζ��f�[�^_(i) = Nothing
    Next i

    LogApiOut "SkyChecker:Class_Terminate()"
End Sub

' @breif �f�[�^���C���|�[�g����
Public Function ImportData()
    LogApiIn "KinmuDataSky:ImportData()"
    
    With Workbooks(WorkbookName_).Sheets("�����ے�o")
        �N_ = .Range(Range�N)
        ��_ = .Range(Range��)
        �ғ�����_ = .Range(Range�ғ�����)
        �`�[��_ = .Range(Range�`�[��)
        ��E_ = .Range(Range��E)
        �O���[�v_ = .Range(Range�O���[�v)
        �Ј�ID_ = .Range(Range�Ј�ID)
        ����_ = .Range(Range����)
        
        Dim var As Variant
        var = .Range(Range�Ζ��f�[�^)
        Dim i As Long
        For i = 1 To UBound(var) Step 2
            �Ζ��f�[�^_(i / 2).SetDate �N_, ��_, Int(i / 2 + 1)
            �Ζ��f�[�^_(i / 2).�o�� = var(i + 1, �Ζ�����_�o�З�)
            �Ζ��f�[�^_(i / 2).�ގ� = var(i + 1, �Ζ�����_�ގЗ�)
            Dim �Α�() As Variant
            �Α� = Array(var(i, �Αӗ�1), var(i, �Αӗ�2), var(i + 1, �Αӗ�1), var(i + 1, �Αӗ�2))
            �Ζ��f�[�^_(i / 2).�Α� = �Α�
            �Ζ��f�[�^_(i / 2).���l = var(i, ���l��)
            �Ζ��f�[�^_(i / 2).���x�e = var(i + 1, �莞����_���x�e��)
            �Ζ��f�[�^_(i / 2).�[�x�e = var(i + 1, �莞����_�[�x�e��)
            �Ζ��f�[�^_(i / 2).���ԊO�Ζ����c��_�J�n = var(i + 1, ���ԊO�Ζ�_�J�n��)
            �Ζ��f�[�^_(i / 2).���ԊO�Ζ����c��_�I�� = var(i + 1, ���ԊO�Ζ�_�I����)
            �Ζ��f�[�^_(i / 2).���ԊO�Ζ����c��_���x�e = var(i + 1, ���ԊO�Ζ�_���x�e��)
            �Ζ��f�[�^_(i / 2).���ԊO�Ζ�_�J�n = var(i + 1, ���ԊO�Ζ�_�J�n��)
            �Ζ��f�[�^_(i / 2).���ԊO�Ζ�_���x�e = var(i + 1, ���ԊO�Ζ�_���x�e��)
            �Ζ��f�[�^_(i / 2).���ԊO�Ζ�_�[�x�e = var(i + 1, ���ԊO�Ζ�_�I����)
            �Ζ��f�[�^_(i / 2).���Ζ����� = var(i + 1, ���Ζ����ԗ�)
        Next i
    End With
    
    LogApiOut "KinmuDataSky:ImportData()"
End Function
