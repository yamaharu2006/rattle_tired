VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SkyChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Password As String = "pass"

Private year_ As Long
Private month_ As Long
Private path_ As String
Private backupPath_ As String
Private fileName_ As String
Private fullName_ As String
Private name_ As String
Private post_ As String
Private employeeId_ As String

Private commonChecker As commonChecker
Private data_ As SkyWorkScheduleBook

' Setter/Getter
Property Let Year(ByVal argYear As Long)
    year_ = argYear
    commonChecker.Year = argYear
    data_.Year = argYear
End Property

Property Let Month(ByVal argMonth As Long)
    month_ = argMonth
    commonChecker.Month = argMonth
    data_.Month = argMonth
End Property

Property Let Name(ByRef argName As String)
    name_ = argName
    commonChecker.Name = argName
    data_.���� = argName
End Property

Property Let FullName(ByRef argFullName As String)
    fullName_ = argFullName
    commonChecker.FullName = argFullName
End Property

' @breif �R���X�g���N�^
Private Sub Class_Initialize()
    LogApiIn "SkyChecker:Class_Initialize()"
    
    Set data_ = New SkyWorkScheduleBook
    Set commonChecker = New commonChecker
    
    LogApiOut "SkyChecker:Class_Initialize()"
End Sub

' @breif �f�X�g���N�^
Private Sub Class_Terminate()
    LogApiIn "SkyChecker:Class_Terminate()"
    
    Set data_ = Nothing
    Set commonChecker = Nothing

    LogApiOut "SkyChecker:Class_Terminate()"
End Sub

' @breif �C���X�^���X������
' @note �����t���R���X�g���N�^�̑���
Public Function Initialize(Path As String, BackupPath As String, fileName As String, Name As String, post As String, employeeId As String)

    'year_ = Year
    'month_ = Month
    path_ = Path
    backupPath_ = BackupPath
    fileName_ = fileName
    fullName_ = GenerateFullName(Path, fileName)
    name_ = Name
    post_ = post
    employeeId_ = employeeId
    
    data_.WorkbookName = fileName_

End Function

' @breif �Ζ��\���`�F�b�N����
Public Function Check()
    LogApiIn "Check()"
    
    ' �t�@�C�����Ȃ���΃`�F�b�N���Ȃ�
    If Not ExistsFile(fullName_) Then
        AddResult ErrorLog, name_, "�t�@�C�����J���܂���ł����B", fullName_
        LogApiOut "Check()"
        Exit Function
    End If
    
    SetUpChecking
    
    ' �Ζ��\�`�F�b�N
    With commonChecker
        .Check_�ғ����� data_.�ғ�����
    End With
    
    Check_sample
    Check_DaysData
    
    TearDownChecking
    
    LogApiOut "Check()"
End Function

' @breif �`�F�b�N�̏������s��
' @note ��ɋΖ��\�̃f�[�^���C���|�[�g���邱�ƂŃV�[�g�ɃA�N�Z�X����񐔂��ւ炷
Private Function SetUpChecking()
    LogApiIn "SetUpChecking()"
    
    ' �t�@�C�����o�b�N�A�b�v����
    BackupFile path_, fileName_, backupPath_, fileName_
    
    ' �t�@�C�����J��
    OpenWorkbook backupPath_, fileName_
    
    ' �`�F�b�N�ɕK�v�ȃf�[�^����荞��
    data_.ImportData
    
    ' �t�@�C�������
    CloseWorkbook (fileName_)
    
    ' �`�F�b�N���ʂɍŏI�X�V�������o��
    Dim dateLastModified As Date
    GetDateLastModified fullName_, dateLastModified
    AddResult InfoLog, name_, "�ŏI�X�V����(" & format(dateLastModified, "yyyy/mm/dd hh:nn") & ")�̃t�@�C�����`�F�b�N���܂����B", fullName_

    LogApiOut "SetUpChecking()"
End Function

' @breif �`�F�b�N�̌㏈�����s��
Private Function TearDownChecking()
    LogApiIn "TearDownChecking()"
    
    ' �G���[�����ƌx���������o���B���[�v���Ԃ��G�O���̂ŁA�o�͗L����Switch������
    Dim countError As Long
    Dim countWarning As Long
    countError = GetCountReuslt(ErrorLog, name_)
    countWarning = GetCountReuslt(WarningLog, name_)
    AddResult InfoLog, name_, "�t�@�C�����`�F�b�N���܂����B(�G���[:" & countError & "��, �x��:" & countWarning & "��)", fullName_
    
    LogApiOut "TearDownChecking()"
End Function

'////////// �Ζ��\�`�F�b�N�֐� //////////
' @breif ����P�ʂŃ`�F�b�N����
' @note Let/Set ���g����For Each�ł��Ȃ��̂�... ʧ
Private Function Check_DaysData()
    LogApiIn "Check_sample()"
    
    Dim i As Long
    Dim dayData As WorkScheduleDayData
    
    ' ������ɑ΂��ă`�F�b�N����
    Const DaysDataArrayMax As Long = 32
    For i = 0 To DaysDataArrayMax
        Set dayData = data_.�Ζ��f�[�^(i)
        
        ' >>>>> �Ζ��\�`�F�b�N(����P��)��������
        Check_���l���ʋΔ�L�� dayData.���l
        
        ' <<<<< �Ζ��\�`�F�b�N(����P��)�����܂�
        
    Next i
    
    LogApiOut "Check_sample()"
End Function

Private Function Check_sample()
    LogApiIn "Check_sample()"

    AddResult ErrorLog, name_, "��������̃G���[���������܂����B", fullName_

    LogApiOut "Check_sample()"
End Function


' @note ����ȑf�ȃ`�F�b�N
Private Function Check_���l���ʋΔ�L��(actual As String)
    LogApiIn "Check_���l���ʋΔ�L��()"
    
    Select Case True
    Case actual Like "�y�ʋΔ�:�����z"
    Case actual Like "�y�ʋΔ�:���̂݁z"
    Case actual Like "�y�ʋΔ�:���̂݁z"
    Case actual Like "�y�ʋΔ�:�Ȃ��z"
    Case actual Like "�y�ʋΔ�:�Ȃ��z"
    Case actual Like "�y�ʋΔ�:�Ȃ��z�y�e�����[�N�z"
    Case Else
        AddResult ErrorLog, name_, "�ȉ��̂����ꂩ���Ζ��\�̔��l���ɋL�ڂ��Ă�������", fullName_
    End Select
    
    LogApiOut "Check_���l���ʋΔ�L��()"
End Function
