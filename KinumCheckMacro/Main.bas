Attribute VB_Name = "Main"
Option Explicit

' @breif �����o�ꗗ_Sky�̗���
Private Enum ColMemberSky
    IsChecked = 1
    EmplyoeeId
    MemberName
    Post
    MainJob
    SubJob
    Empty1
    Empty2
    Empty3
    Prohibited
    FileName
End Enum

' @breif �����o�ꗗ_Sky�̍s���
Private Enum RowMemberSky
    heading = 1
    First
End Enum

' @breif �����o�ꗗ_BP�̗���
Private Enum ColMemberPartner
    IsChecked = 1
    EmplyoeeId
    MemberName
    Corporate
    MainJob
    SubJob
    Empty1
    Empty2
    Empty3
    Prohibited
    FileName
End Enum

' @breif �����o�ꗗ_BP�̍s���
Private Enum RowMemberBP
    heading = 1
    First
End Enum

Const FooterSizeMemberBP As Integer = 1

Public Sub ButtonCheckAll()
    SetUp
    LogInfo "Pressed Button(CheckAll)"
    CheckSky
    CheckPartner
    TearDown
End Sub

Public Sub ButtonCheckSky()
    SetUp
    LogInfo "Pressed Button(CheckSky)"
    CheckSky
    TearDown
End Sub

Public Sub ButtonCheckPartner()
    SetUp
    LogInfo "Pressed Button(Partner)"
    CheckPartner
    TearDown
End Sub

' @breif Sky�̋Ζ��\�����o���`�F�b�N����
Private Function CheckSky()
    LogApiIn "CheckSky()"
    LogInfo "Start Check(Sky)"
    
    ' �Ζ��\�`�F�b�N
    Dim memberList As Variant
    memberList = Worksheets("�`�F�b�N").Range("�����o�ꗗ_Sky")
    
    ' �`�F�b�N�Ώ�(IsChecked = True)�Ȃ�Ζ��\���`�F�b�N����
    Const FooterSizeMemberSky As Integer = 1
    Dim i As Long
    For i = RowMemberSky.First To UBound(memberList, 1) - FooterSizeMemberSky
        If Not IsEmpty(memberList(i, ColMemberSky.IsChecked)) And memberList(i, ColMemberSky.IsChecked) Then
            Dim checker As SkyChecker
            Set checker = New SkyChecker
            
            checker.BackupPath = GetBackupPath
            checker.Path = GetTargetPath
            checker.FileName = memberList(i, ColMemberSky.FileName)
            checker.EmploeeName = memberList(i, ColMemberSky.MemberName)
            checker.Year = GetTargetYear
            checker.Month = GetTargetMonth
            checker.EmploeeId = memberList(i, ColMemberSky.EmplyoeeId)
            checker.Post = memberList(i, ColMemberSky.Post)
            
            checker.Check
        End If
    Next i

    ' ���O�o��
    OutputResult
    
    LogInfo "End Check(Sky)"
    LogApiOut "CheckSky()"
End Function


' @breif BP�̋Ζ��\���`�F�b�N����
Private Function CheckPartner()
    LogApiIn "CheckPartner()"
    LogInfo "Start Check(Partner)"
    
    ' �Ζ��\�`�F�b�N
    Dim memberList As Variant
    memberList = Worksheets("�`�F�b�N").Range("�����o�ꗗ_BP")
    
    ' �`�F�b�N�Ώ�(IsChecked = True)�Ȃ�Ζ��\���`�F�b�N����
    Dim i As Long
    For i = RowMemberBP.First To UBound(memberList, 1) - FooterSizeMemberBP
        If Not IsEmpty(memberList(i, ColMemberBP.IsChecked)) And memberList(i, ColMemberBP.IsChecked) Then
'            Dim checker As PartnerChecker
'            Set checker = New PartnerChecker
'
'            Dim FileName As String
'            Dim Name As String
'            Dim position As String
'            Dim employeeId As String
'            FileName = memberList(i, ColMemberSky.FileName)
'            Name = memberList(i, ColMemberSky.MemberName)
'            position = memberList(i, ColMemberSky.Post)
'            employeeId = memberList(i, ColMemberSky.EmplyoeeId)
'            checker.Initialize GetTargetPath, GetBackupPath, FileName, Name, position, employeeId
'
'            checker.Check
        End If
    Next i

    
    LogInfo "End Check(Partner)"
    LogApiIn "CheckPartner()"
End Function

' @breif �}�N���J�n����
Private Function SetUp()
    Logger_Initialize
    LogApiIn "SetUp()"

    ' VBA������ Setup
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    ' �e�평���������Ăяo��
    CheckedResult_Initialize
    ScheduleProvider_Initialize
    SettingProvider_Initialize
    
    LogApiOut "SetUp()"
End Function

Private Function TearDown()
    LogApiIn "TearDown()"

    ' �e��I�������Ăяo��
    CheckedResult_Terminate
    ScheduleProvider_Terminate
    SettingProvider_Terminate

    ' VBA������ Teardown
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    LogApiOut "TearDown()"
    Logger_Terminate
End Function
