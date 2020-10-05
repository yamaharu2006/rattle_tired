Attribute VB_Name = "Main"
Option Explicit

Sub CheckAll()
    LogApiIn "CheckAll()"

    
    
    LogApiOut "CheckAll()"
End Sub

Sub CheckSky()
    LogApiIn "CheckSky()"
    
    ' ���O����
    SetUp
    
    ' �Ζ��\�`�F�b�N
    Dim checker As SkyChecker
    Set checker = New SkyChecker
    checker.Check

    GetWorkDayCount 2020, 9, True

    ' ���O�o��
    OutputResult

    ' ���㏈��
    TearDown
    
    LogApiOut "CheckSky()"
End Sub

Sub CheckPartner()
    LogApiIn "CheckPartner()"
    
    StartUp

    TearDown
    
    LogApiIn "CheckPartner()"
End Sub

' @note ��x�ǂ�ł����Q�͂Ȃ�(�Ǝv��)�̂Ńu���b�N�����͍��Ȃ�
Private Function SetUp()
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
End Function
