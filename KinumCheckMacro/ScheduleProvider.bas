Attribute VB_Name = "ScheduleProvider"
Option Explicit

' @class �J�����_�[�񋟃N���X
' @breif �J�����_�[����񋟂���
' @note ���x���l����Ȃ�N�����w�肳����ׂ�
' @note ��������݂���΂����񂾂���W�����W���[���ł��悩�����̂���

' @breif Range�擾����Variant�^�z��ɃA�N�Z�X���邽�߂�Enum
Private Enum SceduleRow
    Heading1 = 1
    Heading2
    First
End Enum

Private Enum SceduleColomun
    Pading1 = 1
    Date
    SceduleName
    SceduleType
    Pading2
End Enum

' �V�[�g���
Private Const SceduleSheetName As String = "�J�����_�["
Private Const SceduleRangeName As String = "�X�P�W���[��"

' define
Private daysInfo As Dictionary
Private workDayCount As Dictionary

' @breif �������֐�
Public Function ScheduleProvider_Initialize()
    LogApiIn "ScheduleProvider_Initialize()"
    
    Set daysInfo = New Dictionary
    Set workDayCount = New Dictionary
    
    SetDayInfo
    
    LogApiOut "ScheduleProvider_Initialize()"
End Function

' @breif �I���֐�
Public Function ScheduleProvider_Terminate()
    LogApiIn "ScheduleProvider_Terminate()"

    Set daysInfo = Nothing
    Set workDayCount = Nothing
    
    LogApiOut "ScheduleProvider_Terminate()"
End Function

' @breif �ғ������擾����
' @note �ғ��������v�Z�ł���Όv�Z����B�v�Z�񐔂͗}������
Public Function GetWorkDayCount(year As Integer, month As Integer, isSky As Boolean) As Long
    LogApiIn "GetWorkDayCount()"
    
    Dim key  As Date
    key = DateSerial(year, month, 1)
    
    If workDayCount.Item(key) = "" Then
        GetWorkDayCount = CalculateWorkDayCount(year, month, isSky)
    Else
        GetWorkDayCount = workDayCount.Item(key)
    End If
    
    LogApiOut "GetWorkDayCount()"
End Function

' @brief �o�Γ����v�Z����
Private Function CalculateWorkDayCount(year As Integer, month As Integer, isSky As Boolean) As Long
    LogApiIn "CalculateWorkDayCount()"

    Dim count As Long
    count = 0

    Dim lastDayOfMonth As Integer
    lastDayOfMonth = GetLastDayOfMonth(year, month)
    
    Dim i As Long
    For i = 1 To lastDayOfMonth
        Dim d As Date
        d = DateSerial(year, month, i)
        If IsWorkDay(d, isSky) Then
            count = count + 1
        End If
    Next i
    
    Dim key  As Date
    key = DateSerial(year, month, 1)
    workDayCount(key) = count
    
    CalculateWorkDayCount = count
    
    LogApiOut "CalculateWorkDayCount()"
End Function

' @brief ���̍ŏI�����擾����
Private Function GetLastDayOfMonth(ArgYear As Integer, ArgMonth As Integer) As Integer
    LogApiIn "CalculateWorkDayCount()"

    Dim lastDay As Date
    lastDay = DateSerial(ArgYear, ArgMonth + 1, 0)
    GetLastDayOfMonth = Day(lastDay)
    
    LogApiOut "CalculateWorkDayCount()"
End Function

' @breif �V�[�g����X�P�W���[�������擾����
' @attention Workbook�ɃA�N�Z�X����֐��̂��ߌĂяo���񐔂ɒ���
Private Function SetDayInfo()
    LogApiIn "SetDayInfo()"

    Dim rangeDayInfo As Variant
    rangeDayInfo = Worksheets(SceduleSheetName).Range(SceduleRangeName)
    
    Dim i As Long
    For i = SceduleRow.First To UBound(rangeDayInfo, 1) - 1
        
        Dim KeyDate As Date
        Dim DateType As String
        KeyDate = rangeDayInfo(i, SceduleColomun.Date)
        DateType = rangeDayInfo(i, SceduleColomun.SceduleType)
        daysInfo.Add KeyDate, DateType
        
    Next i
    
    LogApiOut "SetDayInfo()"
End Function

' @breif �o�Γ����ǂ������擾����
' @note �o�O�𖄂ߍ��މ\������ԍ����֐��B���W�b�N������
Public Function IsWorkDay(ArgDate As Date, isSky As Boolean) As Boolean
    LogApiIn "IsWorkDay()"

    Dim dayType As String
    dayType = GetDayType(ArgDate)
        
    IsWorkDay = True
    
    Select Case dayType
    Case "�����̏j��"
        IsWorkDay = False
    Case "Sky���T��"
        If isSky = False Then
            IsWorkDay = False
        End If
    Case Else
        If (Weekday(ArgDate) = vbSunday) Or (Weekday(ArgDate) = vbSaturday) Then
            IsWorkDay = False
        End If
    End Select
    
    
    LogApiOut "IsWorkDay()"
End Function

' @breif �w����t�����擾����
Private Function GetDayType(ArgDate As Date) As String
    LogApiIn "GetDayType()"

    GetDayType = daysInfo.Item(ArgDate)

    LogApiOut "GetDayType()"
End Function



