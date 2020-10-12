Attribute VB_Name = "Logger"
Option Explicit

' @breif ���O�o�̓��x���B�����قǑ������O���o�͂����(���x���������ƃ`�F�b�N���Ԃ��啝�ɑ�����̂Œ���)
#Const LogLevel = 3
#Const IsOutput = True

Private logStorage As String

' �������֐�
Public Function Logger_Initialize()
    logStorage = "" ' Logger_Initialize()�O�ɌĂ΂ꂽ���O�������Ă��܂����ߍŏ��̃��O�o�͂��O�ɃN���A����
    LogApiIn "Logger_Initialize()"
    
    LogApiOut "Logger_Initialize()"
End Function

' �I���֐�
Public Function Logger_Terminate()
    LogApiIn "Logger_Terminate()"

    ' �I���̌_�@�ł܂Ƃ߂ă��O�������o��
    WriteLogFile
    logStorage = ""

    LogApiOut "Logger_Terminate()"
End Function

' @breif ���O�t�@�C���ɏo�͂������o��
Private Function WriteLogFile()
    LogApiIn "OutputLogFile()"
    
    If IsOutputDebugLogFile() = False Then
        Exit Function
    End If

    Dim fileNumber
    fileNumber = FreeFile()
    
    On Error Resume Next
    Open GenerateFullName(GetDirOutputDebugLog, GetNameDebugLogFile) For Output As #fileNumber
    If Err.Number <> 0 Then
        LogError "Cannot open log file(" & GenerateFullName(GetDirOutputDebugLog, GetNameDebugLogFile) & ")! " _
        & "ErrNo:" & Err.Number & "ErrDescription:" & Err.Description & "ErrFunction:OutputLogFile()"
    End If
    Print #fileNumber, logStorage
    Close #fileNumber

    LogApiOut "OutputLogFile()"
End Function


' @note �ǂ����Ō����悤�ȃ��O�̎�ނ�����
Public Function LogError(output As String)
#If LogLevel >= 1 Then
    Log "[ERROR] " & output
#End If
End Function

Public Function LogWarning(output As String)
#If LogLevel >= 2 Then
    Log "[WARNING] " & output
#End If
End Function

Public Function LogInfo(output As String)
#If LogLevel >= 3 Then
    Log "[INFO] " & output
#End If
End Function

Public Function LogDebug(output As String)
#If LogLevel >= 4 Then
    Log "[DEBUG] " & output
#End If
End Function

Public Function LogApiIn(output As String)
#If LogLevel >= 5 Then
    Log "[API_IN] " & output
#End If
End Function

Public Function LogApiOut(output As String)
#If LogLevel >= 5 Then
    Log "[API_OUT] " & output
#End If
End Function

Private Function Log(output As String)
    ' �C�~�f�B�G�C�g�E�B���h�E�ɏo��
    Dim current As String
    current = GetDateTimer()
    Debug.Print current & " " & output
    StockLog current & output
End Function

' @breif ���O������ϐ��ɕۑ�����
' @note VBA�ɂ�����String�̍ő吔�͖�20�������炵���̂Ō����ӂ�͍l�����Ȃ�
Private Function StockLog(Content As String)
    If IsOutputDebugLogFile() Then
        logStorage = logStorage & Content & vbCrLf
    End If
End Function

' https://vbabeginner.net/vba�Ō��ݓ������~���b�P�ʂŎ擾����/
Private Function GetDateTimer() As String

    Dim t       '// Timer�l
    Dim tint    '// Timer�l�̐�������
    Dim m       '// �~���b
    Dim ret     '// �߂�l
    Dim sHour
    Dim sMinute
    Dim sSecond
    
    '// Timer�l���擾
    t = Timer
    
    '// Timer�l�̐����������擾
    tint = Int(t)
    
    '// �����b���擾
    sHour = Int(tint / (60 * 60))
    sMinute = Int((tint - (sHour * 60 * 60)) / 60)
    sSecond = tint - (sHour * 60 * 60 + sMinute * 60)
    
    '// Timer�l�̏����������擾
    m = t - tint
    
    '// hh:mm:ss.fff�ɐ��`
    ret = format(sHour, "00")
    ret = ret & ":"
    ret = ret & format(sMinute, "00")
    ret = ret & ":"
    ret = ret & format(sSecond, "00")
    ret = ret & format(Left(Right(CStr(m), Len(m) - 1), 4), ".000")
    
    GetDateTimer = ret
    
End Function
