Attribute VB_Name = "Logger"
Option Explicit

' @breif ���O�o�̓��x���B�����قǑ������O���o�͂����
#Const LogLevel = 5

' @note �ǂ����Ō����悤�ȃ��O�̎�ނ�����
Public Function LogError(output As String)
#If LogLevel >= 1 Then
    Debug.Print GetDateTimer() & " [ERROR] " & output
#End If
End Function

Public Function LogWarning(output As String)
#If LogLevel >= 2 Then
    Debug.Print GetDateTimer() & " [WARNING] " & output
#End If
End Function

Public Function LogInfo(output As String)
#If LogLevel >= 3 Then
    Debug.Print GetDateTimer() & " [INFO] " & output
#End If
End Function

Public Function LogDebug(output As String)
#If LogLevel >= 4 Then
    Debug.Print GetDateTimer() & " [DEBUG] " & output
#End If
End Function

Public Function LogApiIn(output As String)
#If LogLevel >= 5 Then
    Debug.Print GetDateTimer() & " [API_IN] " & output
#End If
End Function

Public Function LogApiOut(output As String)
#If LogLevel >= 5 Then
    Debug.Print GetDateTimer() & " [API_OUT] " & output
#End If
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
    ret = Format(sHour, "00")
    ret = ret & ":"
    ret = ret & Format(sMinute, "00")
    ret = ret & ":"
    ret = ret & Format(sSecond, "00")
    ret = ret & Format(Left(Right(CStr(m), Len(m) - 1), 4), ".000")
    
    GetDateTimer = ret
End Function
