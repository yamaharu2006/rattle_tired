Attribute VB_Name = "Logger"
Option Explicit

' @breif ログ出力レベル。高いほど多くログが出力される
#Const LogLevel = 5

' @note どこかで見たようなログの種類だこと
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


' https://vbabeginner.net/vbaで現在日時をミリ秒単位で取得する/
Private Function GetDateTimer() As String
    Dim t       '// Timer値
    Dim tint    '// Timer値の整数部分
    Dim m       '// ミリ秒
    Dim ret     '// 戻り値
    Dim sHour
    Dim sMinute
    Dim sSecond
    
    '// Timer値を取得
    t = Timer
    
    '// Timer値の整数部分を取得
    tint = Int(t)
    
    '// 時分秒を取得
    sHour = Int(tint / (60 * 60))
    sMinute = Int((tint - (sHour * 60 * 60)) / 60)
    sSecond = tint - (sHour * 60 * 60 + sMinute * 60)
    
    '// Timer値の小数部分を取得
    m = t - tint
    
    '// hh:mm:ss.fffに整形
    ret = Format(sHour, "00")
    ret = ret & ":"
    ret = ret & Format(sMinute, "00")
    ret = ret & ":"
    ret = ret & Format(sSecond, "00")
    ret = ret & Format(Left(Right(CStr(m), Len(m) - 1), 4), ".000")
    
    GetDateTimer = ret
End Function
