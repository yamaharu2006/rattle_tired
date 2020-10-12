Attribute VB_Name = "Logger"
Option Explicit

' @breif ログ出力レベル。高いほど多くログが出力される(レベルが高いとチェック時間が大幅に増えるので注意)
#Const LogLevel = 3
#Const IsOutput = True

Private logStorage As String

' 初期化関数
Public Function Logger_Initialize()
    logStorage = "" ' Logger_Initialize()前に呼ばれたログを消してしまうため最初のログ出力より前にクリアする
    LogApiIn "Logger_Initialize()"
    
    LogApiOut "Logger_Initialize()"
End Function

' 終了関数
Public Function Logger_Terminate()
    LogApiIn "Logger_Terminate()"

    ' 終了の契機でまとめてログを書き出す
    WriteLogFile
    logStorage = ""

    LogApiOut "Logger_Terminate()"
End Function

' @breif ログファイルに出力を書き出す
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


' @note どこかで見たようなログの種類だこと
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
    ' イミディエイトウィンドウに出力
    Dim current As String
    current = GetDateTimer()
    Debug.Print current & " " & output
    StockLog current & output
End Function

' @breif ログを内部変数に保存する
' @note VBAにおけるStringの最大数は約20億文字らしいので桁あふれは考慮しない
Private Function StockLog(Content As String)
    If IsOutputDebugLogFile() Then
        logStorage = logStorage & Content & vbCrLf
    End If
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
    ret = format(sHour, "00")
    ret = ret & ":"
    ret = ret & format(sMinute, "00")
    ret = ret & ":"
    ret = ret & format(sSecond, "00")
    ret = ret & format(Left(Right(CStr(m), Len(m) - 1), 4), ".000")
    
    GetDateTimer = ret
    
End Function
