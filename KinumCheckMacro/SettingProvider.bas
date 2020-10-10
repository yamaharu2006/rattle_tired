Attribute VB_Name = "SettingProvider"
' @breif 設定情報を提供するライブラリ
'@note あとから追加したもんだからクラスに基づく設定値はそのクラスから参照してしまっている
Option Explicit

' ReadOnly
Private TargetPath As String
Private BackupPath As String

Private isOutputFileCheckedResult_ As Boolean
Private dirCheckedResult_ As String
Private fileCheckedResult_ As String

Private 定時出勤時間 As Date
Private 定時退勤時間 As Date
Private 昼休憩時間 As Date
Private 定時後休憩時間 As Date
Private 定時退社日 As String

Private debugLogLevel_ As Long
Private isOutputDebugLogFile_ As Boolean
Private dirOutputDebugLog_ As String
Private nameDebugLogFile_ As String

Public Function SettingProvider_Initialize()
    LogApiIn "SettingProvider_Initialize()"
    
    ' オブジェクトを作るときはしばしばボトルネックとなる。本当はシート読み込みを一元管理したい。
    With ThisWorkbook.Worksheets("チェック")
    
        TargetPath = .Range("チェック対象フォルダ")
        BackupPath = .Range("バックアップ先")
        
        isOutputFileCheckedResult_ = .Range("IsOutputFile")
        dirCheckedResult_ = .Range("DirCheckedResult")
        fileCheckedResult_ = .Range("FileCheckedResult")
    
        定時出勤時間 = .Range("定時出勤時間")
        定時退勤時間 = .Range("定時退勤時間")
        昼休憩時間 = .Range("昼休憩時間")
        定時後休憩時間 = .Range("定時後休憩時間")
        定時退社日 = .Range("定時退社日")
        
        debugLogLevel_ = .Range("DebugLogLevel")
        isOutputDebugLogFile_ = .Range("IsOutputDebugLogFile")
        dirOutputDebugLog_ = .Range("DirOutputDebugLog")
        nameDebugLogFile_ = .Range("NameDebugLogFile")
    
    End With
    
    LogApiOut "SettingProvider_Initialize()"
End Function

Public Function SettingProvider_Terminate()
    LogApiIn "SettingProvider_Terminate()"
    
    LogApiOut "SettingProvider_Terminate()"
End Function

Public Function GetTargetPath() As String
    GetTargetPath = TargetPath
End Function

Public Function GetBackupPath() As String
    GetBackupPath = BackupPath
End Function

Public Function Get定時出勤時間() As Date
    Get定時出勤時間 = 定時出勤時間
End Function

Public Function Get定時退勤時間() As Date
    Get定時退勤時間 = 定時退勤時間
End Function

Public Function Get昼休憩時間() As Date
    Get昼休憩時間 = 昼休憩時間
End Function

Public Function Get定時後休憩時間() As Date
    Get定時後休憩時間 = 定時後休憩時間
End Function

Public Function Get定時退社日() As String
    Get定時退社日 = 定時退社日
End Function

Public Function GetDebugLogLevel() As String
    GetDebugLogLevel = debugLogLevel
End Function

Public Function IsOutputDebugLogFile() As String
    IsOutputDebugLogFile = isOutputDebugLogFile_
End Function

Public Function GetDirOutputDebugLog() As String
    GetDirOutputDebugLog = dirOutputDebugLog_
End Function

Public Function GetNameDebugLogFile() As String
    GetNameDebugLogFile = nameDebugLogFile_
End Function

Public Function IsOutputFileCheckedResult()
    IsOutputFileCheckedResult = isOutputFileCheckedResult_
End Function

Public Function GetDirCheckedResult()
    GetDirCheckedResult = dirCheckedResult_
End Function

Public Function GetFileCheckedResult()
    GetFileCheckedResult = fileCheckedResult_
End Function
