Attribute VB_Name = "SettingProvider"
' @breif 設定情報を提供するライブラリ
'@note あとから追加したもんだからクラスに基づく設定値はそのクラスから参照してしまっている
Option Explicit

' ReadOnly
Private TargetPath As String
Private BackupPath As String

Private 定時出勤時間 As Date
Private 定時退勤時間 As Date
Private 昼休憩時間 As Date
Private 定時後休憩時間 As Date
Private 定時退社日 As String

Private Function SettingProvider_Initialize()
    LogApiIn "SettingProvider_Initialize()"
    
    ' オブジェクトを作るときはしばしばボトルネックとなる。本当はシート読み込みを一元管理したい。
    With ThisWorkbook.Worksheets("チェック")
    
        TargetPath = .Range("チェック対象フォルダ")
        BackupPath = .Range("バックアップ先")
    
        定時出勤時間 = .Range("定時出勤時間")
        定時退勤時間 = .Range("定時退勤時間")
        昼休憩時間 = .Range("昼休憩時間")
        定時後休憩時間 = .Range("定時後休憩時間")
        定時退社日 = .Range("定時退社日")
    
    End With
    
    LogApiOut "SettingProvider_Initialize()"
End Function

Private Function SettingProvider_Terminate()
    LogApiIn "SettingProvider_Terminate()"
    
    LogApiOut "SettingProvider_Terminate()"
End Function

Public Function GetTargetPath()
    GetTargetPath = TargetPath
End Function

Public Function GetBackupPath()
    GetBackupPath = BackupPath
End Function

Public Function Get定時出勤時間()
    Get定時出勤時間 = 定時出勤時間
End Function

Public Function Get定時退勤時間()
    Get定時退勤時間 = 定時退勤時間
End Function

Public Function Get昼休憩時間()
    Get昼休憩時間 = 昼休憩時間
End Function

Public Function Get定時後休憩時間()
    Get定時後休憩時間 = 定時後休憩時間
End Function

Public Function Get定時退社日()
    Get定時退社日 = 定時退社日
End Function
