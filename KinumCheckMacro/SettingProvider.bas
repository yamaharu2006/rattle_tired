Attribute VB_Name = "SettingProvider"
' @breif �ݒ����񋟂��郉�C�u����
'@note ���Ƃ���ǉ��������񂾂���N���X�Ɋ�Â��ݒ�l�͂��̃N���X����Q�Ƃ��Ă��܂��Ă���
Option Explicit

' ReadOnly
Private TargetPath As String
Private BackupPath As String

Private isOutputFileCheckedResult_ As Boolean
Private dirCheckedResult_ As String
Private fileCheckedResult_ As String

Private �莞�o�Ύ��� As Date
Private �莞�ދΎ��� As Date
Private ���x�e���� As Date
Private �莞��x�e���� As Date
Private �莞�ގГ� As String

Private debugLogLevel_ As Long
Private isOutputDebugLogFile_ As Boolean
Private dirOutputDebugLog_ As String
Private nameDebugLogFile_ As String

Public Function SettingProvider_Initialize()
    LogApiIn "SettingProvider_Initialize()"
    
    ' �I�u�W�F�N�g�����Ƃ��͂��΂��΃{�g���l�b�N�ƂȂ�B�{���̓V�[�g�ǂݍ��݂��ꌳ�Ǘ��������B
    With ThisWorkbook.Worksheets("�`�F�b�N")
    
        TargetPath = .Range("�`�F�b�N�Ώۃt�H���_")
        BackupPath = .Range("�o�b�N�A�b�v��")
        
        isOutputFileCheckedResult_ = .Range("IsOutputFile")
        dirCheckedResult_ = .Range("DirCheckedResult")
        fileCheckedResult_ = .Range("FileCheckedResult")
    
        �莞�o�Ύ��� = .Range("�莞�o�Ύ���")
        �莞�ދΎ��� = .Range("�莞�ދΎ���")
        ���x�e���� = .Range("���x�e����")
        �莞��x�e���� = .Range("�莞��x�e����")
        �莞�ގГ� = .Range("�莞�ގГ�")
        
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

Public Function Get�莞�o�Ύ���() As Date
    Get�莞�o�Ύ��� = �莞�o�Ύ���
End Function

Public Function Get�莞�ދΎ���() As Date
    Get�莞�ދΎ��� = �莞�ދΎ���
End Function

Public Function Get���x�e����() As Date
    Get���x�e���� = ���x�e����
End Function

Public Function Get�莞��x�e����() As Date
    Get�莞��x�e���� = �莞��x�e����
End Function

Public Function Get�莞�ގГ�() As String
    Get�莞�ގГ� = �莞�ގГ�
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
