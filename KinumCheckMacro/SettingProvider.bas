Attribute VB_Name = "SettingProvider"
' @breif �ݒ����񋟂��郉�C�u����
'@note ���Ƃ���ǉ��������񂾂���N���X�Ɋ�Â��ݒ�l�͂��̃N���X����Q�Ƃ��Ă��܂��Ă���
Option Explicit

' ReadOnly
Private TargetPath As String
Private BackupPath As String

Private �莞�o�Ύ��� As Date
Private �莞�ދΎ��� As Date
Private ���x�e���� As Date
Private �莞��x�e���� As Date
Private �莞�ގГ� As String

Private Function SettingProvider_Initialize()
    LogApiIn "SettingProvider_Initialize()"
    
    ' �I�u�W�F�N�g�����Ƃ��͂��΂��΃{�g���l�b�N�ƂȂ�B�{���̓V�[�g�ǂݍ��݂��ꌳ�Ǘ��������B
    With ThisWorkbook.Worksheets("�`�F�b�N")
    
        TargetPath = .Range("�`�F�b�N�Ώۃt�H���_")
        BackupPath = .Range("�o�b�N�A�b�v��")
    
        �莞�o�Ύ��� = .Range("�莞�o�Ύ���")
        �莞�ދΎ��� = .Range("�莞�ދΎ���")
        ���x�e���� = .Range("���x�e����")
        �莞��x�e���� = .Range("�莞��x�e����")
        �莞�ގГ� = .Range("�莞�ގГ�")
    
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

Public Function Get�莞�o�Ύ���()
    Get�莞�o�Ύ��� = �莞�o�Ύ���
End Function

Public Function Get�莞�ދΎ���()
    Get�莞�ދΎ��� = �莞�ދΎ���
End Function

Public Function Get���x�e����()
    Get���x�e���� = ���x�e����
End Function

Public Function Get�莞��x�e����()
    Get�莞��x�e���� = �莞��x�e����
End Function

Public Function Get�莞�ގГ�()
    Get�莞�ގГ� = �莞�ގГ�
End Function
