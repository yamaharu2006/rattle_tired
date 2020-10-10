Attribute VB_Name = "CheckedResultProvider"
' @breif �`�F�b�N���ʂ��Ǘ�����N���X
' @note �`�F�b�N���ʒ񋟕��ƕۊǕ��ŕ����Ă������C�͂���
Option Explicit

' @breif Enum���g�����ق������\�I�ɂ͑����񂾂��ǁA���X�ύX����̂��ʓ|�Ŏg��Ȃ�����
Public Enum resultType
    ErrorLog
    WarningLog
    InfoLog
End Enum

Public Enum ColumnOutputRange
    ColType
    ColTarget
    ColContent
    ColFullPath
    ColCount
End Enum

Private Const OutputWorksheet As String = "�`�F�b�N����"
Private Const OutputCell As String = "E3"
Private Const OutputStartingPoint As String = "C7"

Private CheckedResultList As Collection
Private IsOutputFile As Boolean
Private IsOutputCell As Boolean
Private IsOutputError As Boolean
Private IsOutputWarning As Boolean
Private IsOutputInfo As Boolean
Private IsOutputDetail As Boolean

' @breif ����������
Function CheckedResult_Initialize()
    LogApiIn "CheckedResult_Initialize()"

    LoadWorksheets
    
    Set CheckedResultList = New Collection

    LogApiOut "CheckedResult_Initialize()"
End Function

' @breif �V�[�g����{���W���[���ɕK�v�ȏ���ǂݍ���
Private Function LoadWorksheets()
    LogApiIn "LoadWorksheets()"

    ' �`�F�b�N�}�N������ݒ��ǂݍ���
    ' ���s�������l����Ƃ����̃A�_�v�^�͈ꌳ�Ǘ������ق����ǂ�
    With ThisWorkbook.Worksheets("�`�F�b�N")
        IsOutputFile = .Range("IsOutputFile")
        IsOutputCell = .Range("IsOutputCell")
        IsOutputDetail = .Range("IsOutputDetail")
    
        IsOutputError = .Range("IsOutputErrorSky")
        IsOutputWarning = .Range("IsOutputWarningSky")
        IsOutputInfo = .Range("IsOutputInfoSky")
    End With

    LogApiOut "LoadWorksheets()"
End Function

' @breif �I��������
Function CheckedResult_Terminate()
    LogApiIn "CheckedResult_Terminate()"

    Set CheckedResultList = Nothing

    LogApiOut "CheckedResult_Terminate()"
End Function

' @breif ���ʂ�ǉ�����
Public Function AddResult(rsltType As String, target As String, content As String, fullPath As String)
    LogApiIn "AddResult()"
    
    Dim result As CheckedResult
    Set result = New CheckedResult
    With result
        .resultType = rsltType
        .target = target
        .content = content
        .fullPath = fullPath
    End With
    
    CheckedResultList.Add result
    
    Set result = Nothing
    
    LogApiOut "AddResult()"
End Function

' @breif �`�F�b�N���ʂ��o�͂���
Public Function OutputResult()
    LogApiIn "OutputResult()"
    
    If IsOutputFile Then
        WriteFile
    End If

    If IsOutputCell Then
        WriteCell
    End If
    
    If IsOutputDetail Then
        WriteWorksheet
    End If
    
    LogApiOut "OutputResult()"
End Function

' @breif �`�F�b�N���ʂ��Z���ɏo�͂���
Private Function WriteCell()
    LogApiIn "WriteCell()"
    
    ' �o�͐�N���A
    Dim output As Range
    Set output = ThisWorkbook.Worksheets(OutputWorksheet).Range(OutputCell)
    output.Clear
    
    ' �o��
    output = FormatCheckedResult
    
    LogApiOut "WriteCell()"
End Function

' @breif �o�͗p�Ƀ`�F�b�N���ʂ𐮌`����
' @note �O�̋Ζ��\�`�F�b�N�}�N�����Q�l�ɂ���Ƃ��������Ƃ���𓥏P���Ȃ��Ƃ����Ȃ��̂����邢
' @note �ł������ɂ���Ǝg�p�҂����Ă��Ȃ��Ƃ����̂����邵
Private Function FormatCheckedResult() As String
    LogApiIn "FormatCheckedResult()"
    
    Dim context As String
    Dim name As String
    name = ""
    
    Dim result As CheckedResult
    Set result = New CheckedResult
    For Each result In CheckedResultList
    
        If NeedOutput(result) Then
            FormPersonalCheckedResult result, context
        End If
    
    Next result
    
    FormatCheckedResult = context
    
    LogApiOut "FormatCheckedResult()"
End Function

' @breif ���ʏo�̗͂v�ۂ𔻒肷��
Private Function NeedOutput(result As CheckedResult) As Boolean
    LogApiIn "NeedOutput()"
    
    If result.resultType = "Error" And IsOutputError Then
        NeedOutput = True
    ElseIf result.resultType = "Warning" And IsOutputWarning Then
        NeedOutput = True
    ElseIf result.resultType = "Info" And IsOutputInfo Then
        NeedOutput = True
    Else
        NeedOutput = False
    End If
    
    LogApiOut "NeedOutput()"
End Function

' @breif ��l���̃`�F�b�N���ʂ��o�͂���
' @note [�o�O�L]beforeTarget�̓v���O�����̏I���ŏ���������Ȃ��B
'       ���̂��߁A�`�F�b�N�Ώېl������l�̂Ƃ��Ƀ`�F�b�N������ƁA���ڈȍ~�͌��o������������Ȃ�
'       ���Ȃ�beforeTarget���O���[�o���Ɏ����Ă����B���͌����ȃR�[�h�Ȃ̂ł��Ȃ�
Private Function FormPersonalCheckedResult(ByRef result As CheckedResult, ByRef context As String)
    LogApiIn "FormPersonalCheckedResult()"

    ' �ŏ��̏o�͂̏ꍇ�͌��o��������
    Static beforeTarget As String
    If beforeTarget <> result.target Then
        context = context & FormHeading(result)
    End If
    beforeTarget = result.target
    
    context = context & "[" & result.resultType & "]" & result.content & vbCrLf
    
    LogApiOut "FormPersonalCheckedResult()"
End Function

' @breif �`�F�b�N���ʏo�͗p�̌��o���𐶐�����
Private Function FormHeading(ByRef result As CheckedResult) As String
    LogApiIn "FormHeader()"

    Dim heading As String
    heading = "�� " & result.target & vbCrLf
    heading = heading + result.fullPath & vbCrLf
    
    Dim ret As Boolean
    Dim dateLastModified As Date
    ret = GetDateLastModified(result.fullPath, dateLastModified)
    If ret = True Then
        heading = heading & "�ŏI�X�V����(" & format(dateLastModified, "yyyy/mm/dd hh:nn") & ")���_�̃t�@�C���ɑ΂��ă`�F�b�N���s���܂����B" & vbCrLf
    End If
    
    FormHeading = heading
    
    LogApiOut "FormHeader()"
End Function

' @breif �`�F�b�N���ʂ����[�N�V�[�g�ɏo�͂���
Private Function WriteWorksheet()
    LogApiIn "WriteWorksheet()"
    
    ' �o�͐�N���A
    ClearRange
    
    ' ���X�g���o��
    OutputList
    
    LogApiOut "WriteWorksheet()"
End Function

' @breif ���O�o�͐��Range���N���A����
Private Function ClearRange()
    LogApiIn "ClearRange()"

    ' �N���A����s��
    Const MaxColumnOffset As Long = 2000

    ' �N���A�͈͂̎Z�o
    Dim output As Range
    Set output = ThisWorkbook.Worksheets(OutputWorksheet).Range(OutputStartingPoint) _
                    .Resize(MaxColumnOffset, ColumnOutputRange.ColCount)
        
    ' �N���A
    output.Clear

    LogApiOut "ClearRange()"
End Function

' �`�F�b�N���ʂ�\�o�͂���
Private Function OutputList()
    LogApiIn "OutputList()"

    ' �o�͐�͈͂̎Z�o
    Dim output As Range
    Set output = ThisWorkbook.Worksheets(OutputWorksheet).Range(OutputStartingPoint) _
                    .Resize(CheckedResultList.count(), ColumnOutputRange.ColCount)
    
    ' �o�̓f�[�^�̐���
    Dim data As Variant
    data = GenerateVariant
    
    ' �o��
    output = data

    LogApiOut "OutputList()"
End Function


' @breif �`�F�b�N���ʃ��X�g����Variant�^�z��𐶐�����
Private Function GenerateVariant() As Variant
    LogApiIn "GenerateVariant()"
    
    ' �z��̃T�C�Y������(List�T�C�Y+1�̑傫��)
    Dim ret As Variant
    ReDim ret(CheckedResultList.count, ColumnOutputRange.ColCount)
    
    ' �N���X�^�z��Variant�^�z��ɕϊ�
    Dim i As Long
    For i = 0 To CheckedResultList.count - 1
        Dim result As CheckedResult
        Set result = New CheckedResult
        
        With CheckedResultList.Item(i + 1)
            ret(i, ColType) = .resultType
            ret(i, ColTarget) = .target
            ret(i, ColContent) = .content
            ret(i, ColFullPath) = .fullPath
        End With
        
        Set result = Nothing
    Next i
    
    GenerateVariant = ret

    LogApiOut "GenerateVariant()"
End Function

' @breif �`�F�b�N���ʂ��t�@�C���o�͂���
Private Function WriteFile()
    LogApiIn "WriteFile()"
    
    If IsOutputFile = False Then
        Exit Function
    End If

    Dim fileNumber
    fileNumber = FreeFile()
    
    On Error Resume Next
    Open GenerateFullName(GetDirCheckedResult, GetFileCheckedResult) For Output As #fileNumber
    If Err.Number <> 0 Then
        LogError "Cannot open log file(" & GenerateFullName(GetDirCheckedResult, GetFileCheckedResult) & ")! " _
        & "ErrNo:" & Err.Number & "ErrDescription:" & Err.Description & "ErrFunction:OutputLogFile()"
    End If
    Print #fileNumber, FormatCheckedResult
    Close #fileNumber
    
    LogApiOut "WriteFile()"
End Function

' @breif �����ɍ���Result���������邩�擾����
' @note �����Ȃ��̏ꍇ�͂��ׂč��v�Ƃ����ӂ��ɂ������������A���܂��������@���v�����Ȃ�����
Public Function GetCountReuslt(Optional rsltType As String = "", Optional target As String = "") As Long
    LogApiIn "GenerateVariant()"
    
    Dim count As Long
    count = 0
    
    Dim result As CheckedResult
    Set result = New CheckedResult
    For Each result In CheckedResultList
        If (result.resultType = rsltType) And (result.target = target) Then
            count = count + 1
        End If
    Next result
    
    GetCountReuslt = count
    
    LogApiOut "GenerateVariant()"
End Function


