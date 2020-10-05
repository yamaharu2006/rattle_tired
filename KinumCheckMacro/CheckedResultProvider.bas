Attribute VB_Name = "CheckedResultProvider"
' @breif �`�F�b�N���ʂ��Ǘ�����N���X
' @note �`�F�b�N���ʒ񋟕��ƕۊǕ��ŕ����Ă������C�͂���
Option Explicit

' @breif Enum���g�����ق������\�I�ɂ͑����񂾂��ǁA���X�ύX����̂��ʓ|�Ŏg��Ȃ�����
Public Enum resultType
    Err
    Warning
    Info
End Enum

Public Enum ColumnOutputRange
    ColType
    ColTarget
    ColContent
    ColFullPath
    ColCount
End Enum

Private Const sheetChecking As String = "�`�F�b�N"

Private Const RangeIsOutputFile As String = "IsOutputFile"
Private Const RangeIsOutputCell As String = "IsOutputCell"
Private Const RangeIsOutputError As String = "IsOutputError"
Private Const RangeIsOutputWarning As String = "IsOutputWarning"
Private Const RangeIsOutputInfo As String = "IsOutputInfo"
Private Const RangeIsOutputDetail As String = "IsOutputDetail"

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
    With ThisWorkbook.Worksheets(sheetChecking)
        IsOutputFile = .Range(RangeIsOutputFile)
        IsOutputCell = .Range(RangeIsOutputCell)
    
        IsOutputError = .Range(RangeIsOutputError)
        IsOutputWarning = .Range(RangeIsOutputWarning)
        IsOutputInfo = .Range(RangeIsOutputInfo)
        IsOutputDetail = .Range(RangeIsOutputDetail)
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
Public Function AddResult(rsltType As String, target As String, Content As String, fullPath As String)
    LogApiIn "AddResult()"
    
    Dim result As CheckedResult
    Set result = New CheckedResult
    With result
        .resultType = rsltType
        .target = target
        .Content = Content
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
        ' Unimplemented
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
Private Function FormPersonalCheckedResult(ByRef result As CheckedResult, ByRef context As String)
    LogApiIn "FormPersonalCheckedResult()"

    Static beforeTarget As String

    ' �ŏ��̏o�͂̏ꍇ�͌��o��������
    ' �ăR���p�C�������܂ŐÓI�ϐ��̒l���ς��Ȃ����Ƃ�����̂ł��̗\�h��Ƃ���context��""���ǂ����m�F���Ă���
    If beforeTarget <> result.target Or context = "" Then
    
        Dim dateLastModified As Date
        dateLastModified = GetDateLastModified(result.fullPath)
    
        context = context + "�� " & result.target & vbCrLf
        context = context + result.fullPath & vbCrLf
        context = context + "�ŏI�X�V����(" & Format(dateLastModified, "yyyy/mm/dd hh:nn") & ")���_�̃t�@�C���ɑ΂��ă`�F�b�N���s���܂����B" & vbCrLf
    
    End If
    beforeTarget = result.target
    
    context = context & "[" & result.resultType & "]" & result.Content & vbCrLf
    
    LogApiOut "FormPersonalCheckedResult()"
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
            ret(i, ColContent) = .Content
            ret(i, ColFullPath) = .fullPath
        End With
        
        Set result = Nothing
    Next i
    
    GenerateVariant = ret

    LogApiOut "GenerateVariant()"
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


