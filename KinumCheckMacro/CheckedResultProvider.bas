Attribute VB_Name = "CheckedResultProvider"
' @breif チェック結果を管理するクラス
' @note チェック結果提供部と保管部で分けてもいい気はする
Option Explicit

' @breif Enumを使ったほうが性能的には早いんだけど、今更変更するのが面倒で使わなかった
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

Private Const OutputWorksheet As String = "チェック結果"
Private Const OutputCell As String = "E3"
Private Const OutputStartingPoint As String = "C7"

Private CheckedResultList As Collection
Private IsOutputFile As Boolean
Private IsOutputCell As Boolean
Private IsOutputError As Boolean
Private IsOutputWarning As Boolean
Private IsOutputInfo As Boolean
Private IsOutputDetail As Boolean

' @breif 初期化処理
Function CheckedResult_Initialize()
    LogApiIn "CheckedResult_Initialize()"

    LoadWorksheets
    
    Set CheckedResultList = New Collection

    LogApiOut "CheckedResult_Initialize()"
End Function

' @breif シートから本モジュールに必要な情報を読み込む
Private Function LoadWorksheets()
    LogApiIn "LoadWorksheets()"

    ' チェックマクロから設定を読み込む
    ' 実行効率を考えるとここのアダプタは一元管理したほうが良い
    With ThisWorkbook.Worksheets("チェック")
        IsOutputFile = .Range("IsOutputFile")
        IsOutputCell = .Range("IsOutputCell")
        IsOutputDetail = .Range("IsOutputDetail")
    
        IsOutputError = .Range("IsOutputErrorSky")
        IsOutputWarning = .Range("IsOutputWarningSky")
        IsOutputInfo = .Range("IsOutputInfoSky")
    End With

    LogApiOut "LoadWorksheets()"
End Function

' @breif 終了化処理
Function CheckedResult_Terminate()
    LogApiIn "CheckedResult_Terminate()"

    Set CheckedResultList = Nothing

    LogApiOut "CheckedResult_Terminate()"
End Function

' @breif 結果を追加する
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

' @breif チェック結果を出力する
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

' @breif チェック結果をセルに出力する
Private Function WriteCell()
    LogApiIn "WriteCell()"
    
    ' 出力先クリア
    Dim output As Range
    Set output = ThisWorkbook.Worksheets(OutputWorksheet).Range(OutputCell)
    output.Clear
    
    ' 出力
    output = FormatCheckedResult
    
    LogApiOut "WriteCell()"
End Function

' @breif 出力用にチェック結果を整形する
' @note 前の勤務表チェックマクロを参考にするとこういうところを踏襲しないといけないのがだるい
' @note でも無下にすると使用者がついてこないというのもあるし
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

' @breif 結果出力の要否を判定する
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

' @breif 一人分のチェック結果を出力する
' @note [バグ有]beforeTargetはプログラムの終了で初期化されない。
'       そのため、チェック対象人数が一人のときにチェックをすると、二回目以降は見出しが生成されない
'       やるならbeforeTargetをグローバルに持っていく。私は嫌いなコードなのでやらない
Private Function FormPersonalCheckedResult(ByRef result As CheckedResult, ByRef context As String)
    LogApiIn "FormPersonalCheckedResult()"

    ' 最初の出力の場合は見出しをつける
    Static beforeTarget As String
    If beforeTarget <> result.target Then
        context = context & FormHeading(result)
    End If
    beforeTarget = result.target
    
    context = context & "[" & result.resultType & "]" & result.content & vbCrLf
    
    LogApiOut "FormPersonalCheckedResult()"
End Function

' @breif チェック結果出力用の見出しを生成する
Private Function FormHeading(ByRef result As CheckedResult) As String
    LogApiIn "FormHeader()"

    Dim heading As String
    heading = "■ " & result.target & vbCrLf
    heading = heading + result.fullPath & vbCrLf
    
    Dim ret As Boolean
    Dim dateLastModified As Date
    ret = GetDateLastModified(result.fullPath, dateLastModified)
    If ret = True Then
        heading = heading & "最終更新日時(" & format(dateLastModified, "yyyy/mm/dd hh:nn") & ")時点のファイルに対してチェックを行いました。" & vbCrLf
    End If
    
    FormHeading = heading
    
    LogApiOut "FormHeader()"
End Function

' @breif チェック結果をワークシートに出力する
Private Function WriteWorksheet()
    LogApiIn "WriteWorksheet()"
    
    ' 出力先クリア
    ClearRange
    
    ' リストを出力
    OutputList
    
    LogApiOut "WriteWorksheet()"
End Function

' @breif ログ出力先のRangeをクリアする
Private Function ClearRange()
    LogApiIn "ClearRange()"

    ' クリアする行数
    Const MaxColumnOffset As Long = 2000

    ' クリア範囲の算出
    Dim output As Range
    Set output = ThisWorkbook.Worksheets(OutputWorksheet).Range(OutputStartingPoint) _
                    .Resize(MaxColumnOffset, ColumnOutputRange.ColCount)
        
    ' クリア
    output.Clear

    LogApiOut "ClearRange()"
End Function

' チェック結果を表出力する
Private Function OutputList()
    LogApiIn "OutputList()"

    ' 出力先範囲の算出
    Dim output As Range
    Set output = ThisWorkbook.Worksheets(OutputWorksheet).Range(OutputStartingPoint) _
                    .Resize(CheckedResultList.count(), ColumnOutputRange.ColCount)
    
    ' 出力データの生成
    Dim data As Variant
    data = GenerateVariant
    
    ' 出力
    output = data

    LogApiOut "OutputList()"
End Function


' @breif チェック結果リストからVariant型配列を生成する
Private Function GenerateVariant() As Variant
    LogApiIn "GenerateVariant()"
    
    ' 配列のサイズを決定(Listサイズ+1の大きさ)
    Dim ret As Variant
    ReDim ret(CheckedResultList.count, ColumnOutputRange.ColCount)
    
    ' クラス型配列→Variant型配列に変換
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

' @breif チェック結果をファイル出力する
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

' @breif 条件に合うResultが何件あるか取得する
' @note 引数なしの場合はすべて合致というふうにしたかったが、うまい実装方法が思いつかなかった
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


