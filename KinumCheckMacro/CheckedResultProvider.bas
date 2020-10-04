Attribute VB_Name = "CheckedResultProvider"
' @breif チェック結果を管理するクラス
' @note チェック結果提供部と保管部で分けてもいい気はする
Option Explicit

' @breif Enumを使ったほうが性能的には早いんだけど、今更変更するのが面倒で使わなかった
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

Private Const sheetChecking As String = "チェック"

Private Const RangeIsOutputFile As String = "IsOutputFile"
Private Const RangeIsOutputCell As String = "IsOutputCell"
Private Const RangeIsOutputError As String = "IsOutputError"
Private Const RangeIsOutputWarning As String = "IsOutputWarning"
Private Const RangeIsOutputInfo As String = "IsOutputInfo"
Private Const RangeIsOutputDetail As String = "IsOutputDetail"

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

' @breif 終了化処理
Function CheckedResult_Terminate()
    LogApiIn "CheckedResult_Terminate()"

    Set CheckedResultList = Nothing

    LogApiOut "CheckedResult_Terminate()"
End Function

' @breif 結果を追加する
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

' @breif チェック結果を出力する
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
Private Function FormPersonalCheckedResult(ByRef result As CheckedResult, ByRef context As String)
    LogApiIn "FormPersonalCheckedResult()"

    Static beforeTarget As String

    ' 最初の出力の場合は見出しをつける
    ' 再コンパイルされるまで静的変数の値が変わらないことがあるのでその予防策としてcontextが""かどうか確認している
    If beforeTarget <> result.target Or context = "" Then
    
        Dim dateLastModified As Date
        dateLastModified = GetDateLastModified(result.fullPath)
    
        context = context + "■ " & result.target & vbCrLf
        context = context + result.fullPath & vbCrLf
        context = context + "最終更新日時(" & Format(dateLastModified, "yyyy/mm/dd hh:nn") & ")時点のファイルに対してチェックを行いました。" & vbCrLf
    
    End If
    beforeTarget = result.target
    
    context = context & "[" & result.resultType & "]" & result.Content & vbCrLf
    
    LogApiOut "FormPersonalCheckedResult()"
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
            ret(i, ColContent) = .Content
            ret(i, ColFullPath) = .fullPath
        End With
        
        Set result = Nothing
    Next i
    
    GenerateVariant = ret

    LogApiOut "GenerateVariant()"
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


