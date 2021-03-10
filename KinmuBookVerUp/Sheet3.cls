VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Enum ColumnList
    ColumnOldFullPath
    ColumnNewFullPath
End Enum

Const RangeList As String = ""
Const RangeTemplateFullPath As String = ""

Function UpgradeProperWorkbook()

    Dim list As Variant
    list = Range(RangeList)
    
    Dim i As Long
    For i = 0 To UBound(list)
        FileCopy RangeTemplateFullPath, list(i, ColumnNewFullPath)
        UpgradeProper list(i, ColumnOldFullPath), list(i, ColumnNewFullPath)
    Next

End Function


