Attribute VB_Name = "Module1"
Option Explicit

Function UpgradeProper(oldPath As String, newPath As String)
    
    Dim oldWorkbook As Workbook
    Dim newWorkbook As Workbook
    
    oldWorkbook = Workbooks.Open(oldPath)
    newWorkbook = Workbooks.Open(newPath)
    
    newWorkbook.Worksheets("Sheet1").Range("A1").Value = oldWorkbook.Worksheets("Sheet1").Range("A1").Value
    
    oldWorkbook.Close
    newWorkbook.Close
    
    
End Function

Function UpgradePartner(oldPath As String, newPath As String)
    
    Dim oldWorkbook As Workbook
    Dim newWorkbook As Workbook
    
    oldWorkbook = Workbooks.Open(oldPath)
    newWorkbook = Workbooks.Open(newPath)
    
    newWorkbook.Worksheets("Sheet1").Range("A1").Value = oldWorkbook.Worksheets("Sheet1").Range("A1").Value
    
    oldWorkbook.Close
    newWorkbook.Close
    
    
End Function

