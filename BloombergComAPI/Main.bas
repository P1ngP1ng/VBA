Attribute VB_Name = "Main"
Option Explicit

Public Sub Main()
Dim i As Integer

On Error GoTo Catch

' Clear existing data
Results.Cells.ClearContents

' Get file path
Dim strDir As String
Dim strMask As String
Dim strFile As String

strDir = Control.Range("SourceDirectory") & Application.PathSeparator
strFile = Dir(strDir & "*_VaR_Position.csv")
If Len(strFile) = 0 Then
     Err.Raise vbObjectError + 1, , "Cannot find a file matching the pattern:" & vbCrLf & vbCrLf & strMask & vbCrLf & vbCrLf & "in" & vbCrLf & vbCrLf & strDir
End If

Dim BaseCcy As String
Dim strFund As String
strFund = Control.Range("Fund").Value
Select Case strFund
Case "ARGBF": BaseCcy = "GBP"
Case "SPSF": BaseCcy = "USD"
End Select

Dim RequiredCOB As String
RequiredCOB = Format(Control.Range("RequiredCOBDate").Value, "YYYYMMDD")

Dim arr() As String
Dim bln As Boolean
Dim ccy As String
Dim cob As String
Do While Not bln
     Application.StatusBar = "Checking file: " & strFile
     arr() = ReadFile(strDir & strFile)
     ccy = ReadCcy(arr)
     cob = ReadCOB(arr)
     'Debug.Print ccy, cob, strPath
     If ccy = BaseCcy And cob = RequiredCOB Then
          bln = True
     Else
          strFile = Dir
     End If
Loop

' Get ordinal position of columns
Application.StatusBar = strFile & ": Determining Ordinal Position of Key Columns"
Dim idxSecID As Integer
Dim idxDescr As Integer
Control.Range("COBID").Value = cob

Dim line() As String
line = Split(arr(6), "|")
i = 0
For i = 0 To UBound(line)
     If line(i) = "security_id" Then idxSecID = i
     If line(i) = "description" Then idxDescr = i
     If idxSecID <> 0 And idxDescr <> 0 Then Exit For
Next i


' Output SecID and Descr and clean-up for uniques
Application.StatusBar = "Populating 'Results' worksheet with Source, SecurityID and Description"

Dim r As Long
r = Results.Range("A1048576").End(xlUp).Row
i = 0
Do: i = i + 1: Loop Until arr(i - 1) = "DATA_START"
Do Until arr(i) = "DATA_END"
     line = Split(arr(i), "|")
     With Results
          .Cells(r, 1) = IIf(r = 1, "source", strFund)
          .Cells(r, 2) = line(idxSecID)
          .Cells(r, 3) = line(idxDescr)
     End With
     r = r + 1
     i = i + 1
Loop
Results.Activate
With Results.Range("A1048576").End(xlUp).CurrentRegion
     .RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
     .Sort Key1:=Range("A1"), Header:=xlYes
End With


' Add in field headers
Application.StatusBar = "Populating 'Results' worksheet with Bloomberg FLDS"
FLDS.Range("A1").CurrentRegion.Copy
Results.Range("D1").PasteSpecial Transpose:=True
Results.Rows(1).Font.Name = "Arial"
Results.Rows(1).Font.Size = 8
Results.Rows(1).Font.Bold = True


' Get data from Bloomberg
Application.StatusBar = "Collating securities array..."
Dim b As New BBGCOMAPI
Dim str As String
Dim m As Integer
Dim securities() As Variant
m = Results.Range("B1048576").End(xlUp).Row
ReDim securities(0 To m - 2)
For i = 2 To m
     str = Results.Cells(i, 2).Value
     If Left(str, 3) = "IX-" Then
          securities(i - 2) = Results.Cells(i, 3)
     Else
          securities(i - 2) = "/BUID/" & Results.Cells(i, 2)
     End If
Next i

Application.StatusBar = "Collating fields array..."
Dim n As Integer
Dim fields() As Variant
n = Results.Range("XFD1").End(xlToLeft).Column
ReDim fields(0 To n - 3 - 1)
For i = 4 To n
     str = Results.Cells(1, i).Value
     fields(i - 4) = str
Next i

Application.StatusBar = "Downloading Bloomberg Data..."
Dim arrBBG As Variant
arrBBG = b.getData(securities, fields)
Results.Range("D2", Cells(m, n)) = arrBBG


' Format the results
With Results.Range("A1").CurrentRegion.Font
     .Name = "Arial"
     .Size = 8
End With
Results.Rows(1).Font.Bold = True
With Results.Range("A1").CurrentRegion.Columns
     .AutoFit
     .HorizontalAlignment = xlHAlignLeft
End With


' Sort the results
Dim rng As Range
Set rng = Results.Range("B1048576").End(xlUp).Cells
With Results.Sort
     .SortFields.Clear
     .SortFields.Add Key:=Results.Range("B2", rng), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
     .SetRange rng.CurrentRegion
     .Header = xlYes
     .Apply
End With


Results.Range("A1").Select
Application.StatusBar = ""
MsgBox "Please check the output on the 'Results' worksheet before proceeding to Stage 2.", vbOKOnly, "Stage 1 Complete"
GoTo Finally

Catch:
     
     MsgBox Err.Description, , "Exception"

Finally:
     
     If Not b Is Nothing Then Set b = Nothing
     Application.StatusBar = ""
     Application.ScreenUpdating = True
     
End Sub

Public Sub ExportToCSV()
Dim i As Integer
Dim c As Variant
Dim r As Variant
Dim rng As Range
Dim str As String
Dim strCOBID As String
Dim strPath As String

strCOBID = Control.Range("COBID").Value
strPath = Control.Range("OutputDirectory").Value & Application.PathSeparator & Control.Range("Fund").Value & "_" & strCOBID & "_BBG.csv"

If Len(Dir(strPath)) > 0 Then Kill strPath

i = FreeFile
Open strPath For Output As #i

Print #i, "HEADER_START"
Print #i, "DATA_TYPE=BLOOMBERG"
Print #i, "DATE=" & Mid(strCOBID, 5, 2) & "_" & Right(strCOBID, 2) & "_" & Left(strCOBID, 4)
Print #i, "HEADER_END"
Print #i, "DATA_START"

Set rng = Results.Range("A1048576").End(xlUp).CurrentRegion
For Each r In rng.Rows
     str = ""
     For Each c In rng.Columns
          str = str & rng.Parent.Cells(r.Row, c.Column) & "|"
     Next c
     Print #i, Left(str, Len(str) - 1)
Next r

Print #i, "DATA_END"
Close #i

End Sub

Private Function ReadFile(ByRef strPath As String) As String()
Dim i As Integer
i = FreeFile
Open strPath For Input As #i
ReadFile = Split(Input$(LOF(1), #i), vbLf)
Close #i
End Function

Private Function ReadCcy(ByRef arrFile() As String) As String
ReadCcy = Mid(arrFile(3), 10)
End Function

Private Function ReadCOB(ByRef arrFile() As String) As String
Dim dte() As String
dte = Split(Mid(arrFile(2), 6), "_")
ReadCOB = Format(DateSerial(dte(2), dte(0), dte(1)), "YYYYMMDD")
End Function
