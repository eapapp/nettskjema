VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Update the curation overview table"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8205
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkUseCurrentSheet_Click()

    If txtOverview.Enabled = True Then txtOverview.Enabled = False Else txtOverview.Enabled = True

End Sub

Private Sub CommandButton1_Click()

    Dim strFile

    strFile = Application.GetOpenFilename(FileFilter:="Excel files (*.xlsx*), *.xlsx*", Title:="Select the existing overview table", MultiSelect:=False)
    txtOverview.Text = strFile

End Sub

Private Sub CommandButton2_Click()

    Dim strFile

    strFile = Application.GetOpenFilename(FileFilter:="Excel files (*.xlsx*), *.xlsx*", Title:="Select the export from Nettskjema", MultiSelect:=False)
    txtNettskjema.Text = strFile

End Sub

Private Sub CommandButton3_Click()

Dim shNettskjema As Worksheet
Dim shOverview As Worksheet
Dim lastRow As Integer
Dim lastNR As String
Dim i As Integer, j As Variant, k As Integer
Dim Cols() As String
Dim NewDataset() As String
Dim NewVersion() As String
'Cols = Split("A C F G H BN")
'Cols = Split("A C G H I BO")
'Cols = Split("A C U V BQ")
'Cols = Split("A C U V W BS")
NewDataset = Split("A C U V W BS")
NewVersion = Split("A C U V W J")
Const ResProd = "M:R"  '"N:R"    '"M:Q"
Const Funding = "BU"  '"BS"  '"AC:AF"  '"AB:AE"
Dim prod As String
Dim ext As String
Dim after As Range, c As Range
Dim match As Integer
Dim offset As Integer
Dim sel As Range

If txtNettskjema.Text = "" Or (txtOverview.Text = "" And chkUseCurrentSheet.Value = False) Then Exit Sub

Application.ScreenUpdating = False

If chkUseCurrentSheet.Value = False Then Workbooks.Open txtOverview.Text, ReadOnly:=False
Set shOverview = ActiveWorkbook.Sheets(1)

Workbooks.Open txtNettskjema.Text, ReadOnly:=True
Set shNettskjema = ActiveWorkbook.Sheets(1)

'lastRow = shOverview.UsedRange.Rows.Count
lastRow = shOverview.Cells(shOverview.Rows.Count, 1).End(xlUp).Row
lastNR = shOverview.Cells(lastRow, 1)

Set after = Range("A" & CStr(shNettskjema.UsedRange.Rows.Count - 1))
shNettskjema.Activate
match = FindNR(lastNR, after)

offset = 0
For i = match + 1 To shNettskjema.UsedRange.Rows.Count

    offset = offset + 1
    
    k = 0
    If shNettskjema.Cells(i, 7) = "" Then Cols = NewDataset Else Cols = NewVersion
    For Each j In Cols
        k = k + 1
        shNettskjema.Activate
        shNettskjema.Range(j & CStr(i)).Copy
        shOverview.Activate
        shOverview.Cells(lastRow + offset, k).Select
        ActiveSheet.Paste
        If j = "BS" And shOverview.Cells(lastRow + offset, k) = "" Then
            shNettskjema.Activate
            shNettskjema.Range("BR" & CStr(i)).Copy
            shOverview.Activate
            shOverview.Cells(lastRow + offset, k).Select
            ActiveSheet.Paste
        End If
        
        If j = "J" Then shOverview.Cells(lastRow + offset, k) = "New version for: " & shOverview.Cells(lastRow + offset, k)
                
        If k = 1 Then
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="https://nettskjema.no/user/form/submission/show.html?id=" & ActiveCell.Text, TextToDisplay:=ActiveCell.Text
        'ElseIf k = 3 Then    'Merge name fields
        '    k = k + 1
        '    ActiveSheet.Range("C" & CStr(lastRow + offset) & ":" & "D" & CStr(lastRow + offset)).Select
        '    Selection.Merge
        End If
    Next j

    shNettskjema.Activate
    Range(Left(ResProd, 1) & CStr(i) & ":" & Right(ResProd, 1) & CStr(i)).Select
    prod = ""
    For Each c In Selection
        If Not c.Value = "" Then
            'Select Case c.Value
            Select Case True
                Case InStr(1, c.Value, "experimental") > 0      '"an experimental study"
                    prod = "dataset"
                Case InStr(1, c.Value, "simulation") > 0        '"a simulation"
                    prod = "dataset"
                Case InStr(1, c.Value, "model") > 0             '"a computational model"
                    prod = "model"
                Case InStr(1, c.Value, "software") > 0          '"a software"
                    prod = "software"
                Case InStr(1, c.Value, "atlas") > 0
                    prod = "atlas"
                Case InStr(1, c.Value, "other") > 0             '"Other(s)"
                    'shNettskjema.Range("AO" & CStr(i)).Copy
                    shNettskjema.Range("S" & CStr(i)).Copy
                    shOverview.Activate
                    shOverview.Cells(lastRow + offset, 6).Select
                    ActiveSheet.Paste
                    shNettskjema.Activate
            End Select
            Exit For
        End If
    Next c
        
    'Range(Left(Funding, 2) & CStr(i) & ":" & Right(Funding, 2) & CStr(i)).Select
    Range(Funding & CStr(i)).Select
    ext = ""
    For Each c In Selection
        If Not c.Value = "" Then
            If c.Value = "Yes" Then ext = "internal" Else ext = "external"
            Exit For
        End If
    Next c
    
    shOverview.Activate
    shOverview.Cells(lastRow + offset, 7) = prod
    shOverview.Cells(lastRow + offset, 8) = ext
    
Next i

shOverview.Activate
Range("F:F").Select
With Selection
    .WrapText = True
    .ShrinkToFit = False
    .MergeCells = False
End With

Range("A2:K" & CStr(lastRow + offset)).Select
Selection.VerticalAlignment = xlTop

shNettskjema.Activate
ActiveWorkbook.Close
UserForm1.Hide
Application.ScreenUpdating = True
shOverview.Activate
Range("A" & CStr(lastRow + 1) & ":H" & CStr(lastRow + offset)).Select

End Sub

Private Sub UserForm_Initialize()

    Dim dldPath As String

    ChDir Application.ActiveWorkbook.Path
    dldPath = Environ("USERPROFILE") & "\Downloads\"
    txtNettskjema.Text = NewestFile(dldPath)

End Sub

Private Function FindNR(lastNR, after) As Integer

On Error GoTo Catch
    
    ActiveSheet.UsedRange.Columns(1).Select
    Selection.Find(lastNR, after, xlValues, xlWhole, xlByRows, xlPrevious).Activate
    FindNR = ActiveCell.Row
    Exit Function
    
Catch:
    FindNR = 1

End Function

Private Function NewestFile(FilePath)

Dim FileName As String
Dim FileSpec As String
Dim MostRecentFile As String
Dim MostRecentDate As Date

FileSpec = "data-*-utf.xlsx"
FileName = Dir(FilePath & FileSpec)

If FileName <> "" Then
    MostRecentFile = FileName
    MostRecentDate = FileDateTime(FilePath & FileName)
    Do While FileName <> ""
        If FileDateTime(FilePath & FileName) > MostRecentDate Then
             MostRecentFile = FileName
             MostRecentDate = FileDateTime(FilePath & FileName)
        End If
        FileName = Dir
    Loop
End If

NewestFile = FilePath & MostRecentFile

End Function
