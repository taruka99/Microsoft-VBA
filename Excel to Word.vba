Sub Button1_Click()

Dim wd As Word.Application
Set wd = New Word.Application
Dim ObjDoc As Object
Dim FilePath As String
Dim FileName As String
Dim ReportName As String

'the file path of excel file
FilePath = ThisWorkbook.Worksheets("Setup").Range("D4").Value
'the file name of the template doc
FileName = ThisWorkbook.Worksheets("Setup").Range("D6").Value

TotalFilePath = FilePath & FileName

'report filepath
'doesn't work
'ThisWorkbook.Worksheets("setup").Range("D4").Value = TotalFilePath

'word report name
ReportName = ThisWorkbook.Worksheets("Setup").Range("D8").Value

'check if template document is open in Word, otherwise open it
On Error Resume Next
Set wd = GetObject(, "Word.Application")

If wd Is Nothing Then
    Set wd = CreateObject("Word.Application")
    Set ObjDoc = wd.Documents.Open(TotalFilePath)
Else
    On Error GoTo notOpen
    Set ObjDoc = wd.Documents(FileName)
    GoTo OpenAlready
notOpen:
    Set ObjDoc = wd.Documents.Open(TotalFilePath)
End If
OpenAlready:
On Error GoTo 0

'make word doc visible
wd.Visible = True

'Edit Certain Texts in a Textbox
Dim oShape As Word.Shape
    If wd.ActiveDocument.Shapes.Count > 0 Then
        For Each oShape In wd.ActiveDocument.Shapes
            If oShape.AutoShapeType = msoShapeRectangle Then 'we need to check both if oShape is of type msoShapeRectangle and its textframe contains place for writing
                If oShape.TextFrame.HasText = True Then
                    'oShape.TextFrame.TextRange.InsertAfter "https://www.automateexcel.com/vba-code-library"
                    oShape.TextFrame.TextRange.Find.Execute FindText:="<<customername>>", _
                        ReplaceWith:=ThisWorkbook.Sheets("Customer Information").Range("B19").Text, Replace:=wdReplaceAll
                    
                    oShape.TextFrame.TextRange.Find.Execute FindText:="<<bforef>>", _
                        ReplaceWith:=ThisWorkbook.Sheets("Customer Information").Range("B20").Text, Replace:=wdReplaceAll
                        
                    oShape.TextFrame.TextRange.Find.Execute FindText:="<<country>>", _
                        ReplaceWith:=ThisWorkbook.Sheets("Customer Information").Range("B18").Text, Replace:=wdReplaceAll
                        
                    oShape.TextFrame.TextRange.Find.Execute FindText:="<<contact>>", _
                        ReplaceWith:=ThisWorkbook.Sheets("Customer Information").Range("B21").Text, Replace:=wdReplaceAll

                    'Exit For 'we just want to write into first textbox
                End If
            End If
        Next oShape
    End If


'Edit Certain Text in a Document
wd.Application.Selection.Find.Execute FindText:="<<country>>", _
    ReplaceWith:=ThisWorkbook.Sheets("Customer Information").Range("B18").Text, Replace:=wdReplaceAll

wd.Application.Selection.Find.Execute FindText:="<<currency>>", _
    ReplaceWith:=ThisWorkbook.Sheets("Customer Information").Range("E18").Text, Replace:=wdReplaceAll

wd.Application.Selection.Find.Execute FindText:="<<annualprice>>", _
    ReplaceWith:=ThisWorkbook.Sheets("Customer Information").Range("E22").Text, Replace:=wdReplaceAll



'Import Images from Excel Sheet into Word with Book Marks
Dim bm As Bookmark
Dim bm2 As Bookmark
Dim bm3 As Bookmark
Dim sh As Shape
Dim sh2 As Shape
Dim sh3 As Shape

If ThisWorkbook.Sheets("Customer Information").Range("B18").Text = "Japan" Then
    Set sh = ThisWorkbook.Sheets("Rate Card Japan").Shapes("Picture 1")
    Set bm = wd.ActiveDocument.Bookmarks("ratecard")  ' your bookmark name here
    
    sh.Copy
    bm.Range.Paste

ElseIf ThisWorkbook.Sheets("Customer Information").Range("B18").Text = "Korea" Then
    Set sh = ThisWorkbook.Sheets("Rate Card Korea").Shapes("Picture 1")
    Set bm = wd.ActiveDocument.Bookmarks("ratecard")  ' your bookmark name here
    
    sh.Copy
    bm.Range.Paste
    
ElseIf ThisWorkbook.Sheets("Customer Information").Range("B18").Text = "Thailand" Then
    Set sh = ThisWorkbook.Sheets("Rate Card Thailand").Shapes("Picture 1")
    Set bm = wd.ActiveDocument.Bookmarks("ratecard")  ' your bookmark name here

    sh.Copy
    bm.Range.Paste
    
ElseIf ThisWorkbook.Sheets("Customer Information").Range("B18").Text = "India" Then
    Set sh = ThisWorkbook.Sheets("Rate Card India").Shapes("Picture 1")
    Set bm = wd.ActiveDocument.Bookmarks("ratecard") ' your bookmark name here
    Set sh2 = ThisWorkbook.Sheets("Rate Card India").Shapes("Picture 2")
    Set bm2 = wd.ActiveDocument.Bookmarks("ratecard2")  ' your bookmark name here
    
    sh.Copy
    bm.Range.Paste
    sh2.Copy
    bm2.Range.Paste
    
ElseIf ThisWorkbook.Sheets("Customer Information").Range("B18").Text = "Indonesia" Then
    Set sh = ThisWorkbook.Sheets("Rate Card Indonesia").Shapes("Picture 1")
    Set bm = wd.ActiveDocument.Bookmarks("ratecard") ' your bookmark name here
    Set sh2 = ThisWorkbook.Sheets("Rate Card Indonesia").Shapes("Picture 2")
    Set bm2 = wd.ActiveDocument.Bookmarks("ratecard2")  ' your bookmark name here
    Set sh3 = ThisWorkbook.Sheets("Rate Card Indonesia").Shapes("Picture 4")
    Set bm3 = wd.ActiveDocument.Bookmarks("ratecard3")  ' your bookmark name here
    
    sh.Copy
    bm.Range.Paste
    sh2.Copy
    bm2.Range.Paste
    sh3.Copy
    bm3.Range.Paste

ElseIf ThisWorkbook.Sheets("Customer Information").Range("B18").Text = "Singapore" Then
    Set sh = ThisWorkbook.Sheets("Rate Card Singapore").Shapes("Picture 1")
    Set bm = wd.ActiveDocument.Bookmarks("ratecard") ' your bookmark name here
    Set sh2 = ThisWorkbook.Sheets("Rate Card Singapore").Shapes("Picture 2")
    Set bm2 = wd.ActiveDocument.Bookmarks("ratecard2")  ' your bookmark name here
    
    sh.Copy
    bm.Range.Paste
    sh2.Copy
    bm2.Range.Paste

ElseIf ThisWorkbook.Sheets("Customer Information").Range("B18").Text = "Taiwan" Then
    Set sh = ThisWorkbook.Sheets("Rate Card Taiwan").Shapes("Picture 1")
    Set bm = wd.ActiveDocument.Bookmarks("ratecard") ' your bookmark name here
    Set sh2 = ThisWorkbook.Sheets("Rate Card Taiwan").Shapes("Picture 2")
    Set bm2 = wd.ActiveDocument.Bookmarks("ratecard2")  ' your bookmark name here
    
    sh.Copy
    bm.Range.Paste
    sh2.Copy
    bm2.Range.Paste

ElseIf ThisWorkbook.Sheets("Customer Information").Range("B18").Text = "Australia" Then
    Set sh = ThisWorkbook.Sheets("Rate Card AU").Shapes("Picture 1")
    Set bm = wd.ActiveDocument.Bookmarks("ratecard") ' your bookmark name here
    Set sh2 = ThisWorkbook.Sheets("Rate Card AU").Shapes("Picture 2")
    Set bm2 = wd.ActiveDocument.Bookmarks("ratecard2")  ' your bookmark name here
    
    sh.Copy
    bm.Range.Paste
    sh2.Copy
    bm2.Range.Paste

ElseIf ThisWorkbook.Sheets("Customer Information").Range("B18").Text = "New Zealand" Then
    Set sh = ThisWorkbook.Sheets("Rate Card NZ").Shapes("Picture 1")
    Set bm = wd.ActiveDocument.Bookmarks("ratecard") ' your bookmark name here
    Set sh2 = ThisWorkbook.Sheets("Rate Card NZ").Shapes("Picture 2")
    Set bm2 = wd.ActiveDocument.Bookmarks("ratecard2")  ' your bookmark name here
    
    sh.Copy
    bm.Range.Paste
    sh2.Copy
    bm2.Range.Paste
    
End If

If bm Is Nothing Then
    MsgBox "Bookmark not found."
    Exit Sub
End If
On Error GoTo 0
'sh.Copy
'bm.Range.Paste
'sh2.Copy
'bm2.Range.Paste
'sh3.Copy
'bm3.Range.Paste


'save as a word doc with the report name
With wd
    .ActiveDocument.SaveAs2 FileName:=wd.ActiveDocument.Path & "\" & ReportName, _
FileFormat:=wdFormatDocumentDefault
End With

'close and set everything to nothing
Set oDoc = Nothing
wd.ActiveDocument.Save
wd.ActiveDocument.Close
wd.Quit
Set wd = Nothing

MsgBox "Successfully Exported to Word Document"

End Sub
