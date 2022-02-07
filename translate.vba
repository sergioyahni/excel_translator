Option Explicit     'Force variable declaration

Dim lastCol As Integer
Dim inCol As Integer
Dim language As String
Dim activeLanguageSheet As String
Sub RunTranslation()
   ' Cancel text wrap if exists, set row height to 15.75
    Columns("A:B").Select
    With Selection
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Cells.RowHeight = 15.75
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    ' Cells.WrapText = False
    ' Cells.RowHeight = 15.75

    ' set all IDs in column A as int
    Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
    ' align ID numbers in column A to the right
    Columns("A:A").HorizontalAlignment = xlRight
    
    CreateLanguagesSheet
    saveTranslated
End Sub
Private Sub CreateLanguagesSheet()
    ' Creates Sheets according to the languages specified in Language
    lastCol = Sheets("Language").Cells(1, Columns.Count).End(xlToLeft).Column
    For inCol = 2 To lastCol
    language = Sheets("Language").Cells(1, inCol).Value
    activeLanguageSheet = "Opening Hours_" & language
         With ThisWorkbook
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = activeLanguageSheet
    End With
    'Populate sheet according to language, translate, save translation and quit application
    populateLanguageSheet
    translateContent
    Next
End Sub

Private Sub populateLanguageSheet()
' populate language sheet with content to be translated
    Sheets("Report").Select
        Columns("A:B").Select
        Selection.Copy
        Sheets(activeLanguageSheet).Select
        Range("A1").Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
End Sub

Private Sub translateContent()
    Dim orgValue As String
    Dim destValue As String
    Dim lastRow As Integer
    Dim j As Integer
    
    lastRow = Sheets("Language").Cells(Rows.Count, 1).End(xlUp).Row
    For j = 2 To lastRow
        orgValue = Sheets("Language").Cells(j, 1).Value
        destValue = Sheets("Language").Cells(j, inCol).Value
        
        ' replace terms row by row
        Sheets(activeLanguageSheet).Select
        Cells.Replace What:=orgValue, Replacement:=destValue, LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        
        'fix heading to Opening Hours_language
        Cells(1, 2).Value = activeLanguageSheet
    Next
    Cells.WrapText = False
End Sub

Private Sub saveTranslated()
    Dim MyFolder As String
    Dim MySaveFileName As String
    Dim MyDate As String
    
    'Build the date stamp
    ' MyDate = Mid(Now(), 7, 4) & Mid(Now(), 4, 2) & Left(Now(), 2) & " "
    MyDate = Format(Now, "yyyymmddhhmmss")

    'Specifies where the backup will be stored (ensure there is a slash at the very end)
    MyFolder = "C:\Users\SergioY\Documents\Opening Hours\AfterTranslation"
    
    'Build the file name
    MySaveFileName = MyFolder & "\Opening Hours-" & MyDate & ".xlsx"
        
    ' close VB window
    ' ThisWorkbook.VBProject.VBE.MainWindow.Visible = False
    
    'Save a copy
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=MySaveFileName, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = False
End Sub
