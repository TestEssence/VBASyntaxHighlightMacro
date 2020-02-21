'========================================================================
' @description:
' GherkinExcel()
'  Apply Gherkin Syntax Highlight to the selected cells in MS EXCEL
'  Highlight non-ascii characters - it is suggested to replace them with regular ASCII chars
' ClearHighlight()
'  a procedure to remove highlights
' @Requires: Regular Expressions
' Must add reference (Tools > References) to the
' "Microsoft VBScript Regular Expressions 5.5" Object Library
'
' @Author:
' Anatoliy Sakhno   Anatoly.Sakhno@gmail.com
' @Date   :    5/10/2019
' @Version : 1.2
'========================================================================

Sub GherkinExcel()
 Dim cel As Range
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    For Each cel In selectedRange.Cells
        Call HighlighCellWithGherkin(cel)
    Next cel
End Sub

' highlight regex constant
Sub HighlighCellWithGherkin(anActiveCell As Object)
Dim clauses(7) As String
clauses(0) = "^\s*?Feature:"
clauses(1) = "^\s*?When"
clauses(2) = "^\s*?Then"
clauses(3) = "^\s*?Scenario:"
clauses(4) = "^\s*?Given"
clauses(5) = "^\s*?Scenario Outline:"
clauses(6) = "^\s*?Background:"
Dim operands(2) As String
operands(0) = "^\s*?And"
operands(1) = "^\s*?But"
Dim checks(3) As String
checks(0) = "VERIFY"
checks(1) = "CONFIRM"
checks(2) = "ASSERT"
Dim comments(1) As String
comments(0) = "#.*?$"
Dim constants(3) As String
constants(0) = "\s[0-9]+\s"
constants(1) = "\"".*?\"""
constants(2) = "\<.*?\>"
Dim Tags(1) As String
Tags(0) = "@[0-9A-z_-]+"
Dim unicode(1) As String
unicode(0) = "[^\u0000-\u007F]"

clausesColor = RGB(64, 64, 255)
operandColor = RGB(64, 64, 255)
checkColor = RGB(140, 50, 50)
CommentsColor = RGB(64, 168, 64)
constcolor = RGB(66, 134, 244)
tagcolor = RGB(100, 0, 100)
unicodeColor = RGB(232, 13, 24)

        Call formatString(anActiveCell, clauses, clausesColor, True)    '// given whrn rhen
        Call formatString(anActiveCell, operands, operandColor, False) ' and but
        Call formatString(anActiveCell, checks, checkColor, False)     ' VERIFY CONFIRM
        Call formatString(anActiveCell, constants, constcolor, False)   ' Constants
        Call formatString(anActiveCell, Tags, tagcolor, False)  ' Tags
        Call formatString(anActiveCell, comments, CommentsColor, False)  ' #comments
        Call formatString(anActiveCell, unicode, unicodeColor, True)  ' Unicode

End Sub

'apply particular regex category
'
'
Sub formatString(aCell As Object, SearchPatternArray() As String, color As Variant, isBold As Boolean)

Dim mcolResults As MatchCollection
Dim r As Match
Dim st As String

For Each substring In SearchPatternArray
   st = substring
 Set mcolResults = RegEx(aCell.Text, st, True, True, True)

If Not mcolResults Is Nothing Then
        For Each r In mcolResults
        LPosition = InStr(1, aCell.Text, r)
        Do While LPosition > 0
                aCell.Characters(Start:=LPosition, Length:=Len(r)).Font.color = color
                aCell.Characters(Start:=LPosition, Length:=Len(r)).Font.Bold = isBold
            Debug.Print r ' remove in production
            LPosition = InStr(LPosition + 1, aCell.Text, r)
         Loop
         DoEvents
        Next r
    End If
Next substring
End Sub

' basic regex
Function RegEx(strInput As String, strPattern As String, _
    Optional GlobalSearch As Boolean, Optional MultiLine As Boolean, _
    Optional IgnoreCase As Boolean) As MatchCollection
    
    Dim mcolResults As MatchCollection
    Dim objRegEx As New RegExp
    
    If strPattern <> vbNullString Then
        
        With objRegEx
            .Global = GlobalSearch
            .MultiLine = MultiLine
            .IgnoreCase = IgnoreCase
            .Pattern = strPattern
        End With
    
        If objRegEx.Test(strInput) Then
            Set mcolResults = objRegEx.Execute(strInput)
            Set RegEx = mcolResults
        End If
    End If
End Function

'removes highlighting introduced by Gherkin format
Sub ClearGherkinHighlight()
Dim cel As Range
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    For Each cel In selectedRange.Cells
        cel.Font.Bold = False
        cel.Font.color = RGB(0, 0, 0)
    Next cel
End Sub

Sub HighlightUnicode()
Dim unicode(1) As String
unicode(0) = "[^\u0000-\u007F]"
unicodeColor = RGB(232, 13, 24)

    Dim cel As Range
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    For Each cel In selectedRange.Cells
        Call formatString(cel, unicode, unicodeColor, True)  ' Unicode
    Next cel
End Sub



