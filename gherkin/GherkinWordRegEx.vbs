'========================================================================
' @description:
' GherkinWord()
'  Apply Gherkin Syntax highlight to the current selection in Microsoft Word
'  Highlight non-ascii characters - it is suggested to replace them with regular ASCII chars
' ClearGherkinHighlight()
' Removes formatting introduced by script
' Note: Indent might be impacted
' @Requires: Regular Expressions
' Must add reference (Tools > References) to the
' "Microsoft VBScript Regular Expressions 5.5" Object Library
'
' @Author:
' Anatoliy Sakhno   Anatoly.Sakhno@gmail.com
'  @Date   :    5/9/2019
'  @Version : 1.1
'========================================================================

Sub GherkinWord()
Dim singleLine As Paragraph
Dim clauses(11) As String
clauses(0) = "^\s*?Feature:"
clauses(1) = "^\s*?When\s+"
clauses(2) = "^\s*?Then\s+"
clauses(3) = "^\s*?Scenario:"
clauses(4) = "^\s*?Given\s+"
clauses(5) = "^\s*?Scenario Outline:"
clauses(6) = "^\s*?Background:"
clauses(7) = "When\u00a0"
clauses(8) = "Then\u00a0"
clauses(9) = "^\s*?Scenario:"
clauses(10) = "Given\u00a0"
Dim operands(2) As String
operands(0) = "^\s*?And\s+"
operands(1) = "^\s*?But\s+"
operands(2) = "^\s*?And\u00a0"
operands(3) = "^\s*?But\u00a0"
Dim checks(3) As String
checks(0) = "\s+VERIFY\s+"
checks(1) = "\s+CONFIRM\s+"
checks(2) = "\s+ASSERT\s+"
Dim comments(1) As String
comments(0) = "#.*?$"
Dim constants(3) As String
constants(0) = "\s[0-9]+\s"
constants(1) = "\"".*?\"""
constants(2) = "\<.*?\>"
Dim Tags(1) As String
Tags(0) = "@[0-9A-z_-]+"
Dim unicodes(1) As String
unicodes(0) = "[^\u0000-\u007F]"

clausesColor = RGB(64, 64, 255)
operandColor = RGB(64, 64, 255)
checkColor = RGB(140, 50, 50)
CommentsColor = RGB(64, 168, 64)
constcolor = RGB(66, 134, 244)
tagcolor = RGB(100, 0, 100)
unicodeColor = RGB(232, 13, 24)

 For Each singleLine In Selection.Paragraphs
        singleLine.Outdent ' outdent to makesure lines are not shifted on repeated use
        Call formatString(singleLine, clauses, clausesColor, True, True)    '// given whrn rhen
        Call formatString(singleLine, operands, operandColor, True, False) ' and but
  '      Call formatString(singleLine, checks, checkColor, False, False)    ' VERIFY CONFIRM
        Call formatString(singleLine, constants, constcolor, False, False)   ' Constants
        Call formatString(singleLine, Tags, tagcolor, False, False)  ' Tags
        Call formatString(singleLine, comments, CommentsColor, False, False) ' #comments
        Call formatString(singleLine, unicodes, unicodeColor, False, True)  ' Draw attention to nonascii code Unicode
        '// parse the text here...
 Next singleLine

End Sub

'apply particular regex category
'
'
Sub formatString(line As Paragraph, SearchPatternArray() As String, color As Variant, isStartOfLine As Boolean, isBold As Boolean)

Dim mcolResults As MatchCollection
Dim r As Match
Dim st As String
Dim adjustedRange As Range

For Each substring In SearchPatternArray
   st = substring
 Set mcolResults = RegEx(line.Range.Text, st, True, True, True)

If Not mcolResults Is Nothing Then
        For Each r In mcolResults
        LPosition = InStr(1, line.Range.Text, r)
        Do While LPosition > 0
                Set adjustedRange = line.Range
                adjustedRange.MoveStart Unit:=wdCharacter, Count:=LPosition - 1
                adjustedRange.Collapse Direction:=wdCollapseStart
                adjustedRange.MoveEnd Unit:=wdCharacter, Count:=Len(r)
                adjustedRange.Font.color = color
                adjustedRange.Font.Bold = isBold
                If (Not isBold) And (isStartOfLine) Then
                    line.Indent
                End If
            LPosition = InStr(LPosition + 1, line.Range.Text, r)
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
            '.IgnoreCase = IgnoreCase
            .Pattern = strPattern
        End With
    
        If objRegEx.Test(strInput) Then
            Set mcolResults = objRegEx.Execute(strInput)
            Set RegEx = mcolResults
        End If
    End If
End Function

'
' Clear Gherkin Formatting
'
Sub ClearGherkinHighlight()
For Each singleLine In Selection.Paragraphs
        singleLine.Outdent ' outdent to make sure lines are not shifted on repeated use
        singleLine.Range.Font.Bold = False
        singleLine.Range.Font.color = RGB(0, 0, 0)
    Next singleLine
End Sub
Remove







