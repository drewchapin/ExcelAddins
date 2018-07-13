Attribute VB_Name = "FormatQuote"
Option Explicit

Private Function IsWhitespace(Value As String) As Boolean
    Dim a As Integer: a = Asc(Value)
    IsWhitespace = a <= 32 Or (a >= 127 And a <= 160)
End Function


Public Sub FormatQuote()
    
    ''' https://superuser.com/questions/791151/outlook-vba-how-to-copy-the-currently-selected-text-into-clipboard
    Dim Selection As Object, s As Integer, e As Integer
    Set Selection = Application.ActiveInspector.CurrentItem.GetInspector.WordEditor.Application.Selection
    
    ' Whitespace selected?
    Do While IsWhitespace(Left(Selection.Text, 1)) And Selection.Start < Selection.End
        Selection.Start = Selection.Start + 1
    Loop
    Do While IsWhitespace(Right(Selection.Text, 1)) And Selection.Start < Selection.End
        Selection.End = Selection.End - 1
    Loop
    
    ' Italic Blue
    Selection.Font.Bold = False
    Selection.Font.Color = RGB(0, 112, 192)
    Selection.Font.Italic = True
    Selection.Font.Name = "Calibri"
    
    ' Add quotes?
    If Left(Selection.Text, 1) <> """" Then Selection.Text = """" & Selection.Text
    If Right(Selection.Text, 1) <> """" Then Selection.Text = Selection.Text & """"
    
    ' Get start and end
    s = Selection.Start
    e = Selection.End
    
    ' Format beginning quote
    Selection.End = s + 1
    Selection.Font.Bold = True
    Selection.Font.Color = RGB(192, 0, 0)
    Selection.Font.Italic = False
    Selection.Font.Name = "Consolas"
    
    ' Format ending quote
    Selection.Start = e - 1
    Selection.End = e
    Selection.Font.Bold = True
    Selection.Font.Color = RGB(192, 0, 0)
    Selection.Font.Italic = False
    Selection.Font.Name = "Consolas"
    
    Selection.Start = s
    Selection.End = e

End Sub
