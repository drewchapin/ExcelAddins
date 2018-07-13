Attribute VB_Name = "FormatPicture"
Option Explicit

Public Sub FormatPicture()
    
    Dim WordApp As Word.Application
    Set WordApp = Application.ActiveInspector.CurrentItem.GetInspector.WordEditor.Application
    
    Dim eightPoint5 As Single
    eightPoint5 = WordApp.InchesToPoints(8.5)
    
    Dim m As Word.InlineShape
    For Each m In WordApp.Selection.InlineShapes
       m.Shadow.Style = msoShadowStyleOuterShadow
       m.Shadow.Type = msoShadow21
       m.Borders.OutsideColor = wdColorBlack
       m.Borders.OutsideLineStyle = wdLineStyleSingle
       m.Borders.OutsideLineWidth = wdLineWidth075pt
       If m.Width > eightPoint5 Then: _
            m.Width = eightPoint5
    Next m
    
End Sub

