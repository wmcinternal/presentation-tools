Function get_translation(txt As String, column As Integer, trap As Range) As String
    Dim c As Range
    For Each c In trap.Cells:
        On Error GoTo eh
            If Not StrComp(txt, c.Text) Then
                get_translation = c.Offset(ColumnOffset:=column - 2).Text
                Debug.Print , get_translation
            Else
                get_transaltion = txt
            End If
            
    Next
eh:
        Debug.Print , "Oops"
End Function

Function Scrape(ws As Worksheet)
    ' Note : This scrapes English text, if you want to scrape some other language, changes are outlined below
    Dim slides_ As Object, slide_num_cell As String, elt As shape, row_counter As Integer
    Set slides_ = ActivePresentation.Slides
    Set row_counter = 2
    For I = 1 To slides_.Count
        Set slide_num_cell = "A" + (I + 1)
        For Each shape In slides_(I).Shapes
            If shape.HasTextFrame And shape.TextFrame.HasText
                ws.Range(slide_num_cell).SetValue (slide_num_cell)
                ws.Range("B" + I + 1).SetValue (shape.TextFrame.TextRange.Text) ' Change the row alphabet here
                row_counter = row_counter + 1
            End If
        Next
    Next
End Function

Sub TranslateTradCh()
    Dim xlApp As Excel.Application, ws As Excel.Worksheet
    Set xlApp = New Excel.Application
    Set ws = xlApp.Workbooks.Open("C:\Users\amadika\Downloads\Translations.xlsx", True, True).Worksheets(1)
    Dim trap As Range, lastRow As Integer
    Set trap = Range("B1", Range("B1").End(xlDown))
    
    Dim slides_ As Object, shape As shape, a As String
    Set slides_ = ActivePresentation.Slides
    For I = 1 To slides_.Count
        For Each shape In slides_(I).Shapes
            If shape.HasTextFrame And shape.TextFrame.HasText Then
                a = get_translation(shape.TextFrame.TextRange.Text, 3, trap)
                shape.TextFrame.TextRange.Text = a
            End If
        Next
    Next
    ActivePresentation.SaveCopyAs ("WMC Introduction Chinese_Trad")
    xlApp.Quit
End Sub

Sub TranslateSimpCh()
    Dim xlApp As Excel.Application, ws As Excel.Worksheet
    Set xlApp = New Excel.Application
    Set ws = xlApp.Workbooks.Open("C:\Users\amadika\Downloads\Translations.xlsx", True, True).Worksheets(1)
    Dim trap As Range, lastRow As Integer
    Set trap = Range("B1", Range("B1").End(xlDown))
    
    Dim slides_ As Object, shape As shape, a As String
    Set slides_ = ActivePresentation.Slides
    For I = 1 To slides_.Count
        For Each shape In slides_(I).Shapes
            If shape.HasTextFrame And shape.TextFrame.HasText Then
                a = get_translation(shape.TextFrame.TextRange.Text, 4, trap)
                shape.TextFrame.TextRange.Text = a
            End If
        Next
    Next
    ActivePresentation.SaveCopyAs ("WMC Introduction Chinese_Trad")
    xlApp.Quit
End Sub

