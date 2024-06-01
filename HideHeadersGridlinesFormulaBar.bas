Sub HideHeadersGridlinesFormulaBar()
'
' HideHeadersGridlinesFormulaBar
'
If ActiveWindow.DisplayHeadings = False And _
   ActiveWindow.DisplayGridlines = False And _
   Application.DisplayFormulaBar = False Then
    
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayGridlines = True
    Application.DisplayFormulaBar = True
    
Else
    If ActiveWindow.DisplayHeadings = True Or _
        ActiveWindow.DisplayGridlines = True Or _
        Application.DisplayFormulaBar = True Then
        
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayGridlines = False
        Application.DisplayFormulaBar = False
        
    End If
        
End If
         
End Sub
