'Author: Peter Stine
'Created: 4/8/2023
'Documentation used:
'   https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-conceptual-topics
'   https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-language-reference
'   https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-fornext-statements

'This will create a single column of all the different
'Ticker symbols condensed into camera
'
Sub openPrice()
    'openPriceAZTM initialize to 0
    openPriceAZTM = 0
    'Loop through <open> column 22,770 times
    For i = 0 To 22770
        'If symbol is AZTM
        If cell(i + 1, 1).Value = "AZTM" Then
            'openPriceAZTM = openPriceAZTM + 1
        'If Symbol is not AZTM
            'continue loop
        'End IF
    'Increment counter and jump back to start of loop
    Next i
End Sub

'Sub closePrice()

'End Sub

