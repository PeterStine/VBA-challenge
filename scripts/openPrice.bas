Attribute VB_Name = "openPrice"
'Author: Peter Stine
'Created: 4/8/2023
'Documentation used:
'   https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-conceptual-topics
'   https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-language-reference
'   https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-fornext-statements

'This will create a single column of all the different
'Ticker Symbols
'
Sub openPrice()

    'Declare types
    Dim openPriceAZTM As Single
    Dim listLength As Long
    
    'Set column header name
    Cells(1, 9).Value = Ticker
    
    'Declare default values
    openPriceAZTM = 0
    listLength = 22770
    
    'Loop through <open> column 22,770 times
    For i = 0 To 22770
        openPriceAZTM = openPriceAZTM + Cells(i + 1, 3).Value
    'Increment counter and jump back to start of loop
    Next i
    
    'Write the summative value of the symbol to
    'A conspicuous location
    Range("I2").Value = openPriceAZTM
    
End Sub

