Attribute VB_Name = "Module1"
Sub countTickerSymbols()
    
    'Create a container to hold the number of records'
    'Avoid overflow by using LONG integer types'
    Dim recordQuantity2018, 
        recordQuantity2019, 
        recordQuantity2020 As Long

    'Manually verified record counts in worksheets'
    'Why are there different amounts of records?'
    'Are these records directly comparable?'
    recordQuantity2018 = 753000
    recordQuantity2019 = 756000
    recordQuantity2020 = 759000
    
    'This will pick up trading symbol "AAB"
    tradingSymbol = Range("A2").Value

    'Create for loop to count number of symbols
    'outer loop for each trading symbol
        'inner loop for each individual trading symbol of a set
        'end inner loop
    'end outer loop

    'Forloop to traverse the 2018 worksheet
    For i = 1 To recordQuantity2018 + 1
        If (Cells(i, 1).Value = tradingSymbol) Then

            quantityOfIndividualSymbols += 1
        End If
    Next iterator
    
    MsgBox (quantityOfIndividualSymbols)
    Range("I2").Value = tradingSymbol
    Range("J2").Value = quantityOfIndividualSymbols

'    - WHILE-LOOP
    While Boolean
        
    
'    - **Will only increment counter after a full inner traversal**
'    - *Note:* **This would require unsorted datasets to be sorted.**
'    - FOR-LOOP: 
'        - **This loop will reset it's counter each time it fully    traverses a single trading symbol, such as AAB, AAF, and so on.**
'        - 
'        - innerCounter = innerCounter + 1
'    - END INNER LOOP
'- END OUTER LOOP
End While 'Decided to use a while loop instead 

End Sub