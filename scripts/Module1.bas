Attribute VB_Name = "Module1"
Sub countTickerSymbols()
    
    'Create a container to hold the number of records'
    'Avoid overflow by using LONG integer types'
    Dim recordQuantity2018, recordQuantity2019, recordQuantity2020 As Long

    'Manually verified record counts in worksheets'
    'Why are there different amounts of records?'
    'Are these records directly comparable?'
    recordQuantity2018 = 753000
    recordQuantity2019 = 756000
    recordQuantity2020 = 759000

    'Create for loop to count number of symbols
    'outer loop for each TLA symbol
        'inner loop for each individual TLA symbol of a set
        'end inner loop
    'end outer loop

    'Forloop to traverse the 2018 worksheet
    For iterator = 1 To recordQuantity2018 + 1
        If (Cells(iterator, 1).Value = "AAB") Then
            quantityOfIndividualSymbols = quantityOfIndividualSymbols + 1
        End If
    Next iterator
    
    MsgBox (quantityOfIndividualSymbols)

End Sub

