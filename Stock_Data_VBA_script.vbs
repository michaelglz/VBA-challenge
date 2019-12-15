Sub Stock_Data()

'Preparing for looping through worksheets
Dim a As Integer
Dim ws_num As Integer
Dim starting_ws As Worksheet
Dim totalVolume As Variant

'Setting the worksheet count
ws_num = ThisWorkbook.Worksheets.Count

'Loop through worksheet
For a = 1 To ws_num

'Has it work only on the current workbook
ThisWorkbook.Worksheets(a).Activate

'Declaring variables
Dim firstOpen As Variant
Dim lastClose As Variant
Dim I As Double
Dim x As Double

'Headings
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

x = 2
I = 2

'Ticker
Cells(x, 9).Value = Cells(x, 1).Value

'First open of year
firstOpen = Cells(I, 3).Value

LastRow = Cells(Rows.Count, 1).End(xlDown).Row

For I = 2 To LastRow

        
        If Cells(I, 1).Value = Cells(x, 9).Value Then
        
                totalVolume = totalVolume + Cells(I, 7).Value
        
                lastClose = Cells(I, 6).Value
        
             Else
             
                Cells(x, 10).Value = lastClose - firstOpen
        
                    If lastClose <= 0 Or firstOpen <= 0 Then
                            
                            Cells(x, 11).Value = 0
                                    
                        Else
                        'Percent Change
                            Cells(x, 11).Value = (lastClose / firstOpen) - 1
                    
                    End If
        
                'Setting Cells(x, 11) as Percent type
                Cells(x, 11).Style = "Percent"
                
                'Conditional formatting for positive and negative change
                If Cells(x, 10).Value >= 0 Then
                                            
                            Cells(x, 10).Interior.ColorIndex = 4
                                                
                                Else
                                            
                            Cells(x, 10).Interior.ColorIndex = 3
                                
                End If
                
                'Reset variables
                Cells(x, 12).Value = totalVolume
        
                firstOpen = Cells(I, 3).Value
        
                totalVolume = Cells(I, 7).Value
                x = x + 1
                Cells(x, 9).Value = Cells(I, 1).Value

        End If

Next I

'Resizing columns
Columns("I:Q").EntireColumn.AutoFit
Cells(1, 1).Select

        Next a
        
End Sub

