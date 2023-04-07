Sub stockcounter() 'Assignment #2 - Steve Bennett

    Dim rowCount As Long                ' how many rows
    Dim howmanyrows As Long             ' automated value to know the upper limit size for Range area size
    Dim openamount As Double            ' open amount for the stock
    Dim closeamount As Double           ' closing amount at end of year
    Dim volamount As Double             ' total volume through the year
    Dim tickername As String            ' name of the stock
    Dim isfirst As Integer              ' controls my initial IF THEN statement
    Dim namecheck As String             ' checks the upcoming name to compare against tickername
    Dim yearlychange As Double          ' YOY value
    Dim percentchange As Double         ' YOY % value
    Dim outputrow As Long               ' formatting placeholder
    
    Dim currentsheet() As String        'my array to store the sheet names
    Dim sheetcycle As Integer           'the counter for the array
    
    Dim bestinc As Double               'tracker for Greatest % Increase
    Dim bestdec As Double               'tracker for Greatest % Decrease
    Dim bestvol As Double               'tracker for Greatest total volume
    Dim bestincname As String           'storing the names for the above 3
    Dim bestdecname As String
    Dim bestvolname As String
    
    ReDim currentsheet(ActiveWorkbook.Sheets.Count) 'sets the array size to match the qty of sheets
    
    'START OF MY PRIMARY LOOP CYLCING THROUGH EACH SHEET
    For sheetcycle = 1 To ActiveWorkbook.Sheets.Count 'runs the primary outer loop for how many sheets we have
        currentsheet(sheetcycle) = ActiveWorkbook.Sheets(sheetcycle).Name 'sets the name for the sheet to the array string
    
        Worksheets(currentsheet(sheetcycle)).Cells.ClearFormats  'makes sure we're starting fresh no formatting in the way
        Worksheets(currentsheet(sheetcycle)).Columns("I:Q").Delete 'also clears old data so there is no confusion
    
        howmanyrows = (Worksheets(currentsheet(sheetcycle)).UsedRange.Rows.Count - 1) ' figures out how many vertical rows it has and removes the header
    
        outputrow = 2 'this ensures my outputs are starting below the header
        isfirst = 1 'makes sure I enter my first if then loop
    
        Worksheets(currentsheet(sheetcycle)).Cells(1, 9).Value = "Ticker" 'adds the new column name for the output values
        Worksheets(currentsheet(sheetcycle)).Cells(1, 10).Value = "Yearly Change" 'adds the new column name for the output values
        Worksheets(currentsheet(sheetcycle)).Cells(1, 11).Value = "Percent Change" 'adds the new column name for the output values
        Worksheets(currentsheet(sheetcycle)).Cells(1, 12).Value = "Total Stock Volume" ' adds the new column name for the output values
        Worksheets(currentsheet(sheetcycle)).Cells(1, 16).Value = "Ticker" ' adds the new column name for the output values
        Worksheets(currentsheet(sheetcycle)).Cells(1, 17).Value = "Value" ' adds the new column name for the output values
        Worksheets(currentsheet(sheetcycle)).Cells(2, 15).Value = "Greatest % Increase" ' adds the new column name for the output values
        Worksheets(currentsheet(sheetcycle)).Cells(3, 15).Value = "Greatest % Decrease" ' adds the new column name for the output values
        Worksheets(currentsheet(sheetcycle)).Cells(4, 15).Value = "Greatest Total Volume" ' adds the new column name for the output values
    
        'START OF MY SECONDARY LOOP - CYCLES WITHIN EVERY ROW ON THE CURRENT SHEET
        For rowCount = 1 To howmanyrows 'this loops for how many rows DOWN we have
    
            If isfirst = 1 Then 'If its the first time we've seen this stock name, we store its name and open amount
                tickername = Worksheets(currentsheet(sheetcycle)).Range("A" & (rowCount + 1)).Value
                openamount = Worksheets(currentsheet(sheetcycle)).Range("C" & (rowCount + 1)).Value
                isfirst = 0
            End If
    
            volamount = Worksheets(currentsheet(sheetcycle)).Range("G" & (rowCount + 1)).Value + volamount 'start adding the volume amount for that date
    
            namecheck = Worksheets(currentsheet(sheetcycle)).Range("A" & rowCount + 2).Value 'store the string name for the next  stock on the row below
        
            If StrComp(tickername, namecheck, 1) <> 0 Then ' checking if the current stock name matches the next one, if not we output the results
                closeamount = Range("F" & (rowCount + 1)).Value 'grabs the final end of year value for where the stock closed
                yearlychange = closeamount - openamount ' calculates the year over year
                percentchange = yearlychange / openamount 'calculates the percentage change YOY

                Worksheets(currentsheet(sheetcycle)).Cells(outputrow, 9).Value = tickername 'adds the stock name under Ticker
                Worksheets(currentsheet(sheetcycle)).Cells(outputrow, 10).Value = yearlychange 'adds the YOY value under Yearly Change
            
                If yearlychange > 0 Then  'Color conditional formatting for the YOY if a loss = RED, gain = GREEN
                    Worksheets(currentsheet(sheetcycle)).Cells(outputrow, 10).Interior.ColorIndex = 4
                ElseIf yearlychange < 0 Then
                    Worksheets(currentsheet(sheetcycle)).Cells(outputrow, 10).Interior.ColorIndex = 3
                End If
            
                Worksheets(currentsheet(sheetcycle)).Cells(outputrow, 11).Value = percentchange 'adds the percent YOY change value
                Worksheets(currentsheet(sheetcycle)).Cells(outputrow, 11).NumberFormat = "0.00%" 'conditional formats to % value
                Worksheets(currentsheet(sheetcycle)).Cells(outputrow, 12).Value = volamount 'finally adds the total stock volume for the year

                'THIS SECTION CHECKS FOR THE GREATEST VALUES  AND STORES ONLY IF THE VALUE IS GREATEST
                If volamount > bestvol Then
                    bestvol = volamount
                    bestvolname = tickername
                End If
            
                If percentchange > bestinc Then
                    bestinc = percentchange
                    bestincname = tickername
                ElseIf percentchange < bestdec Then
                    bestdec = percentchange
                    bestdecname = tickername
                End If
            
                outputrow = outputrow + 1 'sends it to the next row for when we do our next stock calculations
                volamount = 0 'resets the volume amount for the new stock
                isfirst = 1 'resets the initial IF THEN loop so we start fresh
        
            End If
        
        Next rowCount

        'OUTPUTS THE GREATEST VALUES SECTION OF THE DOCUMENT
        Worksheets(currentsheet(sheetcycle)).Cells(2, 16).Value = bestincname
        Worksheets(currentsheet(sheetcycle)).Cells(2, 17).NumberFormat = "0.00%" 'conditional formats to % value
        Worksheets(currentsheet(sheetcycle)).Cells(2, 17).Value = bestinc
        Worksheets(currentsheet(sheetcycle)).Cells(3, 16).Value = bestdecname
        Worksheets(currentsheet(sheetcycle)).Cells(3, 17).NumberFormat = "0.00%" 'conditional formats to % value
        Worksheets(currentsheet(sheetcycle)).Cells(3, 17).Value = bestdec
        Worksheets(currentsheet(sheetcycle)).Cells(4, 16).Value = bestvolname
        Worksheets(currentsheet(sheetcycle)).Cells(4, 17).Value = bestvol
        Worksheets(currentsheet(sheetcycle)).Range("O:O").ColumnWidth = 21 'formats the width to match the row text widtch
        Worksheets(currentsheet(sheetcycle)).Range("Q:Q").ColumnWidth = 11.3 'formats the width to match the row text widtch
        Worksheets(currentsheet(sheetcycle)).Range("I1:P1").Columns.AutoFit 'auto formats the width to match the header text widtch
            
        'BEGIN PURGE OF DATA SO NOTHING CARRIES OVER INTO THE NEXT SHEET LOOP CALCULATIONS
        bestvol = Empty
        bestinc = Empty
        bestdec = Empty
        bestvolname = ""
        bestinvname = ""
        bestdecname = ""
        volamount = Empty
        percentchange = Empty
        openamount = Empty
        tickername = ""
        nextname = ""
        'PURGE COMPLETE - START NEXT SHEET LOOP
            
    Next sheetcycle

End Sub
