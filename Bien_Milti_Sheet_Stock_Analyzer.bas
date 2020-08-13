Attribute VB_Name = "Module1"
Sub Mult_Sheet_Stock_Analyzer()

Dim lastrow As Long
Dim tickername As String
Dim summarytablerow As Integer
Dim startprice As Double
Dim endprice As Double
Dim volumetotal As Double
Dim yearlychange As Double
Dim percentchange As Double

Dim lastrowyearlychange As Long
Dim maxpercentincrease As Double
Dim maxpercentdecrease As Double
Dim maxtotalvolume As Double
Dim maxpercentincreaseticker As String
Dim maxpercentdecreaseticker As String
Dim maxtotalvolumeticker As String

Dim ws As Worksheet

'Loop through each worksheet
For Each ws In Worksheets

'Check worksheet name
'MsgBox (ws.Name)


'Set all variable to zero
    startprice = 0
    endprice = 0
    volumetotal = 0
    yearlychange = 0
    percentchange = 0
    lastrow = 0
    
'Set first startprice
startprice = ws.Cells(2, 6).Value

'Set first summary table row = 2
summarytablerow = 2

'Determne last row of column A
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Create Summary Table headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Loop through rows
    For i = 2 To lastrow
    
    'If <ticker> value changes
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

        'set the ticker name value
        tickername = ws.Cells(i, 1).Value
        
        'Add to volume total
        volumetotal = volumetotal + ws.Cells(i, 7).Value
        
        'Capture end price
            endprice = ws.Cells(i, 6).Value
            
        'MsgBox ("startprice " + Str(startprice) + "endprice " + Str(endprice))
            
            
        'Calculate yearly change
        yearlychange = startprice - endprice
        
        'Calculate percent change
            'check if start price = 0 to avoid divide by zero error
            If startprice = 0 Then
                percentchange = 0
            Else
            percentchange = ((startprice - endprice) / startprice) * 100
            
            End If
            
        'Write <ticker> value to summary table
            ws.Cells(summarytablerow, 9).Value = tickername
            
        'Write volume total
            ws.Cells(summarytablerow, 12).Value = volumetotal
            
        'Write yearly change to summary table
            ws.Cells(summarytablerow, 10) = yearlychange
        'Format yearly change
            If ws.Cells(summarytablerow, 10).Value < 0 Then
                ws.Cells(summarytablerow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(summarytablerow, 10).Interior.ColorIndex = 4
            End If
            
        'Write percent change to summary table
            ws.Cells(summarytablerow, 11).Value = Str(percentchange) + "%"

        'Increment a row in the summary table
        summarytablerow = summarytablerow + 1

        'Reset volumetotal
        volumetotal = 0
        
        'Set new start price
        startprice = ws.Cells(i + 1, 4).Value

    'If <ticker> value does not change
        Else
        
        'Increase volume total
         volumetotal = volumetotal + ws.Cells(i, 7).Value
        
    End If

    Next i
    
    '**************************** Evalulate Ticker/Yearly Change/Percent Change/Total Stock Volume Table ***********************************************
    
    'Write table labels
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    'Set all variable to zero
    
    lastrowyearlychange = 0
    
    'Determine last row of column I
    lastrowyearlychange = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Set initial values
        maxpercentincrease = 0
        maxpercentdecrease = 0
        maxtotalvolume = 0
        
For j = 2 To lastrowyearlychange
'MsgBox (lastrowyearlychange)
        
    'Determine Greatest % Increase
    If maxpercentincrease < ws.Cells(j, 11).Value Then
        maxpercentincreaseticker = ws.Cells(j, 9).Value
        maxpercentincrease = ws.Cells(j, 11).Value
    Else
    End If
    
    'MsgBox ("maxpercentincreaseticker =" + maxpercentincreaseticker + ", maxpercentincrease =" + Str(maxpercentincrease))
        
    'Determine Greatest % Decrease
    
    If maxpercentdecrease > ws.Cells(j, 11).Value Then
        maxpercentdecreaseticker = ws.Cells(j, 9).Value
        maxpercentdecrease = ws.Cells(j, 11).Value
    End If
    
    'MsgBox ("maxpercentdecreaseticker =" + maxpercentdecreaseticker + ", maxpercentdecrease =" + Str(maxpercentdecrease))
    
    'Determine Max Volume
    
    If maxtotalvolume < ws.Cells(j, 12).Value Then
        maxtotalvolumeticker = ws.Cells(j, 9).Value
        maxtotalvolume = ws.Cells(j, 12).Value
    Else
    End If
    
    'MsgBox ("maxtotalvolumeticker =" + maxtotalvolumeticker + ", maxtotalvolume =" + Str(maxtotalvolume))

    Next j
   
   'Write vaulue to table
    ws.Cells(2, 15).Value = maxpercentincreaseticker
    ws.Cells(2, 16).Value = Str(maxpercentincrease * 100) + "%"
    ws.Cells(3, 15).Value = maxpercentdecreaseticker
    ws.Cells(3, 16).Value = Str(maxpercentdecrease * 100) + "%"
    ws.Cells(4, 15).Value = maxtotalvolumeticker
    ws.Cells(4, 16).Value = maxtotalvolume
    
    
    Next ws
'whyerror:
'MsgBox ("startprice = " + Str(startprice) + " endprice = " + Str(endprice))

End Sub
