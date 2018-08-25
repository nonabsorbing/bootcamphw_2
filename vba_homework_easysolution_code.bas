Attribute VB_Name = "Module1"
Sub stockticker():

'what am i trying to do here?
'start process with cycling thru first column,
'add name in first column to summary table
'add counter to summary table,
'keep adding to counter until new company comes up, and then start a new summary line

'pulled from ticker
    Dim company_name As String

' company volume - start at 0
    Dim company_volume As Long
        company_volume = 0
    
'use to build new rows for summary table
    Dim summary_row As Long
    summary_row = 1
    
'for variables - for columns and rows. use long because they'll be big numbers

    Dim i As Long
    
'   Dim lastrow As Long
          
        lastrow = Range("A" & Rows.Count).End(xlUp).Row

'define for loop from 2 to the last row in the first column
 
For i = 2 To lastrow

'compare ticker cell to ticker cell below it

    If Range("A" & (i + 1)).Value <> Range("A" & i).Value Then
        summary_row = (summary_row + 1)
        company_name = Range("A" & i).Value
         
          'if cell doesn't match, add a new line to the summary table
          'and start adding to the volume counter - and divide by 1000 to avoid overflow
            
            company_volume = company_volume + ((Range("G" & i).Value) / 1000)
            
       Range("I" & summary_row).Value = company_name
       Range("J" & summary_row).Value = company_volume
      
       company_volume = 0

    'if cell matches, keep cycling through and adding to volume counter, divide by 1000

    Else: company_volume = company_volume + Range("G" & i).Value / 1000
     
            
    End If
    
Next i


End Sub


