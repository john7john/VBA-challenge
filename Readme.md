


Sub final():
  For Each ws In Worksheets							' Looping through worksheets'								
 
 
 Dim lr As Long
 
 ws.Range("I1").Value = "Ticker"						' Giving titles to cells'
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change($)"
 ws.Range("l1").Value = "Total Stock Volume"
 ws.Range("o3").Value = "Greatest % Increase"
 ws.Range("o4").Value = "Greatest % Decrease"
 ws.Range("o5").Value = "Greatest Total Volume"
 ws.Range("p2").Value = "Ticker"
ws.Range("q2").Value = "Value"
 ws.Range("K2:K1000").NumberFormat = "0.00%"
 
 lr = ws.Cells(Rows.Count, 1).End(xlUp).Row        			' length of rows with data'

 j = 1
 Start = 2
 For i = 2 To lr
 closing1 = 0
 
 
 If ws.Cells(i, 1).Value <> ws.Cells(j, 9).Value Then  		'Finding Ticker'
  j = j + 1
  t = 0
 ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
 End If

 
If ws.Cells(i, 1).Value = ws.Cells(j, 9).Value Then  			 'Finding Total -   Values are getting addes to 't' while running through each cells in same ticeker'
t = t + ws.Cells(i, 7).Value
ws.Cells(j, 12).Value = t

End If



If ws.Cells(i, 1).Value = ws.Cells(j, 9).Value And ws.Cells(i, 2).Value = ws.Cells(2, 2).Value Then

Start = ws.Cells(i, 3).Value                       			 'finding start value of stock'

End If

If ws.Cells(i, 1).Value = ws.Cells(j, 9).Value Then
    
   closing1 = ws.Cells(i, 6).Value            				 'finding closing value of stock at the end of the year'

End If

ws.Cells(j, 10).Value = closing1 - Start          			  'Percentage change and conditional formating'
ws.Cells(j, 11).Value = (closing1 - Start) / Start
If ws.Cells(j, 11).Value >= 0 Then
ws.Cells(j, 11).Interior.ColorIndex = 4
Else
ws.Cells(j, 11).Interior.ColorIndex = 3
End If


If Cells(j, 10).Value >= 0 Then                				 'Yearly change conditional formating'
ws.Cells(j, 10).Interior.ColorIndex = 4
Else
ws.Cells(j, 10).Interior.ColorIndex = 3
End If


Next i
 
 grt1 = ws.Cells(2, 11).Value
lst1 = ws.Cells(2, 11).Value
tots1 = ws.Cells(2, 12).Value



 lr = ws.Cells(Rows.Count, 11).End(xlUp).Row
For i = 2 To lr
If ws.Cells(i, 11).Value > grt1 Then                			'  Finding Greatest % Increase,% Decrease and Total'
grt1 = ws.Cells(i, 11).Value
grtTicker = i
End If

If ws.Cells(i, 11).Value < lst1 Then

lst1 = ws.Cells(i, 11).Value
lstTicker = i

End If
If ws.Cells(i, 12).Value > tots1 Then
tots1 = ws.Cells(i, 12).Value
TotsTicker = i

End If
          
  
 Next i
 
 ws.Cells(3, 17).Value = grt1               				' Assigning values to cells'
 ws.Cells(3, 17).NumberFormat = "0.00%"
  ws.Cells(4, 17).Value = lst1
  ws.Cells(4, 17).NumberFormat = "0.00%"
  ws.Cells(5, 17).Value = tots1
 ws.Cells(3, 16).Value = ws.Cells(grtTicker, 9).Value
 ws.Cells(4, 16).Value = ws.Cells(lstTicker, 9).Value
 ws.Cells(5, 16).Value = ws.Cells(TotsTicker, 9).Value

 Next ws
 
 MsgBox ("processed")

 End Sub

 
Sub clear():
For Each ws In Worksheets

ws.Range("I1:I25000").ClearContents						'clearing all resutls  in all worksheets'   ' Click Process button to run the script' 
ws.Range("j1:j25000").ClearContents
ws.Range("l1:l25000").ClearContents
ws.Range("k1:k25000").ClearContents
ws.Range("k2:k25000").Interior.ColorIndex = 2
ws.Range("j2:k25000").Interior.ColorIndex = 2
ws.Range("o2:o10").ClearContents
ws.Range("q2:q10").ClearContents
 ws.Range("p2:p10").ClearContents

Next ws
End Sub

