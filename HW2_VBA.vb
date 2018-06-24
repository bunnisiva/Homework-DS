Sub test_summer()

   'declare the variable

   
   Dim my_variable As String
   
   Dim sum As Double
   sum = 0
 
   Dim Summary_Table_Row As Integer
   Summary_Table_Row = 2

 
   For i = 2 To 705714
   

   'column 1 is "ticker, column 7 is "Volume", column 8 is "ticker Name", column 9 is "ticker Sum"
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

           my_variable = Cells(i, 1).Value
           
           sum = sum + Cells(i, 7).Value
           
           Range("I" & Summary_Table_Row).Value = my_variable
           
           Range("J" & Summary_Table_Row).Value = sum
           
           Summary_Table_Row = Summary_Table_Row + 1
           
           sum = 0
           
           
       Else
       
           sum = sum + Cells(i, 7).Value
           
           
       End If
       
   Next i

End Sub
••••ˇˇˇˇ