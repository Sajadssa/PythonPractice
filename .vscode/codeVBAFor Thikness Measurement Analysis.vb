Private Sub cbpoint_AfterUpdate()

  '----graph4
On Error Resume Next


'----graph 4============================================
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim pointNo As String
    Dim Location As String
    Dim Position As String
    Dim ReportDate As Date
    Dim MRT As Double
    Dim prevValue As Double
    Dim currentValue As Variant ' Use Variant to handle Null
    Dim lastReportDate As Date
    
    Set db = CurrentDb
    

    pointNo = Trim(Me.cbpoint)
    Location = Trim(Me.cbloc)
    Position = "0/12" 
    

    On Error Resume Next
    

    strSQL = "SELECT Report_Date, WT_Measurement, NWTBase, CA FROM Cor_qrThkMeasurement " & _
             "WHERE Point_No = '" & pointNo & "' AND Location = '" & Location & "' AND Position = '" & Position & "' " & _
             "ORDER BY Report_Date ASC"
    
    Set rst = db.OpenRecordset(strSQL, dbOpenDynaset)
    
    If rst.EOF Then
        MsgBox "No data input", vbExclamation
        Exit Sub
    End If
    

    rst.MoveFirst
    ReportDate = rst!Report_Date

    If IsNull(rst!WT_Measurement) Then
        calculatedValue = Nz(rst!NWTBase, 0)
    Else
        calculatedValue = rst!WT_Measurement
    End If
    prevValue = calculatedValue
  
    MRT = Nz(rst!NWTBase, 0) - Nz(rst!CA, 0)

    db.Execute "INSERT INTO TempTable_ThkMeasurement (Report_Date, Point_No, Location, Position, WT_Measurement, MRT) " & _
               "VALUES (#" & Format(ReportDate, "mm\/dd\/yyyy") & "#, '" & pointNo & "', '" & Location & "', '" & Position & "', " & calculatedValue & ", " & MRT & ")"
    
    rst.MoveNext 
    

    Do While Not rst.EOF
        ReportDate = rst!Report_Date
        currentValue = rst!WT_Measurement
       
        If IsNull(currentValue) Then
            calculatedValue = prevValue
        Else
            
            If currentValue > prevValue Then
                calculatedValue = prevValue
            Else
                calculatedValue = currentValue
            End If
        End If
        
      
        MRT = Nz(rst!NWTBase, 0) - Nz(rst!CA, 0)
        
     
        db.Execute "INSERT INTO TempTable_ThkMeasurement (Report_Date, Point_No, Location, Position, WT_Measurement, MRT) " & _
                   "VALUES (#" & Format(ReportDate, "mm\/dd\/yyyy") & "#, '" & pointNo & "', '" & Location & "', '" & Position & "', " & _
                   Nz(calculatedValue, "NULL") & ", " & MRT & ")"
        
  
        prevValue = calculatedValue
        lastReportDate = ReportDate 
        
        rst.MoveNext
    Loop

    If lastReportDate <> 0 Then
        strSQL = "UPDATE TempTable_ThkMeasurement " & _
                 "SET WT_Measurement = " & prevValue & " " & _
                 "WHERE Report_Date = #" & Format(lastReportDate, "mm\/dd\/yyyy") & "# " & _
                 "AND Point_No = '" & pointNo & "' " & _
                 "AND Location = '" & Location & "' " & _
                 "AND Position = '" & Position & "' " & _
                 "AND WT_Measurement IS NULL;"
        db.Execute strSQL
    End If
    
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    

    Me.Graph4.RowSource = "SELECT Report_Date, WT_Measurement, MRT FROM TempTable_ThkMeasurement " & _
                          "WHERE Point_No = '" & pointNo & "' AND Location = '" & Location & "' AND Position = '" & Position & "' " & _
                          "ORDER BY Report_Date ASC;"
    Me.Graph4.Requery
    

'----graph 5============================================

    Dim rst5 As DAO.Recordset
    Dim strSQL5 As String
    Dim pointNo5 As String
    Dim Location5 As String
    Dim position5 As String
    Dim reportDate5 As Date
    Dim MRT5 As Double
    Dim prevValue5 As Double
    Dim currentValue5 As Variant ' Use Variant to handle Null
    Dim lastReportDate5 As Date
    
    Set db = CurrentDb
    
    pointNo5 = Trim(Me.cbpoint)
    Location5 = Trim(Me.cbloc)
    position5 = "3" 
    

    On Error Resume Next
    

    strSQL5 = "SELECT Report_Date, WT_Measurement, NWTBase, CA FROM Cor_qrThkMeasurement " & _
             "WHERE Point_No = '" & pointNo5 & "' AND Location = '" & Location5 & "' AND Position = '" & position5 & "' " & _
             "ORDER BY Report_Date ASC"
    
    Set rst5 = db.OpenRecordset(strSQL5, dbOpenDynaset)
    
    If rst5.EOF Then
        MsgBox "No data input", vbExclamation
        Exit Sub
    End If
    

    rst5.MoveFirst
    reportDate5 = rst5!Report_Date
    
    If IsNull(rst5!WT_Measurement) Then
        calculatedValue5 = Nz(rst5!NWTBase, 0)
    Else
        calculatedValue5 = rst5!WT_Measurement
    End If
    prevValue5 = calculatedValue5

    MRT5 = Nz(rst5!NWTBase, 0) - Nz(rst5!CA, 0)
    

    db.Execute "INSERT INTO TempTable_ThkMeasurement (Report_Date, Point_No, Location, Position, WT_Measurement, MRT) " & _
               "VALUES (#" & Format(reportDate5, "mm\/dd\/yyyy") & "#, '" & pointNo5 & "', '" & Location5 & "', '" & position5 & "', " & calculatedValue5 & ", " & MRT5 & ")"
    
    rst5.MoveNext
    
   
    Do While Not rst5.EOF
        reportDate5 = rst5!Report_Date
        currentValue5 = rst5!WT_Measurement
        
     
        If IsNull(currentValue5) Then
            calculatedValue5 = prevValue5
        Else
           
            If currentValue5 > prevValue5 Then
                calculatedValue5 = prevValue5
            Else
                calculatedValue5 = currentValue5
            End If
        End If
        

        MRT5 = Nz(rst5!NWTBase, 0) - Nz(rst5!CA, 0)
        

        db.Execute "INSERT INTO TempTable_ThkMeasurement (Report_Date, Point_No, Location, Position, WT_Measurement, MRT) " & _
                   "VALUES (#" & Format(reportDate5, "mm\/dd\/yyyy") & "#, '" & pointNo5 & "', '" & Location5 & "', '" & position5 & "', " & _
                   Nz(calculatedValue5, "NULL") & ", " & MRT5 & ")"
        
     
        prevValue5 = calculatedValue5
        lastReportDate5 = reportDate5 
        
        rst5.MoveNext
    Loop
    

    If lastReportDate5 <> 0 Then
        strSQL5 = "UPDATE TempTable_ThkMeasurement " & _
                 "SET WT_Measurement = " & prevValue5 & " " & _
                 "WHERE Report_Date = #" & Format(lastReportDate5, "mm\/dd\/yyyy") & "# " & _
                 "AND Point_No = '" & pointNo5 & "' " & _
                 "AND Location = '" & Location5 & "' " & _
                 "AND Position = '" & position5 & "' " & _
                 "AND WT_Measurement IS NULL;"
        db.Execute strSQL5
    End If
    
    rst5.Close
    Set rst5 = Nothing
    Set db = Nothing
    
  
    Me.Graph5.RowSource = "SELECT Report_Date, WT_Measurement, MRT FROM TempTable_ThkMeasurement " & _
                          "WHERE Point_No = '" & pointNo5 & "' AND Location = '" & Location5 & "' AND Position = '" & position5 & "' " & _
                          "ORDER BY Report_Date ASC;"
    Me.Graph5.Requery
    

'----graph 6============================================
   
    Dim rst6 As DAO.Recordset
    Dim strSQL6 As String
    Dim pointNo6 As String
    Dim Location6 As String
    Dim position6 As String
    Dim reportDate6 As Date
    Dim MRT6 As Double
    Dim prevValue6 As Double
    Dim currentValue6 As Variant ' Use Variant to handle Null
    Dim lastReportDate6 As Date
    
    Set db = CurrentDb

    pointNo6 = Trim(Me.cbpoint)
    Location6 = Trim(Me.cbloc)
    position6 = "6" ' ?? ????? ???? ?? ????? ???? ???
    
 
    On Error Resume Next
    
    
 
    strSQL6 = "SELECT Report_Date, WT_Measurement, NWTBase, CA FROM Cor_qrThkMeasurement " & _
             "WHERE Point_No = '" & pointNo6 & "' AND Location = '" & Location6 & "' AND Position = '" & position6 & "' " & _
             "ORDER BY Report_Date ASC"
    
    Set rst6 = db.OpenRecordset(strSQL6, dbOpenDynaset)
    
    If rst6.EOF Then
        MsgBox "No data input", vbExclamation
        Exit Sub
    End If
    

    rst6.MoveFirst
    reportDate6 = rst6!Report_Date
 
    If IsNull(rst6!WT_Measurement) Then
        calculatedValue6 = Nz(rst6!NWTBase, 0)
    Else
        calculatedValue6 = rst6!WT_Measurement
    End If
    prevValue6 = calculatedValue6

    MRT6 = Nz(rst6!NWTBase, 0) - Nz(rst6!CA, 0)

    db.Execute "INSERT INTO TempTable_ThkMeasurement (Report_Date, Point_No, Location, Position, WT_Measurement, MRT) " & _
               "VALUES (#" & Format(reportDate6, "mm\/dd\/yyyy") & "#, '" & pointNo6 & "', '" & Location6 & "', '" & position6 & "', " & calculatedValue6 & ", " & MRT6 & ")"
    
    rst6.MoveNext 
    

    Do While Not rst6.EOF
        reportDate6 = rst6!Report_Date
        currentValue6 = rst6!WT_Measurement
        
       
        If IsNull(currentValue6) Then
            calculatedValue6 = prevValue6
        Else
         
            If currentValue6 > prevValue6 Then
                calculatedValue6 = prevValue6
            Else
                calculatedValue6 = currentValue6
            End If
        End If
        
      
        MRT6 = Nz(rst6!NWTBase, 0) - Nz(rst6!CA, 0)
        
  
        db.Execute "INSERT INTO TempTable_ThkMeasurement (Report_Date, Point_No, Location, Position, WT_Measurement, MRT) " & _
                   "VALUES (#" & Format(reportDate6, "mm\/dd\/yyyy") & "#, '" & pointNo6 & "', '" & Location6 & "', '" & position6 & "', " & _
                   Nz(calculatedValue6, "NULL") & ", " & MRT6 & ")"
        
     
        prevValue6 = calculatedValue6
        lastReportDate6 = reportDate6 
        
        rst6.MoveNext
    Loop
  
    If lastReportDate6 <> 0 Then
        strSQL6 = "UPDATE TempTable_ThkMeasurement " & _
                 "SET WT_Measurement = " & prevValue6 & " " & _
                 "WHERE Report_Date = #" & Format(lastReportDate6, "mm\/dd\/yyyy") & "# " & _
                 "AND Point_No = '" & pointNo6 & "' " & _
                 "AND Location = '" & Location6 & "' " & _
                 "AND Position = '" & position6 & "' " & _
                 "AND WT_Measurement IS NULL;"
        db.Execute strSQL6
    End If
    
    rst6.Close
    Set rst6 = Nothing
    Set db = Nothing
    
 
    Me.Graph6.RowSource = "SELECT Report_Date, WT_Measurement, MRT FROM TempTable_ThkMeasurement " & _
                          "WHERE Point_No = '" & pointNo6 & "' AND Location = '" & Location6 & "' AND Position = '" & position6 & "' " & _
                          "ORDER BY Report_Date ASC;"
    Me.Graph6.Requery
    
   '----graph 7============================================
   
    Dim rst7 As DAO.Recordset
    Dim strSQL7 As String
    Dim pointNo7 As String
    Dim Location7 As String
    Dim position7 As String
    Dim reportDate7 As Date
    Dim MRT7 As Double
    Dim prevValue7 As Double
    Dim currentValue7 As Variant ' Use Variant to handle Null
    Dim lastReportDate7 As Date
    
    Set db = CurrentDb

    pointNo7 = Trim(Me.cbpoint)
    Location7 = Trim(Me.cbloc)
    position7 = "9" 
    
   
    On Error Resume Next

    strSQL7 = "SELECT Report_Date, WT_Measurement, NWTBase, CA FROM Cor_qrThkMeasurement " & _
             "WHERE Point_No = '" & pointNo7 & "' AND Location = '" & Location7 & "' AND Position = '" & position7 & "' " & _
             "ORDER BY Report_Date ASC"
    
    Set rst7 = db.OpenRecordset(strSQL7, dbOpenDynaset)
    
    If rst7.EOF Then
        MsgBox "No data input", vbExclamation
        Exit Sub
    End If
    
    rst7.MoveFirst
    reportDate7 = rst7!Report_Date
   
    If IsNull(rst7!WT_Measurement) Then
        calculatedValue7 = Nz(rst7!NWTBase, 0)
    Else
        calculatedValue7 = rst7!WT_Measurement
    End If
    prevValue7 = calculatedValue7

    MRT7 = Nz(rst7!NWTBase, 0) - Nz(rst7!CA, 0)

    db.Execute "INSERT INTO TempTable_ThkMeasurement (Report_Date, Point_No, Location, Position, WT_Measurement, MRT) " & _
               "VALUES (#" & Format(reportDate7, "mm\/dd\/yyyy") & "#, '" & pointNo7 & "', '" & Location7 & "', '" & position7 & "', " & calculatedValue7 & ", " & MRT7 & ")"
    
    rst7.MoveNext 
    
  
    Do While Not rst7.EOF
        reportDate7 = rst7!Report_Date
        currentValue7 = rst7!WT_Measurement
        
    
        If IsNull(currentValue7) Then
            calculatedValue7 = prevValue7
        Else
            
            If currentValue7 > prevValue7 Then
                calculatedValue7 = prevValue7
            Else
                calculatedValue7 = currentValue7
            End If
        End If
  
        MRT7 = Nz(rst7!NWTBase, 0) - Nz(rst7!CA, 0)
        

        db.Execute "INSERT INTO TempTable_ThkMeasurement (Report_Date, Point_No, Location, Position, WT_Measurement, MRT) " & _
                   "VALUES (#" & Format(reportDate7, "mm\/dd\/yyyy") & "#, '" & pointNo7 & "', '" & Location7 & "', '" & position7 & "', " & _
                   Nz(calculatedValue7, "NULL") & ", " & MRT7 & ")"

        prevValue7 = calculatedValue7
        lastReportDate7 = reportDate7 
        
        rst7.MoveNext
    Loop

    If lastReportDate7 <> 0 Then
        strSQL7 = "UPDATE TempTable_ThkMeasurement " & _
                 "SET WT_Measurement = " & prevValue7 & " " & _
                 "WHERE Report_Date = #" & Format(lastReportDate7, "mm\/dd\/yyyy") & "# " & _
                 "AND Point_No = '" & pointNo7 & "' " & _
                 "AND Location = '" & Location7 & "' " & _
                 "AND Position = '" & position7 & "' " & _
                 "AND WT_Measurement IS NULL;"
        db.Execute strSQL7
    End If
    
    rst7.Close
    Set rst7 = Nothing
    Set db = Nothing
  
    Me.Graph7.RowSource = "SELECT Report_Date, WT_Measurement, MRT FROM TempTable_ThkMeasurement " & _
                          "WHERE Point_No = '" & pointNo7 & "' AND Location = '" & Location7 & "' AND Position = '" & position7 & "' " & _
                          "ORDER BY Report_Date ASC;"
    Me.Graph7.Requery

End Sub