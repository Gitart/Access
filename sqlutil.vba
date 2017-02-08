


Public Function q(Optional strSQL As String = "", Optional intWidth As Integer = 10, Optional intMax As Integer = 100) As Boolean
 
  Dim tmpRCDSet As Recordset, tmpFeld As Field, tmpString As String, I As Integer, intTemplen As Integer
  Dim intNr As Integer
  
  On Error GoTo Err_SQL
  Debug.Print «Running „ & strSQL & “...»
  Set tmpRCDSet = CurrentDb.OpenRecordset(strSQL)
    tmpRCDSet.MoveLast
    Debug.Print «Query returned „ & tmpRCDSet.RecordCount & “ entries.»
    tmpRCDSet.MoveFirst
    tmpString = "| Nr | "
    For Each tmpFeld In tmpRCDSet.Fields
      tmpString = tmpString & padleft(tmpFeld.Name, intWidth) & " | "
    Next
    
    Debug.Print String(Len(tmpString) — 1, "-")
    Debug.Print tmpString
    Debug.Print String(Len(tmpString) — 1, "-")
    
    intNr = 1
    While (Not (tmpRCDSet.EOF)) And (intNr <= intMax)
      tmpString = "| " & padleft(Str(intNr), 4) & "| "
      For Each tmpFeld In tmpRCDSet.Fields
        tmpString = tmpString & padleft(Nz(tmpFeld.Value, ""), intWidth) & " | "
      Next
      Debug.Print tmpString
      intNr = intNr + 1
      tmpRCDSet.MoveNext
    Wend
    Debug.Print String(Len(tmpString) — 1, "-")
  
  Exit Function
 
 
Err_SQL:
  Debug.Print Err.Number & " " & Err.Description
  Debug.Print «Bad SQL string»
 
 
End Function
 
Function padleft(strLineIn As String, intWidth As Integer) As String
 
If Len(strLineIn) = intWidth Then
  padleft = strLineIn
ElseIf Len(strLineIn) > intWidth Then
  padleft = Mid(strLineIn, 1, intWidth)
Else
  padleft = String(intWidth — Len(strLineIn), " ") & strLineIn
End If
 
End Function
