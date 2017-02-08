
'Утилита это называется q (чтобы было удобно вызывать, и намек на Query — запрос), принимает она в качестве параметров сам запрос (либо полностью в виде SQL, либо только название запроса или таблицы, главное, чтобы это можно было открыть через CurrentDB.OpenRecordset), максимальную ширину поля при выдаче (если поле больше — оно обрезается, по умолчанию — 10 символов) и максимальное количество записей (по умолчанию — 100 записей), и выдает в консоль содержимое результатов данного запроса в текстовом виде, выглядит это вот так:
'?q("qryStatBestVertriebler")
'Running qryStatBestVertriebler...
'Query returned 9 entries.
'-------------------------------------------------------------------------------------
'|  Nr |        VMB |      FName |    Einkauf |    Verkauf |     Gewinn |        Num | 
'-------------------------------------------------------------------------------------
'|    1|        757 | Sönke Doba |  492661,52 |   718774,8 |  226113,28 |        231 | 
'|    2|        877 | Johannes W |   10464,99 |   59677,25 |   49212,26 |         39 | 
'|    3|       1098 | Marco Müll |    8233,18 |   12244,77 |    4011,59 |         36 | 
'|    4|       5527 | Torben Jas |   12974,64 |   24642,42 |   11667,78 |         16 | 
'|    5|       6214 | Thiemo Wol |    5932,17 |   12175,97 |     6243,8 |         23 | 
'|    6|       7833 | Florian Mi |  207384,93 |  293553,82 | 86168,8900 |        254 | 
'|    7|       8310 | Daniel Sch |    3525,56 |    4338,92 |     813,36 |          4 | 
'|    8|       8917 | Daniela He |  187881,29 |  638726,06 |  450844,77 |        559 | 
'|    9|       9330 | Konrad Cyw |   94142,67 |  133056,71 |   38914,04 |        139 | 
'-------------------------------------------------------------------------------------


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
