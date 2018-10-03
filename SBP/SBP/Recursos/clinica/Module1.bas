Attribute VB_Name = "Module1"
 Public rs As New ADODB.Recordset
  Public cn As New ADODB.Connection
 Public gsede1 As String
 Public ngsede1 As String
 Public opcion1 As Integer
 Public dgusuario As String
 Public globaldat As String
 

 Public Function conectar()
 On Error GoTo cmd1_error
 cn.CursorLocation = adUseClient
 cn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=osi;Data Source=(local)"
 conectar = 1
 Exit Function
cmd1_error:
 Exit Function
 End Function
 Function busca_combox(xxyzcontrol As Control, buf As String)
On Error GoTo cmd45_err
Dim i As Integer
Dim sw As Integer
sw = 0
For i = 0 To xxyzcontrol.ListCount - 1
   If xxyzcontrol.List(i) = buf Then
      busca_combox = i
      sw = 1
      Exit For
   End If
Next i
If sw = 0 Then
   busca_combox = 0
End If
Exit Function
cmd45_err:
busca_combox = 0
Exit Function
End Function
Function extra_loquesea(buf As String) As String
Dim j
Dim buf1 As String
buf1 = ""
If InStr(buf, "|") > 0 Then
   j = InStr(buf, "|")
   buf1 = Mid$(buf, 1, j - 1)
   Else
   buf1 = buf
End If
extra_loquesea = buf1
End Function


