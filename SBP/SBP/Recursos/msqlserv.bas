Attribute VB_Name = "Module8"
Option Explicit

Public cn        As New ADODB.Connection

Public cn1       As New ADODB.Connection

Public basedatos As String

Global mytable11 As New ADODB.Recordset

Global mytable12 As New ADODB.Recordset

Public Function conectar(buf As String)

    Dim dbuser     As String

    Dim dbpassword As String

    Dim dbname     As String

    Dim dbserver   As String

    Dim xservidor  As String

    basedatos = buf

    On Error GoTo cmd1_error

    Set cn = Nothing
    'MsgBox "hola0"
    cn.CursorLocation = adUseClient
    'MsgBox "hola1"
    cn.CommandTimeout = 0 '200000 '1024
    'MsgBox "Hola2"
    'cn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & buf & ";Data Source=(local)"
    'cn.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=orion;Uid=sa"
    'cn.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=" & buf & " ;Uid=sa;Pwd="
    'If cn.State = 1 Then
    '   cn.Close
    'End If
    'xservidor = "" & menup.vservidor
    'MsgBox xservidor & " " & buf & " " & clave_servidor
    cn.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=" & basedatos & " ;Uid=sa;pwd=" & clave_servidor & ""
    'cn.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=" & buf & " ;Uid=sa;pwd=;"
 
    'cn.Open "Driver={SQL Server};Server=ventas\kali;Database=calipso;Uid=sa"
    'cn.Open "Driver={SQL Server};Server=(local);Database=calipso;Uid=sa"
    'autentiacion windows
    'cn.Open "Driver={SQL Server};Data Source = " & menup.vservidor & "; Initial Catalog = " & buf & "'; Integrated Security = True"
    'autentuicacion nservidor
    'cn.Open "data source = ServidorSQL; initial catalog = " & buf & "; user id = sa; password = "
    'servidor remoto
    '
    'cn.Open "Driver={SQL Server}; data source = " & menup.vservidor & "; initial catalog = " & buf & "; user id = sa; password = "
 
    'funciona bien localmente
    'cn.Open "Provider=SQLNCLI; " & "Initial Catalog=" & buf & "; " & "Data Source=" & menup.vservidor & "; " & "integrated security=SSPI; persist security info=True;"
    'tinajas sqlserver 2005
    'cn.Open "Provider=SQLOLEDB.1; " & "Initial Catalog=" & buf & "; " & "Data Source=" & menup.vservidor & "; " & "integrated security=SSPI; persist security info=True;"
 
    conectar = 1
    Exit Function
cmd1_error:
    MsgBox " Aviso en Conectar " & error$, 48, "Aviso"
    Exit Function

End Function

Public Function conectara()

    Dim dbuser     As String

    Dim dbpassword As String

    Dim dbname     As String

    Dim dbserver   As String

    On Error GoTo cmd223_error
 
    cn1.CursorLocation = adUseClient
    cn1.CommandTimeout = 200000 '1024
    'cn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=calipso;Data Source=(local)"
    'cn.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=orion;Uid=sa"
    'cn.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=" & buf & " ;Uid=sa;Pwd="
    cn1.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=calipsot;Uid=sa;pwd="
 
    'cn.Open "Driver={SQL Server};Server=ventas\kali;Database=calipso;Uid=sa"
    'cn.Open "Driver={SQL Server};Server=(local);Database=calipso;Uid=sa"
    conectar1 = 1
    Exit Function
cmd223_error:
    MsgBox " conectara " & error$, 48, "Aviso"
    Exit Function

End Function
 
Function conectar1()

    Dim dbuser     As String

    Dim dbpassword As String

    Dim dbname     As String

    Dim dbserver   As String
 
    dbuser = "hackeem"              'This is a MS SQL Server User Login Name
    dbpassword = "kilburn"          'Login Password
    dbname = "programDB"            'MS SQL Server Database Name
    dbserver = "192.168.1.108"      'This is the Host computer on a network
    ' You may change it into localhost if you are running on a server.
    ' or if you are on a network use the computer name or IP address where MS SQL Server resides.
    'SQLcon.Open "Provider=SQLOLEDB.1; User ID=" & dbuser & ";Password=" & dbpassword & ";Initial Catalog=" & dbname & "; Data Source=" & dbserver

End Function

Public Function fin_del_Mes(fecha As Variant) As Date

    If IsDate(fecha) Then
        fin_del_Mes = DateAdd("m", 1, fecha)
        fin_del_Mes = DateSerial(Year(fin_del_Mes), Month(fin_del_Mes), 1)
        fin_del_Mes = DateAdd("d", -1, fin_del_Mes)

    End If
  
End Function

'Devuelve el último día de la semana
 
Function fin_de_Semana(ByVal fecha As Date) As Date
  
    If IsDate(fecha) Then
        fin_de_Semana = FormatDateTime(fecha - Weekday(fecha) + 7, vbGeneralDate)

    End If
  
End Function
 
