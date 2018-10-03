Attribute VB_Name = "Module18"
Option Explicit
  
'Agregar la referencia de ADO
'----------------------------------------------------

Sub Excel_a_Access(Path_XLS As String, Filas As Integer, columnas As Integer)
  
    Dim Obj_Excel      As Object

    Dim Obj_Hoja       As Object

    Dim rst_Ado        As New ADODB.Recordset

    Dim Fila_Actual    As Integer

    Dim Columna_Actual As Integer

    Dim DATO
  
    Screen.MousePointer = vbHourglass
  
    'Nueva instancia de Excel
    Set Obj_Excel = CreateObject("Excel.Application")
  
    ' Abre el libro de Excel
    Obj_Excel.Workbooks.Open FileName:=Path_XLS
  
    ' si es la versión de Excel 97, asigna la hoja activa ( ActiveSheet )
    If Val(Obj_Excel.Application.Version) >= 8 Then
        Set Obj_Hoja = Obj_Excel.ActiveSheet
    Else
        Set Obj_Hoja = Obj_Excel

    End If
      
    cn.Execute ("delete from clientesborrar")
    
    rst_Ado.Open "Select * from clientesborrar ", cn, adOpenStatic, adLockOptimistic
      
    'Se posiciona al final    If rst_Ado.RecordCount <> 0 Then rst_Ado.MoveLast
    ' Recorre las filas y columnas de la hoja
    For Fila_Actual = 1 To Filas
        'Nuevo registro
        rst_Ado.AddNew

        For Columna_Actual = 0 To columnas - 1
            ' Va leyendo los datos de la celda indicada
            DATO = Obj_Hoja.Cells(Fila_Actual, Columna_Actual + 1)
            'MsgBox DATO
            'Agrega los datos al campo indicado
            rst_Ado.Fields(Columna_Actual) = "" & DATO
        Next
        rst_Ado.Update
    Next
      
    Call Descargar_Objetos(rst_Ado, Obj_Excel, Obj_Hoja)
    Screen.MousePointer = vbDefault
    MsgBox " Datos copiados ", vbInformation
  
    Exit Sub
  
    'Error
ErrSub:
  
    Call Descargar_Objetos(rst_Ado, Obj_Excel, Obj_Hoja)
    MsgBox Err.Description, vbCritical
    Screen.MousePointer = vbDefault
      
End Sub
  
'Descarga los objetos y los cierra
Sub Descargar_Objetos(rst_Ado As ADODB.Recordset, Obj_Excel As Object, Obj_Hoja As Object)
      
    Set rst_Ado = Nothing
    Obj_Excel.ActiveWorkbook.Close False
    Obj_Excel.Quit
    Set Obj_Hoja = Nothing
    Set Obj_Excel = Nothing
  
End Sub

'Correo Reportes kenyo 18/04/2017
Sub envio_correosReportes()

    'Dim txtserver As String
    'Dim txtusername As String
    'Dim txtpassword As String
    'Dim txtport As String
    'Dim txtto As String
    'Dim chkssl As String
    'Dim txtfromname As String
    'Dim txtfromemail As String
    'Dim txtattach As String
    'Dim txtsubject As String
    'Dim txtmsg As String
    'Dim retval As String
    'Dim txthtml As String
    'Dim txtselecciona As String
    ''Dim txtselecciona As String
    'Dim mytablex As New ADODB.Recordset
    'Dim buf As String
    'On Error GoTo cmd0905677_err
    '
    '   mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
    '   If mytablex.RecordCount > 0 Then
    '     If "" & mytablex.Fields("correocierre") <> "S" Then
    '        mytablex.Close
    '        Exit Sub
    '     End If
    '   End If
    '   mytablex.Close
    '
    ''buf = extra_loquesea1(perfil)
    ''If Trim(buf) = 0 Then Exit Sub
    'mytablex.Open "select * from correos where cosms='11'", cn, adOpenStatic, adLockOptimistic
    'If mytablex.RecordCount > 0 Then
    'txtserver = Trim("" & mytablex.Fields("txtserver"))
    'txtusername = Trim("" & mytablex.Fields("txtusername"))
    'txtpassword = Trim("" & mytablex.Fields("txtpassword"))
    'txtfromname = Trim("" & mytablex.Fields("txtfromname"))
    'txtfromemail = Trim("" & mytablex.Fields("txtfromemail"))
    'txtport = Trim("" & mytablex.Fields("txtport"))
    'txtselecciona = Trim("" & mytablex.Fields("txtselecciona"))
    'chkssl = Trim("" & mytablex.Fields("chkssl"))
    '
    'txtto = Trim("" & mytablex.Fields("txtfromemail"))
    '
    'txtattach = "C:\REPORTE.xls"
    '
    'txtsubject = Trim("" & mytablex.Fields("txtsubject"))
    'txtmsg = Trim("" & mytablex.Fields("txtmsg"))
    'txtmsg = txtmsg & Chr$(10) & Chr$(13) & ""
    'txtmsg = txtmsg & Format(Now, "dd/mm/yyyy") + " " + Format(Now, "hh:mm:ss")
    '
    'If Len(Trim("" & mytablex.Fields("txtfromemail"))) > 0 Then
    '   txtto = Trim("" & mytablex.Fields("txtfromemail"))
    '   retval = SendMail(Trim$(txtto), Trim$(txtsubject), Trim$(txtfromname) & "<" & Trim$(txtfromemail) & ">", Trim$(txtmsg), Trim$(txtserver), CInt(Trim$(txtport)), Trim$(txtusername), Trim$(txtpassword), Trim$(txtattach), True, txtselecciona, txthtml)
    'End If
    '
    'MsgBox "Correo Enviado ", 48, "Aviso"
    'End If
    'mytablex.Close
    '
    'Exit Sub
    'cmd0905677_err:
    'MsgBox "No se Pudo enviar Correo... " + error$, 48, "Aviso"
    'Exit Sub
    '
End Sub
  
Private Sub Command1_Click()
  
    ' Pasar como parámetro el nombre y path de la _
      base de datos y del libro excel, el nombre de la tabla _
      y la cantidad de filas y columnas de la hoja a leer
  
    'Call Excel_a_Access(App.path & "C:\ORION.V5\tclientes.xlsx", "CLIENTESBORRAR", 10, 3)
                      
End Sub

