VERSION 5.00
Begin VB.Form replanil 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte Planilla"
   ClientHeight    =   3300
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox tipopla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox division 
      Height          =   375
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   8
      Text            =   "*"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox nrolineas 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "44"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox titulo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   4
      Text            =   "Reporte Planilla Periodo"
      Top             =   2040
      Width           =   4935
   End
   Begin VB.TextBox periodo 
      Height          =   375
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   1
      Text            =   "*"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox codigo 
      Height          =   375
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   0
      Text            =   "*"
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Planilla"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Division"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro.Lineas Reporte"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo Reporte"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodo:mmaaaa"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Ejecutar"
   End
   Begin VB.Menu ldo342 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "replanil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub djuer1_Click()
Dim found As Integer
Dim mytablex As New ADODB.Recordset
Dim buf As String
contlin = 0
suma1 = 0
suma2 = 0
suma3 = 0
ssuma1 = 0
ssuma2 = 0
ssuma3 = 0
If tipopla = "%" Then
   MsgBox "Seleecione un tipo de planilla", 48, "Aviso"
   Exit Sub
End If
If periodo = "%" Then
   MsgBox "Seleecione un Periodo de planilla", 48, "Aviso"
   Exit Sub
End If

found = sql_documento(mytablex)
If found = 0 Then
   mytablex.Close
    
   Exit Sub
End If
    Filename = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & Filename)
    Open Filename For Append As #1
    '------------------------------------
    If opcion2 = "1" Then  'reporte de planillas
    cabecera_documento
    cuerpo_programa_documento mytablex
    End If
    If opcion2 = "2" Then 'planilla salarios cada empleado
    cuerpo_programa_documento1 mytablex
    End If
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    genver.File = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub
Sub cabecera_documento1(mytablex As Table)
Dim buf As String
Dim i As Integer
Dim j As Integer
Dim found As Integer
Dim mytabley As Table
Dim ynombre As String
Dim ycodigo As String
Dim ycargo As String
Dim ysspp As String
Dim yfechaingr As String
Dim yfechavaca As String
Dim yfechacese As String
ReDim remune(20, 20) As String
ReDim xcolum(4) As Integer
Dim sdx As Double
Dim may As Integer
    buf = "EMPRESA PRUEBA"
       found = formateaa(buf, 90, 2, 0)
       found = formateaa("Direccion   : " & "Las mercedes xx", 40, 2, 0)
       found = formateaa("Ruc         : " & "2043333333", 40, 2, 0)
       found = formateaa("Reg.Patronal: " & "RP1212", 40, 2, 0)
    buf = String(150, "_")
    found = formateaa(buf, 90, 2, 0)
    cabecera_tipico "", "", "" & "" & gusuario
      ynombre = ""
      ycodigo = ""
      ycargo = ""
      ysspp = ""
      yfechaingr = ""
      yfechavaca = ""
      yfechacese = ""
    
    Set mytabley = mydbxglo.OpenTable("vendedor")
    mytabley.Index = "codigo"
    mytabley.Seek "=", "" & mytablex.Fields("Codigo")
    If Not mytabley.NoMatch Then
      ynombre = "" & mytabley.Fields("nombre")
      ycodigo = "" & mytabley.Fields("codigo")
      ycargo = "" & mytabley.Fields("cargo")
      ysspp = "" & mytabley.Fields("ipss")
      yfechaingr = "" & mytabley.Fields("fechaingr")
      yfechavaca = "" & mytabley.Fields("fechavaca")
      yfechacese = "" & mytabley.Fields("fechacese")
    End If
    mytabley.Close
     
       
    found = formateaa("Nombres y Apellidos :" & ynombre, 60, 0, 0)
    found = formateaa(" ", 2, 0, 0)
    found = formateaa("Codigo:" & "" & ycodigo, 15, 2, 0)
    found = formateaa("Ocupacion           :" & ycargo, 60, 0, 0)
    found = formateaa(" ", 2, 0, 0)
    found = formateaa("Cod.SSP:" & "" & ysspp, 15, 2, 0)
    found = formateaa("Fecha Ingreso       :" & yfechaingr, 30, 0, 0)
    found = formateaa(" ", 2, 0, 0)
    found = formateaa("Fecha Vaca  :" & yfechavaca, 30, 0, 0)
    found = formateaa(" ", 2, 0, 0)
    found = formateaa("Fecha Cese  :" & yfechacese, 30, 0, 0)
    found = formateaa(" ", 2, 2, 0)
    buf = String(150, "_")
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("REMUNERACIONES ", 29, 0, 0)
    found = formateaa("|", 1, 0, 0)
    found = formateaa("APORT. Y DESCTOS TRABAJ.", 29, 0, 0)
    found = formateaa("|", 1, 0, 0)
    found = formateaa("APORTACIONES PATRONALES ", 29, 0, 0)
    found = formateaa("|", 1, 2, 0)
    buf = String(150, "_")
    found = formateaa(buf, 90, 2, 0)
    'ahora hay que imprimir los 3 archivos
    'en tres columnas
    For i = 1 To 20
        For j = 1 To 20
            remune(i, j) = ""
        Next j
    Next i
   i = 0
   
   Set mytabley = mydbxglo.OpenTable("remune02")
   mytabley.Index = "remune020"
   mytabley.Seek "=", "" & mytablex.Fields("tipopla"), "" & mytablex.Fields("codigo"), "" & mytablex.Fields("periodo")
   If Not mytabley.NoMatch Then
      Do
      If mytabley.EOF Then Exit Do
      If "" & mytabley.Fields("tipopla") = "" & mytablex.Fields("tipopla") And "" & mytabley.Fields("codigo") = "" & mytablex.Fields("codigo") And "" & mytabley.Fields("periodo") = "" & mytablex.Fields("periodo") Then
             i = i + 1
             remune(i, 1) = "" & mytabley.Fields("tipo")
             remune(i, 2) = "" & mytabley.Fields("concepto")
             remune(i, 3) = "" & mytabley.Fields("importe")
         Else: Exit Do
      End If
      mytabley.MoveNext
      Loop
   End If
   mytabley.Close
    
   xcolum(1) = i
   
   i = 0
   
   Set mytabley = mydbxglo.OpenTable("descue02")
   mytabley.Index = "descue020"
   mytabley.Seek "=", "" & mytablex.Fields("tipopla"), "" & mytablex.Fields("codigo"), "" & mytablex.Fields("periodo")
   If Not mytabley.NoMatch Then
      Do
      If mytabley.EOF Then Exit Do
      If "" & mytabley.Fields("tipopla") = "" & mytablex.Fields("tipopla") And "" & mytabley.Fields("codigo") = "" & mytablex.Fields("codigo") And "" & mytabley.Fields("periodo") = "" & mytablex.Fields("periodo") Then
             i = i + 1
             remune(i, 4) = "" & mytabley.Fields("tipo")
             remune(i, 5) = "" & mytabley.Fields("concepto")
             remune(i, 6) = "" & mytabley.Fields("importe")
         Else: Exit Do
      End If
      mytabley.MoveNext
      Loop
   End If
   mytabley.Close
    
   xcolum(2) = i
   
   i = 0
   
   Set mytabley = mydbxglo.OpenTable("aporta02")
   mytabley.Index = "aporta020"
   mytabley.Seek "=", "" & mytablex.Fields("tipopla"), "" & mytablex.Fields("codigo"), "" & mytablex.Fields("periodo")
   If Not mytabley.NoMatch Then
      Do
      If mytabley.EOF Then Exit Do
      If "" & mytabley.Fields("tipopla") = "" & mytablex.Fields("tipopla") And "" & mytabley.Fields("codigo") = "" & mytablex.Fields("codigo") And "" & mytabley.Fields("periodo") = "" & mytablex.Fields("periodo") Then
             i = i + 1
             remune(i, 7) = "" & mytabley.Fields("tipo")
             remune(i, 8) = "" & mytabley.Fields("concepto")
             remune(i, 9) = "" & mytabley.Fields("importe")
         Else: Exit Do
      End If
      mytabley.MoveNext
      Loop
   End If
   mytabley.Close
    
   xcolum(3) = i
   '--------- imprimiendo
           may = xcolum(1)
        If xcolum(1) > xcolum(2) And xcolum(1) > xcolum(3) Then
              may = xcolum(1)
        End If
        If xcolum(2) > xcolum(1) And xcolum(2) > xcolum(3) Then
              may = xcolum(2)
        End If
        If xcolum(3) > xcolum(1) And xcolum(3) > xcolum(2) Then
              may = xcolum(3)
        End If
        'impresiones detalles ---------------------------------------
        contando = 0
        For i = 1 To may
            'Open Filename For Append As #2
            found = formateaa(remune(i, 1), 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa(remune(i, 2), 15, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa(Format(Val(remune(i, 3)), "0.00"), 8, 0, 1)
            found = formateaa("", 2, 0, 0)

            found = formateaa(remune(i, 4), 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa(remune(i, 5), 15, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa(Format(Val(remune(i, 6)), "0.00"), 8, 0, 1)
            found = formateaa("", 2, 0, 0)

            found = formateaa(remune(i, 7), 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa(remune(i, 8), 15, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa(Format(Val(remune(i, 9)), "0.00"), 8, 0, 1)
            found = formateaa("", 1, 2, 0)
            'Close #2
            contando = contando + 1
        Next i
        
        buf = String(150, "_")
        found = formateaa(buf, 90, 2, 0)
        
            found = formateaa("", 4, 0, 0)
            found = formateaa("TOTAL REMUN.", 16, 0, 0)
            found = formateaa(Format(Val("" & mytablex.Fields("ingreso")), "0.00"), 8, 0, 1)
            found = formateaa("", 2, 0, 0)
            
            found = formateaa("", 4, 0, 0)
            found = formateaa("TOTAL DESCTOS.", 16, 0, 0)
            found = formateaa(Format(Val("" & mytablex.Fields("descuento")), "0.00"), 8, 0, 1)
            found = formateaa("", 2, 0, 0)
            
            found = formateaa("", 4, 0, 0)
            found = formateaa("TOTAL APORTE EMP.", 16, 0, 0)
            found = formateaa(Format(Val("" & mytablex.Fields("aporta")), "0.00"), 8, 0, 1)
            found = formateaa("", 2, 2, 0)
            
        buf = String(150, "_")
        found = formateaa(buf, 90, 2, 0)
        
            found = formateaa("", 4, 0, 0)
            found = formateaa("NETO PAGAR.S/.", 16, 0, 0)
            sdx = Val("" & mytablex.Fields("ingreso")) - Val("" & mytablex.Fields("descuento"))
            buf = Format(sdx, "0.00")
            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 2, 2, 0)
            
            found = formateaa("", 1, 2, 0)
            found = formateaa("", 1, 2, 0)
            
            found = formateaa("", 4, 0, 0)
            found = formateaa("__________________", 16, 0, 0)
            found = formateaa("", 20, 0, 0)
            found = formateaa("__________________", 16, 2, 0)
            
            
            found = formateaa("", 4, 0, 0)
            found = formateaa("Empleador", 16, 0, 0)
            found = formateaa("", 20, 0, 0)
            found = formateaa("Recibi Conforme", 16, 2, 0)
            
            found = formateaa("", 1, 2, 0)
            found = formateaa("", 1, 2, 0)
            
            
            

   
   
End Sub

Sub cabecera_documento()
Dim buf As String
Dim i As Integer
Dim found As Integer
    If contlin > 0 Then
       buf = Chr$(12)
       found = formateaa(buf, Len(buf), 0, 0)
    End If
    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fecha   : " & Format(Now, "dd/mm/yyyy"), 25, 2, 0)
    found = formateaa("Periodo : " & periodo, 25, 2, 0)
    found = formateaa("Codigo  : " & codigo, 25, 2, 0)
    found = formateaa("Division: " & division, 25, 2, 0)
    
    
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)
    found = formateaa("tipopla", 7, 0, 0)
    found = formateaa("Codigo", 12, 0, 0)
    found = formateaa("Nombre", 31, 0, 0)
    found = formateaa("Ingreso ", 12, 0, 1)
    found = formateaa("Aporte ", 12, 0, 1)
    found = formateaa("Descuento ", 12, 0, 1)
    found = formateaa("Total ", 12, 0, 1)
    found = formateaa("Diatra ", 8, 0, 0)
    found = formateaa("Horatra", 8, 0, 0)
    found = formateaa("HoraExt ", 8, 2, 0)
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)
    

End Sub
Sub cuerpo_programa_documento1(mytablex As ADODB.Recordset)
Dim found As Integer
Do
If mytablex.EOF Then Exit Do
   cabecera_documento1 mytablex
   mytablex.MoveNext
Loop
   
End Sub

Sub cuerpo_programa_documento(mytablex As ADODB.Recordset)
Dim tmp As String
Dim sw As Integer
Dim buf As String
Dim found As Integer
Dim sdx As Double
sdx = 0
sw = 0
suma1 = 0
suma2 = 0
suma3 = 0
suma4 = 0
ssuma1 = 0
ssuma2 = 0
ssuma3 = 0
ssuma4 = 0

Do
If mytablex.EOF Then Exit Do
If sw = 0 Then
   buf = "" & mytablex.Fields("periodo")
   found = formateaa(buf, 6, 0, 0)
   found = formateaa("", 1, 2, 0)
   nlineas
   sw = 1
   suma1 = 0
   suma2 = 0
   suma3 = 0
   suma4 = 0
   tmp = "" & mytablex.Fields("periodo")
End If
If tmp <> "" & mytablex.Fields("periodo") Then
   found = formateaa("", 50, 0, 0)
   buf = Format(suma1, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma2, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma3, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma4, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 2, 0)
   nlineas

      
   buf = "" & mytablex.Fields("periodo")
   found = formateaa(buf, 6, 0, 0)
   found = formateaa("", 1, 2, 0)
   nlineas
   tmp = "" & mytablex.Fields("periodo")
   suma1 = 0
   suma2 = 0
   suma3 = 0
   suma4 = 0
End If
   buf = "" & mytablex.Fields("tipopla")
   found = formateaa(buf, 6, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("Codigo")
   found = formateaa(buf, 11, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = busca_vendedor("" & mytablex.Fields("Codigo"))
   found = formateaa(buf, 30, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("ingreso")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("aporta")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("descuento")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   sdx = Val("" & mytablex.Fields("ingreso")) - Val("" & mytablex.Fields("descuento"))
   buf = Format(sdx, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   
   buf = "" & mytablex.Fields("diatraba")
   found = formateaa(buf, 7, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("horatraba")
   found = formateaa(buf, 7, 0, 0)
   found = formateaa("", 1, 0, 0)
   buf = "" & mytablex.Fields("horaextr")
   found = formateaa(buf, 7, 0, 0)
   found = formateaa("", 1, 2, 0)
   
   
   suma1 = suma1 + Val("" & mytablex.Fields("ingreso"))
   suma2 = suma2 + Val("" & mytablex.Fields("aporta"))
   suma3 = suma3 + Val("" & mytablex.Fields("descuento"))
   suma4 = suma4 + sdx
   
   ssuma1 = ssuma1 + Val("" & mytablex.Fields("ingreso"))
   ssuma2 = ssuma2 + Val("" & mytablex.Fields("aporta"))
   ssuma3 = ssuma3 + Val("" & mytablex.Fields("descuento"))
   ssuma4 = ssuma4 + sdx
   
   nlineas
mytablex.MoveNext
Loop
   found = formateaa("", 50, 0, 0)
   buf = Format(suma1, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma2, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma3, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(suma4, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 2, 0)
   nlineas
   
   found = formateaa("Total-->   ", 50, 0, 1)
   buf = Format(ssuma1, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(ssuma2, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(ssuma3, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 0, 0)
   buf = Format(ssuma4, "0.00")
   found = formateaa(buf, 11, 0, 1)
   found = formateaa("", 1, 2, 0)
   
   
   
End Sub
Function sql_documento(mytablex As Snapshot)
Dim buf As String
buf = "select * from sisper02 where "
buf = buf & "  codigo like '" & codigo & "'"
If tipopla <> "%" Then
buf = buf & " and tipopla like '" & tipopla & "'"
End If
If periodo <> "%" Then
  buf = buf & " and periodo like '" & periodo & "'"
End If
If division <> "%" Then
  buf = buf & " and division like '" & periodo & "'"
End If
buf = buf & "order by Periodo,tipopla,Codigo"

mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
sql_documento = 1
End Function
Function busca_vendedor(buf As String) As String
Dim mytabley As Table


   Set mytabley = mydbxglo.OpenTable("vendedor")
   mytabley.Index = "codigo"
   mytabley.Seek "=", buf
   If Not mytabley.NoMatch Then
      busca_vendedor = "" & mytabley.Fields("NOMBRE")
   End If
   mytabley.Close
    

End Function





Sub nlineas()
    contlin = contlin + 1
    If contlin > Val(nrolineas) Then
       cabecera_documento
    End If
End Sub

Private Sub Form_Activate()
carga_inicial
End Sub

Private Sub ldo342_Click()
replanil.Hide
Unload replanil
End Sub
Sub carga_inicial()
Dim mytablex As New ADODB.Recordset

'tipopla.Clear
'tipopla.AddItem "%"

 '  Set mytablex = mydbxglo.OpenTable("tipopla")
 '  Do
 '  If mytablex.EOF Then Exit Do
 '  tipopla.AddItem "" & mytablex.Fields("tipopla")
 '  mytablex.MoveNext
 '  Loop
 '  mytablex.Close
 '
 '  tipopla.ListIndex = 0


'periodo.Clear
'periodo.AddItem "%"

   Set mytablex = mydbxglo.OpenTable("plaperiodo")
   Do
   If mytablex.EOF Then Exit Do
   periodo.AddItem "" & mytablex.Fields("periodo")
   mytablex.MoveNext
   Loop
   mytablex.Close
    
   'periodo.ListIndex = 0

End Sub

