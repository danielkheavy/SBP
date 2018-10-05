VERSION 5.00
Begin VB.Form exporac 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportaciones.."
   ClientHeight    =   7980
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   10965
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Preparacion archivo Sunat"
      Height          =   4095
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox CDR 
         Height          =   375
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "CDR"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox BB 
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   23
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox AAAA 
         Height          =   375
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   22
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox MM 
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   21
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox DD 
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   20
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox XX 
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   19
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Generar Archivo"
         Height          =   615
         Left            =   5880
         TabIndex        =   18
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cerrar ventana"
         Height          =   615
         Left            =   5880
         TabIndex        =   17
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox ruc 
         Height          =   375
         Left            =   1320
         MaxLength       =   11
         TabIndex        =   16
         Text            =   "20420605006"
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CDR"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BB"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AAAA"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MM"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DD"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "XX"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUC"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.Frame panel3d1 
      BackColor       =   &H0080FF80&
      Caption         =   "Procesos..."
      Height          =   2895
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox CheCK3d1 
         BackColor       =   &H0080FF80&
         Caption         =   "Paso1"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   3015
      End
      Begin VB.CheckBox Check3d2 
         BackColor       =   &H0080FF80&
         Caption         =   "Paso2"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   3015
      End
      Begin VB.CheckBox Check3d3 
         BackColor       =   &H0080FF80&
         Caption         =   "Paso3"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CheckBox Check3d4 
         BackColor       =   &H0080FF80&
         Caption         =   "Paso4"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No tocar computadora hasta esperar mensaje de Finalizacion.."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   5175
      End
   End
   Begin VB.TextBox nombre 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   7
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox fechaf 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox fechai 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NombreArchivoClie..  Solo importacionCliente"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo Operacion"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "fecha Final"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "fecha Inicio"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu flosa92 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "exporac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AAAA_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
MM.SetFocus

End Sub

Private Sub BB_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
AAAA.SetFocus

End Sub

Private Sub CDR_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
BB.SetFocus
End Sub

Private Sub Command1_Click()
If Len(fechai) <> 10 Then Exit Sub
If Len(fechaf) <> 10 Then Exit Sub
If Val(Mid$(fechai, 1, 2)) < 1 And Val(Mid$(fechai, 1, 2)) > 31 Then Exit Sub
If Val(Mid$(fechai, 4, 2)) < 1 And Val(Mid$(fechai, 4, 2)) > 12 Then Exit Sub
If Val(Mid$(fechai, 7, 4)) < 2008 Then Exit Sub

If Val(Mid$(fechaf, 1, 2)) < 1 And Val(Mid$(fechaf, 1, 2)) > 31 Then Exit Sub
If Val(Mid$(fechaf, 4, 2)) < 1 And Val(Mid$(fechaf, 4, 2)) > 12 Then Exit Sub
If Val(Mid$(fechaf, 7, 4)) < 2008 Then Exit Sub
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub


If Combo1 = "ImportaCliente" Then
   If Len(nombre) = 0 Then
      nombre.SetFocus
      Exit Sub
   End If
   If Dir$(globaldat & "\exporta\" & nombre & ".dbf") <> "" Then
      panel3d1.Visible = True
      sql_importa_cliente
      panel3d1.Visible = False
      Else
      nombre.SetFocus
      Exit Sub
   End If
   Exit Sub
End If
If Combo1 = "Exportar" Then
   panel3d1.Visible = True
   sql_exporta
   panel3d1.Visible = False
   

   Exit Sub
End If
'If Combo1 = "Importar" Then
'   panel3d1.Visible = True
'   sql_importa
'   panel3d1.Visible = False
''
'
'End If
If Combo1 = "ExporOracle" Then
   panel3d1.Visible = True
   exporta_sdf
   panel3d1.Visible = False

End If
If Combo1 = "ExportarDetraccion" Then
   exporta_sunat
   Exit Sub

End If

'If Combo1 = "CentralizaCajas" Then
'
'   panel3d1.Visible = True
'   central_caja
'   panel3d1.Visible = False
''
'
'End If


   



End Sub
Sub sql_importa_cliente()
Dim buf As String
Dim mydbx As Database
Dim mydby As Database
Dim mytabley As Table
Dim mytablex As Table
Dim buf1 As String
Dim found As Integer
On Error GoTo cmd4412_err
    '----copiar los datos de exporta -------
    'Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
    'buf1 = "DELETE FROM clientes "
    'mydbx.Execute buf1
    'mydbx.Close
    
    Set mydby = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
    Set mydbx = OpenDatabase(globaldir & "\EXPORTA", False, False, "foxpro 2.5;")
    Set mytablex = mydbx.OpenTable(nombre)
    Set mytabley = mydby.OpenTable("clientes")
    mytabley.Index = "CODIGO"
   Do
    If mytablex.EOF Then Exit Do
       found = valida_ruc("" & mytablex.Fields("codigo"))
       If found <> 0 Then
          mytabley.Seek "=", "" & mytablex.Fields("codigo")
          If mytabley.NoMatch Then
          mytabley.AddNew
          mytabley.Fields("codigo") = "" & mytablex.Fields("codigo")
          'mytableY.Fields("ruc") = "" & mytablex.Fields("ruc")
          mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
          'mytableY.Fields("sexo") = "" & mytablex.Fields("sexo")
          mytabley.Fields("moneda") = "S"
          'mytableY.Fields("tipo") = "C"
          'mytableY.Fields("ESTADO") = "0"
          mytabley.Update
          End If
       End If
    mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    mydbx.Close
    mydby.Close
    CheCK3d1.Value = 1
    MsgBox "Proceso Terminado ", 48, "Aviso"
    'dfslo34_Click
    Exit Sub
cmd4412_err:
    MsgBox "Error,Verificar y comenzar de Nuevo1  " & error$, 48, "Aviso"
    End
    Exit Sub

End Sub

Sub exporta_sdf()
Dim buf As String
Dim mydbx As Database
Dim mytablex As Table
Dim buf1 As String
Dim buf2 As String
Dim found As Integer
On Error GoTo cmd439_err
    buf1 = Format(Month(fechai), "00") & Mid$(Format(Year(fechai), "0000"), 3, 2) & "EC" 'cabecera
    If Dir$(globaldat & "\_" & buf1 & ".dbf") <> "" Then
      Kill globaldat & "\_" & buf1 & ".dbf"
    End If
    buf = "select * into _" & buf1 & " from factura where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    buf = buf & " and (tipo='1' or tipo='2' or tipo='3' or tipo='4')"
    Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
    mydbx.Execute buf
    mydbx.Close
    CheCK3d1.Value = 1
    'AHORA TRABAJANDO LA EXPORTACION
    If Dir$(globaldat & "\_" & buf1 & ".dbf") <> "" Then
       'existe
       Else
       MsgBox "Sin Datos ", 48, "Aviso"
       Exit Sub
    End If
    Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
    Set mytablex = mydbx.OpenTable("_" & buf1)
    
    found = borra_nombre(globaldir & "\exporta\oracle.csv")
    Open globaldir & "\exporta\oracle.csv" For Append As #1
    Do
    If mytablex.EOF Then Exit Do
       buf2 = ""
            '----------------------------------------
            buf1 = ""
            If "" & mytablex.Fields("tipo") = "1" Then
            buf1 = "12,"
            End If
            If "" & mytablex.Fields("tipo") = "2" Then
            buf1 = "12,"
            End If
            If "" & mytablex.Fields("tipo") = "3" Then
            buf1 = "03,"
            End If
            If "" & mytablex.Fields("tipo") = "4" Then
            buf1 = "01,"
            End If
      
            If Val("" & mytablex.Fields("local")) = 1 Then
               buf1 = buf1 & "1" & ","
               buf2 = "" & mytablex.Fields("serie")
            End If
            If Val("" & mytablex.Fields("local")) = 2 Then
               buf1 = buf1 & "2" & ","
               buf2 = "" & mytablex.Fields("serie")
            End If
            If Val("" & mytablex.Fields("local")) = 3 Then
               buf1 = buf1 & "3" & ","
               buf2 = "" & mytablex.Fields("serie")
            End If
            If Val("" & mytablex.Fields("local")) = 4 Then
               buf1 = buf1 & "4" & ","
               buf2 = "" & mytablex.Fields("serie")
            End If
            If Val("" & mytablex.Fields("local")) = 5 Then
               buf1 = buf1 & "5" & ","
               buf2 = "" & mytablex.Fields("serie")
            End If
            buf1 = buf1 & Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy") & ","
            If "" & mytablex.Fields("tipo") = "1" Or "" & mytablex.Fields("tipo") = "2" Then  'serie
               buf1 = buf1 & buf2 & ","
            End If
            If "" & mytablex.Fields("tipo") = "2" Or "" & mytablex.Fields("tipo") = "4" Then
               buf1 = buf1 & Mid$("" & mytablex.Fields("numero"), 1, 3) & ","
            End If
            buf1 = buf1 & Mid$("" & mytablex.Fields("numero"), 5, 7) & ","
            buf1 = buf1 & "" & mytablex.Fields("moneda") & "/." & ","

            'codigo
            
               If "" & mytablex.Fields("estado") = "1" Then  'anulado
                  buf1 = buf1 & "00000000,"
               End If
               If "" & mytablex.Fields("estado") = "2" Then  'normal
                  buf1 = buf1 & rellena_ceros("" & mytablex.Fields("codigo")) & ","
               End If
            
            'razon social
            If "" & mytablex.Fields("estado") = "1" Then  'anulado
            buf1 = buf1 & "Anulado,"
            End If
            If "" & mytablex.Fields("estado") = "2" Then  'No anulado
            buf1 = buf1 & "" & limpiar_coma("" & mytablex.Fields("nombre")) & ","
            End If
            'subtotal
            If "" & mytablex.Fields("estado") = "1" Then  'anulado
            buf1 = buf1 & "0.00,"  'subtotal
            buf1 = buf1 & "0.00,"  'impuesto
            buf1 = buf1 & "0.00,"  'total
            
            End If
            If "" & mytablex.Fields("estado") = "2" Then  'No anulado
            buf1 = buf1 & Format(Val("" & mytablex.Fields("subtotal")), "0.00") & ","
            buf1 = buf1 & Format(Val("" & mytablex.Fields("impuesto")), "0.00") & ","
            buf1 = buf1 & Format(Val("" & mytablex.Fields("total")), "0.00") & ","
            End If
            'ccosto
            If Val("" & mytablex.Fields("local")) = 1 Then
               buf1 = buf1 & "K001" & ","
            End If
            If Val("" & mytablex.Fields("local")) = 2 Then
               buf1 = buf1 & "K002" & ","
            
            End If
            If Val("" & mytablex.Fields("local")) = 3 Then
               buf1 = buf1 & "K003" & ","
            
            End If
            If Val("" & mytablex.Fields("local")) = 4 Then
               buf1 = buf1 & "K004" & ","
            
            End If
            If Val("" & mytablex.Fields("local")) = 5 Then
               buf1 = buf1 & "K005" & ","
            End If
            'anulada
            If "" & mytablex.Fields("estado") = "1" Then  'anulado
            buf1 = buf1 & "S,"
            End If
            If "" & mytablex.Fields("estado") = "2" Then  'No anulado
            buf1 = buf1 & "N,"
            End If
            buf1 = buf1 & "01090002,"
            buf1 = buf1 & "40110002,"
            buf1 = buf1 & "07015001,"
            buf1 = buf1 & "16100050"
            Print #1, buf1
    mytablex.MoveNext
    Loop
    
    mytablex.Close
    mydbx.Close
    Close #1
    Check3d2.Value = 1
    MsgBox "Proceso Terminado ", 48, "Aviso"
    Exit Sub
cmd439_err:
    MsgBox "Error,Verificar y comenzar de Nuevo  " & error$, 48, "Aviso"
    End
    Exit Sub

End Sub
Function rellena_ceros(buf As String) As String
Dim i As Integer
Dim buf1 As String
Dim X As Integer
X = 11 - Len(buf)
buf1 = ""
If X = 0 Then
   rellena_ceros = buf
   Exit Function
End If
For i = 1 To X
    buf1 = buf1 + "9"
Next i
'MsgBox buf1
rellena_ceros = buf1 + buf

End Function

Function limpiar_coma(buf As String)
Dim buf1 As String
Dim i As Integer
buf1 = ""
If Len(buf) > 1 Then
For i = 1 To Len(buf)
  If Mid$(buf, i, 1) = "," Then
    Else
    buf1 = buf1 & Mid$(buf, i, 1)
  End If
Next i
End If
limpiar_coma = buf1
End Function


Private Sub Command3_Click()
Dim buf As String
Dim buf1 As String
Dim buf2 As String
Dim found As Integer
Dim xplaca As String
Dim xeje As String
Dim mytablex As Table
Dim mydbx As Database
Dim mysnapx As Snapshot

Dim xocurreg As Double
Dim xocurrei As Double
Dim xocurrec As Double
Dim xocurrep As Double

Dim i As Integer
On Error GoTo cmd4322_err
    If CDR <> "CDR" Then
       CDR.SetFocus
       Exit Sub
    End If
    If Len(BB) = 0 Then
       BB.SetFocus
       Exit Sub
    End If
    If Len(AAAA) <> 4 Then
       AAAA.SetFocus
       Exit Sub
    End If
    If Not IsNumeric(AAAA) Then
      AAAA.SetFocus
    End If
    If Len(DD) <> 2 Then
       DD.SetFocus
       Exit Sub
    End If
    If Not IsNumeric(DD) Then
      DD.SetFocus
      Exit Sub
    End If
    If Val(DD) < 1 And Val(DD) > 31 Then
       DD.SetFocus
       Exit Sub
    End If
    If Not IsNumeric(XX) Then
       XX.SetFocus
       Exit Sub
    End If
   Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
   buf2 = "select * from factura where "
   buf2 = buf2 & " fecha>=" & "DateValue('" & fechai & "'" & ")"
   buf2 = buf2 & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"
   buf2 = buf2 & " and estado='2' "
   Set mysnapx = mydbxglo.CreateSnapshot(buf2)
   xocurreg = 0
   xocurrei = 0
   xocurrec = 0
   xocurrep = 0
   Do
   If mysnapx.EOF Then Exit Do
   xocurreg = xocurreg + 1
   xocurrei = xocurrei + Val("" & mysnapx.Fields("total"))
   xocurrec = xocurrec + Val("" & mysnapx.Fields("total"))
   xocurrep = xocurrep + Val("" & mysnapx.Fields("total"))
   mysnapx.MoveNext
   Loop
   
   Set mytablex = mydbx.OpenTable("detalle")
   mytablex.Index = "tdetalle"
   buf = CDR + BB + AAAA + MM + DD + XX
   Open globaldir & "\sunat\" & buf & ".txt" For Append As #1
   'cabecera----------------------------------------
   buf1 = ruc & ","
   buf1 = buf1 & Format(Val(xocurreg), "000000") & ","
   buf1 = buf1 & Format(Val(xocurrei), "000000") & ","
   buf1 = buf1 & convierte_decimal(Val(xocurrec)) & ","
   buf1 = buf1 & convierte_decimal(Val(xocurrep))
   Print #1, buf1
   mysnapx.MoveFirst
   
Do
If mysnapx.EOF Then Exit Do
   
   mytablex.Seek "=", "" & mysnapx.Fields("local"), "" & mysnapx.Fields("tipo"), "" & mysnapx.Fields("serie"), "" & mysnapx.Fields("numero")
    If Not mytablex.NoMatch Then
    Do
      If "" & mytablex.Fields("local") = "" & mysnap.Fields("local") And "" & mytablex.Fields("tipo") = "" & mysnap.Fields("tipo") And "" & mytablex.Fields("serie") = "" & mysnap.Fields("serie") And "" & mytablex.Fields("numero") = "" & mysnap.Fields("numero") Then
         buf1 = "" & mysnapx.Fields("codigo") & ","    'ruc transportista
         buf1 = buf1 + "" & mytablex.Fields("placa") + ","   'placa
         buf1 = buf1 + "" & mytablex.Fields("subfamilia") + ","  'ejes
         buf1 = buf1 + "0,"  '0 no esta afecto  1 si esta afecto   2 no tiene calcomania
         buf1 = buf1 + "" & mytablex.Fields("numero") + ","  'numero constancia pagao
         buf1 = buf1 + "" & mytablex.Fields("denumero") + ","  'numero constancia detraccion
         buf1 = buf1 + "000,"   'codigo estacion peaje
         buf1 = buf1 + "000,"   'codigo garita o caseta peaje
         buf1 = buf1 + "0,"     'sentido  1,cobro bidireccional  0 cobro unidireccional
         buf1 = buf1 + "00000000,"  'fecha emision aaaammdd
         buf1 = buf1 + "00000000,"  'hora emision hh:mm:ss
         buf1 = buf1 & "000000000000000,"  'monto a cobrar por detraccion
         buf1 = buf1 & "000000000000000,"  'monto pagado por la detraccion si no esta afecto 0 debe ser cero
         buf1 = buf1 & "000000000000000"   'monto de pago x concepto de peaje
         Print #1, buf1
         Else: Exit Do
      End If
      mytablex.MoveNext
    Loop
    End If
 mysnapx.Recordset.MoveNext
Loop
mytablex.Close
mysnapx.Close
mydbx.Close
    Close #1
    MsgBox "proceso Terminado"
    Exit Sub
cmd4322_err:
    MsgBox "Error " & error$, 24, "Aviso"
    Close #1
    Exit Sub

End Sub

Private Sub Command4_Click()
Frame1.Visible = False
End Sub

Private Sub DD_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
XX.SetFocus

End Sub

Private Sub flosa92_Click()
exporac.Hide
Unload exporac
End Sub

Private Sub Form_Load()
fechai = Format(Now, "dd/mm/yyyy")
fechaf = Format(Now, "dd/mm/yyyy")
Combo1.AddItem "Exportar"
Combo1.AddItem "ExporOracle"
Combo1.AddItem "ImportaCliente"
Combo1.AddItem "ExportarDetraccion"
Combo1.ListIndex = 0


End Sub
Function borra_nombre(buf As String)
On Error GoTo cmd457_err
   Kill buf
   borra_nombre = 1
   Exit Function
cmd457_err:
   Exit Function
End Function
Sub sql_exporta()
Dim buf As String
Dim mydbx As Database
Dim buf1 As String
On Error GoTo cmd43_err
    buf1 = Format(Month(fechai), "00") & Mid$(Format(Year(fechai), "0000"), 3, 2) & "EC" 'cabecera
    If Dir$(globaldat & "\_" & buf1 & ".dbf") <> "" Then
      Kill globaldat & "\_" & buf1 & ".dbf"
    End If
    buf = "select * into _" & buf1 & " from factura where    "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
    mydbx.Execute buf
    mydbx.Close
    CheCK3d1.Value = 1
    
    
    buf1 = Format(Month(fechai), "00") & Mid$(Format(Year(fechai), "0000"), 3, 2) & "ED" 'DETALLE
    If Dir$(globaldat & "\_" & buf1 & ".dbf") <> "" Then
      Kill globaldat & "\_" & buf1 & ".dbf"
    End If
    buf = "select * into _" & buf1 & " from detalle where    "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
    mydbx.Execute buf
    mydbx.Close
    Check3d2.Value = 1
    

    buf1 = Format(Month(fechai), "00") & Mid$(Format(Year(fechai), "0000"), 3, 2) & "EF" 'fpagov
    If Dir$(globaldat & "\_" & buf1 & ".dbf") <> "" Then
      Kill globaldat & "\_" & buf1 & ".dbf"
    End If
    buf = "select * into _" & buf1 & " from fpagov where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    Set mydbx = OpenDatabase(globaldat, False, False, "foxpro 2.5;")
    mydbx.Execute buf
    mydbx.Close
    Check3d3.Value = 1

    buf1 = Format(Month(fechai), "00") & Mid$(Format(Year(fechai), "0000"), 3, 2) & "EX"  'CLIENTES
    If Dir$(globaldir & "\_" & buf1 & ".dbf") <> "" Then
      Kill globaldir & "\_" & buf1 & ".dbf"
    End If
    buf = "select * into _" & buf1 & " from clientes "
    Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
    mydbx.Execute buf
    mydbx.Close
    'ahora copiando a exporta
    buf1 = Format(Month(fechai), "00") & Mid$(Format(Year(fechai), "0000"), 3, 2) & "EC"  'cabecera
    copiando globaldat & "\_" & buf1 & ".dbf", globaldir & "\exporta\_" & buf1 & ".dbf"
    buf1 = Format(Month(fechai), "00") & Mid$(Format(Year(fechai), "0000"), 3, 2) & "ED" 'DETALLE
    copiando globaldat & "\_" & buf1 & ".dbf", globaldir & "\exporta\_" & buf1 & ".dbf"
    buf1 = Format(Month(fechai), "00") & Mid$(Format(Year(fechai), "0000"), 3, 2) & "EF" 'fpagov
    copiando globaldat & "\_" & buf1 & ".dbf", globaldir & "\exporta\_" & buf1 & ".dbf"
    buf1 = Format(Month(fechai), "00") & Mid$(Format(Year(fechai), "0000"), 3, 2) & "EX"  'CLIENTES
    copiando globaldir & "\_" & buf1 & ".dbf", globaldir & "\exporta\_" & buf1 & ".dbf"
    Check3d4.Value = 1
    MsgBox "Proceso Terminado ", 48, "Aviso"
    
    Exit Sub
cmd43_err:
    MsgBox "Error,Verificar y comenzar de Nuevo  " & error$, 48, "Aviso"
    End
    Exit Sub

    
End Sub
Sub exporta_sunat()
Frame1.Visible = True
CDR = "CDR"
BB = "XX"
AAAA = Format(Year(Now), "YYYY")
MM = Format(Month(Now), "MM")
DD = Format(Day(Now), "DD")
XX = "01"
CDR.SetFocus
End Sub


Private Sub MM_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
DD.SetFocus

End Sub

Private Sub XX_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
ruc.SetFocus

End Sub
Function convierte_decimal(sdx As Double) As String
Dim buf As String
Dim bufd As String
buf = Format(sdx, "0000000000000.00")
bufd = Mid$(buf, Len(buf) - 1, 2)
buf = Mid$(buf, 1, 13) + bufd
convierte_decimal = buf
End Function

