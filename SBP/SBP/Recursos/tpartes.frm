VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tpartes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Explorador  Tarjetas de Produccion"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   14880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   2160
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   8895
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Ejecutar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox buffer 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "tpartes.frx":0000
         Height          =   4215
         Left            =   120
         OleObjectBlob   =   "tpartes.frx":0014
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1080
         Width           =   8655
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   14820
      TabIndex        =   8
      Top             =   0
      Width           =   14880
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tpartes.frx":09DF
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tpartes.frx":1BF1
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tpartes.frx":2E03
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Consulta"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consulta"
      Height          =   3855
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox tarjeta 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "*"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox plano 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   14
         Text            =   "*"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox seccion 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmdGrabar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6360
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tpartes.frx":4015
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6360
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tpartes.frx":47C3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tarjeta"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plano"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seccion"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "tpartes.frx":4F71
      Height          =   7575
      Left            =   0
      OleObjectBlob   =   "tpartes.frx":4F85
      TabIndex        =   0
      Top             =   600
      Width           =   13935
   End
   Begin VB.Menu djiowewe 
      Caption         =   "&Partes"
   End
   Begin VB.Menu imro331 
      Caption         =   "&Imprime"
      Begin VB.Menu kfdi343 
         Caption         =   "&1.Tarjetas-Sticker Solo toma el Registro"
      End
      Begin VB.Menu dki3434 
         Caption         =   "&3.Reporte"
      End
   End
   Begin VB.Menu dmui821 
      Caption         =   "&Consulta"
   End
   Begin VB.Menu dehy723 
      Caption         =   "&Cierre"
   End
   Begin VB.Menu ldfo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tpartes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub boy723_Click()

End Sub

Private Sub anul91_Click()

End Sub

Private Sub cmdAddEntry_Click()

End Sub

Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   ldfo33_Click
   Exit Sub
End If
Command1_Click
End Sub

Private Sub cmdCancelar_Click()
ldfo33_Click
End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdGrabar_Click()
sql_partes
ldfo33_Click
End Sub

Private Sub dehy723_Click()
On Error GoTo cmd13_err
If "" & Data2.Recordset.Fields("cierre") = "C" Then
   If MsgBox("Desea Invertir el proceso", 1, "Aviso") <> 1 Then Exit Sub
   proceso_grabacion 1
   Exit Sub
End If
If "" & Data2.Recordset.Fields("cierre") <> "C" Then
   If MsgBox("Desea Cerrar la  " + Data2.Recordset.Fields("descripcio"), 1, "Aviso") <> 1 Then Exit Sub
   proceso_grabacion 0
   Exit Sub
End If
Exit Sub
cmd13_err:
MsgBox "Seleccione un dato " + error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub cmdSort_Click()
dmui821_Click
End Sub

Private Sub Command1_Click()
Dim buf As String
If opcion1 = "0" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,Seccion from Pseccion "
   Else
   buf = "select Descripcio,Seccion from Pseccion  where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = buf
               Data1.Refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer.SetFocus
                  Exit Sub
               End If
               DBGrid1.columns(0).Width = 4000
               DBGrid1.columns(1).Width = 2000
               DBGrid1.SetFocus


End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = "0" Then
      Data2.Recordset.Edit
      Data2.Recordset.Fields("seccion") = "" & Data1.Recordset.Fields("seccion")
      Data2.Recordset.Update
      Frame1.Visible = False
      dbgrid2.SetFocus
   End If
End If
   



End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   If "" & Data2.Recordset.Fields("cierre") <> "C" Then
      consulta_seccion
   End If
End If
End Sub

Private Sub djiowewe_Click()
On Error GoTo cmd45_err
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
If Len("" & Data2.Recordset.Fields("producto")) > 0 Then
   If Len("" & Data2.Recordset.Fields("nro")) = 0 Then
      MsgBox "Ingresar el numero de Formula", 48, "Aviso"
      Exit Sub
   End If
   'ingresar insumos
   rproducc.VIENE = "S"
   rproducc.numero = "" & Data2.Recordset.Fields("numero")
   rproducc.nro = "" & Data2.Recordset.Fields("nro")
   rproducc.tarjeta = "" & Data2.Recordset.Fields("tarjeta")
   rproducc.cantidad = "" & Data2.Recordset.Fields("cantidad")
   rproducc.xlinea = "" & Data2.Recordset.Fields("linea")
   rproducc.producto = "" & Data2.Recordset.Fields("producto")
   rproducc.descripcio = "" & Data2.Recordset.Fields("descripcio")
   rproducc.xt1 = "" & Data2.Recordset.Fields("t1")
   rproducc.xt2 = "" & Data2.Recordset.Fields("t2")
   rproducc.xt3 = "" & Data2.Recordset.Fields("t3")
   rproducc.xt4 = "" & Data2.Recordset.Fields("t4")
   rproducc.xt5 = "" & Data2.Recordset.Fields("t5")
   rproducc.xt6 = "" & Data2.Recordset.Fields("t6")
   rproducc.xt7 = "" & Data2.Recordset.Fields("t7")
   rproducc.xt8 = "" & Data2.Recordset.Fields("t8")
   rproducc.xt9 = "" & Data2.Recordset.Fields("t9")
   rproducc.xt10 = "" & Data2.Recordset.Fields("t10")
   rproducc.xt11 = "" & Data2.Recordset.Fields("t11")
   rproducc.xt12 = "" & Data2.Recordset.Fields("t12")
   rproducc.xt13 = "" & Data2.Recordset.Fields("t13")
   rproducc.xt14 = "" & Data2.Recordset.Fields("t14")
   rproducc.xt15 = "" & Data2.Recordset.Fields("t15")
   rproducc.xt16 = "" & Data2.Recordset.Fields("t16")
   rproducc.Show 1
End If
Exit Sub
cmd45_err:
Exit Sub

End Sub

Private Sub dki3434_Click()
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
End Sub

Private Sub dmui821_Click()
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
Frame2.Visible = True
fechai.SetFocus

End Sub

Private Sub Form_Activate()

fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
fechaf = Format(Now, "dd/mm/yyyy")
carga_iniciales
sql_partes

End Sub
Sub carga_iniciales()

Dim mytablex As Table
seccion.Clear
seccion.AddItem "%"

Set mytablex = mydbxglo.OpenTable("pseccion")
Do
If mytablex.EOF Then Exit Do
seccion.AddItem "" & mytablex.Fields("seccion")
mytablex.MoveNext
Loop
seccion.ListIndex = 0
mytablex.Close
 
End Sub

Sub proceso_grabacion(sw As Integer)

If sw = 0 Then
   actualiza_procesos "" & Data2.Recordset.Fields("numero"), "" & Data2.Recordset.Fields("producto"), -1
   Data2.Recordset.Edit
   Data2.Recordset.Fields("cierre") = "C"
   Data2.Recordset.Fields("fechac") = Format(Now, "dd/mm/yyyy")
   Data2.Recordset.Update
   
End If
If sw = 1 Then
   actualiza_procesos "" & Data2.Recordset.Fields("numero"), "" & Data2.Recordset.Fields("producto"), 1
   Data2.Recordset.Edit
   Data2.Recordset.Fields("cierre") = ""
   Data2.Recordset.Fields("fechac") = Null
   Data2.Recordset.Update
   
End If

End Sub


Private Sub kfdi343_Click()
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub

tconetiq.producto = "" & Data2.Recordset.Fields("producto")
tconetiq.descripcio = "" & Data2.Recordset.Fields("descripcio")
tconetiq.unidad = "" & Data2.Recordset.Fields("unidad")
tconetiq.factor = "" & Data2.Recordset.Fields("factor")
tconetiq.linea = "" & Data2.Recordset.Fields("linea")
tconetiq.plano = "" & Data2.Recordset.Fields("numero")
tconetiq.tarjeta = "" & Data2.Recordset.Fields("tarjeta")
tconetiq.cantidad = "" & Data2.Recordset.Fields("cantidad")
tconetiq.T1 = "" & Data2.Recordset.Fields("t1")
tconetiq.t2 = "" & Data2.Recordset.Fields("t2")
tconetiq.t3 = "" & Data2.Recordset.Fields("t3")
tconetiq.t4 = "" & Data2.Recordset.Fields("t4")
tconetiq.t5 = "" & Data2.Recordset.Fields("t5")
tconetiq.t6 = "" & Data2.Recordset.Fields("t6")
tconetiq.t7 = "" & Data2.Recordset.Fields("t7")
tconetiq.t8 = "" & Data2.Recordset.Fields("t8")
tconetiq.t9 = "" & Data2.Recordset.Fields("t9")
tconetiq.t10 = "" & Data2.Recordset.Fields("t10")
tconetiq.t11 = "" & Data2.Recordset.Fields("t11")
tconetiq.t12 = "" & Data2.Recordset.Fields("t12")
tconetiq.t13 = "" & Data2.Recordset.Fields("t13")
tconetiq.t14 = "" & Data2.Recordset.Fields("t14")
tconetiq.t15 = "" & Data2.Recordset.Fields("t15")
tconetiq.t16 = "" & Data2.Recordset.Fields("t16")


tconetiq.Show 1
End Sub

Private Sub ldfo33_Click()
If Frame2.Visible = True Then
   Frame2.Visible = False
   dbgrid2.SetFocus
   Exit Sub
End If
If opcion1 = "0" Then
If Frame1.Visible = True Then
   Frame1.Visible = False
   dbgrid2.SetFocus
   Exit Sub
End If
End If
tpartes.Hide
Unload tpartes
End Sub
Sub sql_partes()
On Error GoTo cmd37_err
Dim buf As String
If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub
buf = "select * from dproduCc where "
buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
If seccion <> "%" Then
   buf = buf & " and seccion='" & seccion & "'"
End If
If plano <> "%" Then
   buf = buf & " and numero like '" & plano & "'"
End If
If tarjeta <> "%" Then
   buf = buf & " and tarjeta like '" & tarjeta & "'"
End If
buf = buf & " order by Fecha,val(tarjeta)"
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
               dbgrid2.SetFocus
Exit Sub
cmd37_err:
MsgBox "Error en select " & error$, 48, "Aviso"
Exit Sub
End Sub

Private Sub Nueo1_Click()
End Sub

Private Sub mo89wew_Click()
End Sub

Private Sub sdki2zom_Click()
End Sub
Sub consulta_seccion()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Seccion"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "0"
Command1_Click

End Sub
Sub actualiza_procesos(buf As String, buf1 As String, signo As Double)
Dim mytablex As Table
Dim mytably As Table
Dim sdx As Double
Dim xbod1 As String
Dim xbod2 As String
Set mytablex = mydbxglo.OpenTable("cproducc")  'buscamos principal
mytablex.Index = "cproducc"
mytablex.Seek "=", buf
If mytablex.NoMatch Then
   mytablex.Close
   Exit Sub
End If
xbod1 = "" & mytablex.Fields("bodegai")
xbod2 = "" & mytablex.Fields("bodega")
mytablex.Close

Set mytablex = mydbxglo.OpenTable("rproducc")  'buscamos las recetas
mytablex.Index = "rproducc"
mytablex.Seek "=", buf, buf1
If mytablex.NoMatch Then
   mytablex.Close
   Exit Sub
End If
Set mytabley = mydbxglo.OpenTable("almacen")
mytabley.Index = "almacen"
Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("numero") = buf And "" & mytablex.Fields("producto") = buf1 Then
   '--------------------------------------------------------------------
   mytabley.Seek "=", "01", "" & mytablex.Fields("productor"), xbod1 'act bodegas
   If Not mytabley.NoMatch Then
      mytabley.Edit
      sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad"))
      mytabley.Fields("saldo") = sdx
      mytabley.Update
   End If
   If mytabley.NoMatch Then
      mytabley.AddNew
      mytabley.Fields("producto") = "" & mytablex.Fields("productor")
      mytabley.Fields("bodega") = xbod1
      mytabley.Fields("local") = "01"
      sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad"))
      mytabley.Fields("saldo") = sdx
      mytabley.Update
   End If
   '--------------------------------------------------------------------
   Else: Exit Do
End If
mytablex.MoveNext
Loop
mytablex.Close
'aqui hacemos lo principal de lo fabricado
   mytabley.Seek "=", "01", "" & Data2.Recordset.Fields("producto"), xbod2
   If Not mytabley.NoMatch Then
      mytabley.Edit
      sdx = Val("" & mytabley.Fields("saldo")) - signo * Val("" & Data2.Recordset.Fields("cantidad"))
      mytabley.Fields("saldo") = sdx
      mytabley.Update
   End If
   If mytabley.NoMatch Then
      mytabley.AddNew
      mytabley.Fields("producto") = "" & Data2.Recordset.Fields("producto")
      mytabley.Fields("bodega") = xbod2
      mytabley.Fields("local") = "01"
      sdx = Val("" & mytabley.Fields("saldo")) - signo * Val("" & Data2.Recordset.Fields("cantidad"))
      mytabley.Fields("saldo") = sdx
      mytabley.Update
   End If
End Sub
Function graba_kardex(acu As String, xingreso As Double)
On Error GoTo cmd781_err
Dim mytablez As Table
Set mytablez = mydbxglo.OpenTable("detalle")
mytablez.AddNew
mytablez.Fields("estado") = "2"
mytablez.Fields("acu") = acu
If acu = "S" Then  'entrada
mytablez.Fields("tipo") = "E"
mytablez.Fields("cantidad") = xingreso
End If
If acu = "T" Then 'salidas
mytablez.Fields("cantidad") = xingreso
mytablez.Fields("tipo") = "S"
End If
mytablez.Fields("local") = "" & Data2.Recordset.Fields("local")
mytablez.Fields("serie") = ""
mytablez.Fields("numero") = "PR" & Data2.Recordset.Fields("numero")
mytablez.Fields("tipoclie") = "I"
mytablez.Fields("codigo") = "FABRICA"
mytablez.Fields("acu1") = ""
'mytablez.Fields("fecha") = Format(Now, "dd/mm/yyyy")
mytablez.Fields("moneda") = "S"
mytablez.Fields("producto") = "" & Data2.Recordset.Fields("producto")
mytablez.Fields("descripcio") = "" & Data2.Recordset.Fields("descripcio")
mytablez.Fields("unidad") = "" & Data2.Recordset.Fields("unidad")
mytablez.Fields("factor") = Val("" & Data2.Recordset.Fields("factor"))
mytablez.Fields("precio") = 0
mytablez.Fields("igv") = 19
mytablez.Fields("neto") = 0
mytablez.Fields("descuento") = 0
mytablez.Fields("subtotal") = 0
mytablez.Fields("impuesto") = 0
mytablez.Fields("total") = 0
mytablez.Fields("fecha") = Format("" & Data2.Recordset.Fields("fecha"), "dd/mm/yyyy")
mytablez.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
mytablez.Fields("hora") = Format(Now, "hh:mm:ss")
mytablez.Fields("vendedor") = ""
mytablez.Fields("bodega") = "" & Data2.Recordset.Fields("bodega")
mytablez.Fields("bodegaf") = ""
mytablez.Fields("deslipo") = 0
mytablez.Fields("flage") = ""
mytablez.Fields("linea") = "" & "" & Data2.Recordset.Fields("linea")
mytablez.Fields("t1") = 0
mytablez.Fields("t2") = 0
mytablez.Fields("t3") = 0
mytablez.Fields("t4") = 0
mytablez.Fields("t5") = 0
mytablez.Fields("t6") = 0
mytablez.Fields("t7") = 0
mytablez.Fields("t8") = 0
mytablez.Fields("t9") = 0
mytablez.Fields("t10") = 0
mytablez.Fields("t11") = 0
mytablez.Fields("t12") = 0
mytablez.Fields("t13") = 0
mytablez.Fields("t14") = 0
mytablez.Fields("t15") = 0
mytablez.Fields("t16") = 0
mytablez.Fields("l1") = ""
mytablez.Fields("l2") = ""
mytablez.Fields("l3") = ""
mytablez.Fields("l4") = ""
'mytablez.Fields("local") = ""
mytablez.Fields("proveedorp") = ""
mytablez.Fields("observa1") = ""
mytablez.Fields("observa2") = ""
mytablez.Fields("observa3") = ""
mytablez.Fields("observa4") = ""
mytablez.Fields("zona") = ""
mytablez.Fields("isc") = 0
mytablez.Fields("tax") = 0
mytablez.Fields("vtaneta") = 0
mytablez.Fields("tcosto") = 0
mytablez.Fields("ganancia") = 0
mytablez.Fields("comision") = 0
mytablez.Fields("usuario") = ""
mytablez.Fields("caja") = ""
mytablez.Fields("turno") = ""
mytablez.Fields("servicio") = ""
mytablez.Fields("comanda") = ""
mytablez.Fields("mesa") = ""
mytablez.Fields("salon") = ""
mytablez.Fields("mesero") = ""
'mytablez.Fields("local") = extra_loquesea(local1)
'MsgBox "x"
mytablez.Update
mytablez.Close
graba_kardex = 1
Exit Function
cmd781_err:
MsgBox "Error " + error$, 48, "Aviso"
Exit Function
End Function


