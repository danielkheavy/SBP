VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form explocli 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorador Clientes"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   14235
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   14175
      TabIndex        =   0
      Top             =   0
      Width           =   14235
      Begin VB.ComboBox vendedor 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox clasifica 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox tipoclie 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   120
         Width           =   1815
      End
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
         Picture         =   "explocli.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explocli.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explocli.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Consulta"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explocli.frx":3636
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Borrar registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   6480
         MaxLength       =   11
         TabIndex        =   3
         Text            =   "%"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox nombre 
         Height          =   375
         Left            =   6480
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "%"
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11160
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explocli.frx":4848
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vended"
         Height          =   375
         Left            =   3240
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clasifica"
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoClie"
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   375
         Left            =   5760
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid dbGrid1 
      Height          =   7575
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   13361
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Nombre"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Ruc"
         Caption         =   "Ruc"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Dni"
         Caption         =   "Dni"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Extranjeria"
         Caption         =   "Extranjeria"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Tipoclie"
         Caption         =   "Tipoclie"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Clasifica"
         Caption         =   "Clase"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Telefono"
         Caption         =   "Telefono"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "vendedor"
         Caption         =   "Vendedor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   5114.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin VB.Menu dk232 
      Caption         =   "&Add"
   End
   Begin VB.Menu dkjiw3 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu dkw66 
      Caption         =   "&Borra"
   End
   Begin VB.Menu zom912 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu dj2323 
      Caption         =   "&Reporte"
      Begin VB.Menu celo3ex 
         Caption         =   "&0.Excell"
      End
      Begin VB.Menu dhne71 
         Caption         =   "&1.General"
      End
      Begin VB.Menu dfk88221 
         Caption         =   "&2.Generador"
      End
   End
   Begin VB.Menu fdoo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "explocli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub celo3ex_Click()
menu_excell
End Sub

Private Sub cmdDelete_Click()
dkw66_Click
End Sub

Private Sub cmdExit_Click()
fdoo232_Click
End Sub

Private Sub cmdPrint_Click()
trepclie.Show 1
End Sub

Private Sub cmdSort_Click()
zom912_Click
End Sub

Private Sub Command1_Click()
sql_clientes 1
End Sub

Private Sub Command2_Click()
fdoo232_Click
End Sub

Private Sub Command3_Click()
fdoo232_Click
End Sub


Private Sub DBGrid1_DblClick()
Dim buf As String
On Error GoTo cmd435_err
buf = "" & dbclie.Fields("codigo")
dk88221_Click
Exit Sub
cmd435_err:
Exit Sub
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
Dim buf As String
Dim buf2 As String
If KeyAscii <> 13 And KeyAscii <> 27 Then
         If KeyAscii = 8 Then
            If Len(nombre) > 0 Then
               buf = Mid$(nombre, 1, Len(nombre) - 1)
               nombre = buf
               KeyAscii = 0
               Else
               KeyAscii = 0
               Exit Sub
            End If
         End If
         buf = Chr(KeyAscii)
         If Chr(KeyAscii) = "*" Then
            buf = ""
            nombre = buf
         End If
         If KeyAscii <> 13 Then
            nombre = nombre + buf
         End If
         buf = nombre
         sql_clientes 0
End If
End Sub

Private Sub DBGrid2_Click()

End Sub

Private Sub DBGrid3_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If ColIndex <> 4 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 4
            If Not IsNumeric(dbGrid3.Columns(3)) Then
               Cancel = True
               Exit Sub
            End If
End Select
      
End Sub

Private Sub dfk88221_Click()
reporgen.NAMETABLA = "clientes"
reporgen.Show 1

End Sub

Private Sub dhne71_Click()


trepclie.Show 1
End Sub

Private Sub dk232_Click()


tclieta.moneda = "S"
tclieta.Caption = "NUEVO"
tclieta.Show 1
End Sub

Private Sub dk88221_Click()
End Sub

Private Sub dk892321_Click()

End Sub

Private Sub dkjiw3_Click()
On Error GoTo cmd4_err
tclieta.Caption = "MODIFICA"

tclieta.profesion = Trim("" & dbclie.Fields("profesion"))
tclieta.religion = Trim("" & dbclie.Fields("religion"))
tclieta.nrodepe = Trim("" & dbclie.Fields("nrodepe"))
tclieta.Trabajo = Trim("" & dbclie.Fields("trabajo"))
tclieta.cargo = Trim("" & dbclie.Fields("cargo"))
tclieta.hobbie = Trim("" & dbclie.Fields("hobbie"))
tclieta.civil = Trim("" & dbclie.Fields("civil"))
tclieta.tipovive = Trim("" & dbclie.Fields("tipovive"))


tclieta.barras = Trim("" & dbclie.Fields("barras"))
tclieta.ruc = Trim("" & dbclie.Fields("ruc"))
tclieta.dni = Trim("" & dbclie.Fields("dni"))
tclieta.especial = Trim("" & dbclie.Fields("especial"))
tclieta.clasifica = Trim("" & dbclie.Fields("clasifica"))
tclieta.tipoclie = Trim("" & dbclie.Fields("tipoclie"))

tclieta.zona = Trim("" & dbclie.Fields("zona"))
tclieta.lunes.Value = Val("" & dbclie.Fields("lunes"))
tclieta.martes.Value = Val("" & dbclie.Fields("martes"))
tclieta.miercoles.Value = Val("" & dbclie.Fields("miercoles"))
tclieta.jueves.Value = Val("" & dbclie.Fields("jueves"))
tclieta.viernes.Value = Val("" & dbclie.Fields("viernes"))
tclieta.sabado.Value = Val("" & dbclie.Fields("sabado"))
tclieta.domingo.Value = Val("" & dbclie.Fields("domingo"))
tclieta.fechalta = Trim("" & dbclie.Fields("fechanac"))
tclieta.referencias = Trim("" & dbclie.Fields("observa"))
tclieta.referencia = Trim("" & dbclie.Fields("referencia"))
tclieta.garantia = Trim("" & dbclie.Fields("garantia"))
tclieta.flete = Trim("" & dbclie.Fields("flete"))
tclieta.moneda = Trim("" & dbclie.Fields("moneda"))
tclieta.descuento1 = Trim("" & dbclie.Fields("descuento1"))
tclieta.credito = Trim("" & dbclie.Fields("credito"))
tclieta.vendedor = Trim("" & dbclie.Fields("vendedor"))
tclieta.descuento = Trim("" & dbclie.Fields("descuento"))
tclieta.diapago = Trim("" & dbclie.Fields("diapago"))
tclieta.fpago = Trim("" & dbclie.Fields("fpago"))
tclieta.cuenta = Trim("" & dbclie.Fields("cuenta"))
tclieta.codigo = Trim("" & dbclie.Fields("codigo"))
tclieta.codigo1 = Trim("" & dbclie.Fields("extranjeria"))
tclieta.nombre = Trim("" & dbclie.Fields("nombre"))
tclieta.nombrec = Trim("" & dbclie.Fields("nombrec"))
tclieta.contacto = Trim("" & dbclie.Fields("contacto"))
tclieta.direccion = Trim("" & dbclie.Fields("direccion"))
tclieta.dpto = Trim("" & dbclie.Fields("dpto"))
tclieta.distrito = Trim("" & dbclie.Fields("distrito"))
tclieta.telefono = Trim("" & dbclie.Fields("telefono"))
tclieta.telefono1 = Trim("" & dbclie.Fields("telefono1"))
tclieta.telefono2 = Trim("" & dbclie.Fields("telefono2"))
tclieta.correo = Trim("" & dbclie.Fields("correo"))
tclieta.estado = Trim("" & dbclie.Fields("estado"))
tclieta.codigo.Enabled = False
tclieta.Show 1
Exit Sub
cmd4_err:
MsgBox "Seleccione un registro " + error$, 48, "Aviso"
Exit Sub
End Sub

Private Sub dkw66_Click()
On Error GoTo cmd7812_err
If MsgBox("Desea Borrar " + "" & dbclie.Fields("nombre"), 1, "Aviso") <> 1 Then Exit Sub
cn.Execute ("delete from clientes where codigo='" & dbclie.Fields("codigo") & "'")
Command1_Click
Exit Sub
cmd7812_err:
MsgBox "Seleccione un registro", 48, "Aviso"
Exit Sub

End Sub

Private Sub fdoo232_Click()

explocli.Hide
Unload explocli
End Sub
Sub sql_clientes(sw As Integer)
On Error GoTo cmd37_err

Dim buf As String
buf = "select * from clientes where codigo like '" & codigo & "'"
If nombre <> "%" Then
   buf = buf & " and nombre like '" & nombre & "%'"
End If
If tipoclie <> "%" Then
   buf = buf & " and tipoclie like '" & extra_loquesea(tipoclie) & "'"
End If
If clasifica <> "%" Then
   buf = buf & " and clasifica like '" & extra_loquesea(clasifica) & "'"
End If
If vendedor <> "%" Then
   buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"
End If




   If dbclie.State = 1 Then dbclie.Close
   dbclie.Open buf, cn, adOpenStatic, adLockOptimistic
   Set dbGrid1.DataSource = dbclie
Exit Sub
cmd37_err:
MsgBox "Error en sql Clientes " & error$, 48, "Aviso"
Exit Sub
End Sub

Private Sub Form_Activate()
sql_clientes 1
End Sub

Private Sub Form_Load()
Dim mytablex As New ADODB.Recordset
tipoclie.Clear
tipoclie.AddItem "%"
   mytablex.Open "SELECT * FROM tipoclie ", cn, adOpenKeyset, adLockOptimistic
   Do
     If mytablex.EOF Then Exit Do
       tipoclie.AddItem "" & mytablex.Fields("tipoclie") & "|" & mytablex.Fields("descripcio")
       mytablex.MoveNext
   Loop
   tipoclie.ListIndex = 0
   mytablex.Close
   
clasifica.Clear
clasifica.AddItem "%"
   mytablex.Open "SELECT * FROM clasifi ", cn, adOpenKeyset, adLockOptimistic
   Do
     If mytablex.EOF Then Exit Do
       clasifica.AddItem "" & mytablex.Fields("clasifica") & "|" & mytablex.Fields("descripcio")
       mytablex.MoveNext
   Loop
   clasifica.ListIndex = 0
   mytablex.Close
   
vendedor.Clear
vendedor.AddItem "%"
   mytablex.Open "SELECT * FROM vendedor ", cn, adOpenKeyset, adLockOptimistic
   Do
     If mytablex.EOF Then Exit Do
       vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
       mytablex.MoveNext
   Loop
   vendedor.ListIndex = 0
   mytablex.Close
   
   
End Sub

Private Sub nombre_Change()
Command1_Click
End Sub

Private Sub zom912_Click()
On Error GoTo cmd3_err
tclieta.Caption = "ZOOM"
tclieta.dk23.Enabled = False
tclieta.cmdSave.Enabled = False
tclieta.barras = Trim("" & dbclie.Fields("barras"))
tclieta.ruc = Trim("" & dbclie.Fields("ruc"))
tclieta.dni = Trim("" & dbclie.Fields("dni"))
tclieta.especial = Trim("" & dbclie.Fields("especial"))
tclieta.clasifica = Trim("" & dbclie.Fields("clasifica"))
tclieta.tipoclie = Trim("" & dbclie.Fields("tipoclie"))

tclieta.zona = Trim("" & dbclie.Fields("zona"))
tclieta.lunes.Value = Val("" & dbclie.Fields("lunes"))
tclieta.martes.Value = Val("" & dbclie.Fields("martes"))
tclieta.miercoles.Value = Val("" & dbclie.Fields("miercoles"))
tclieta.jueves.Value = Val("" & dbclie.Fields("jueves"))
tclieta.viernes.Value = Val("" & dbclie.Fields("viernes"))
tclieta.sabado.Value = Val("" & dbclie.Fields("sabado"))
tclieta.domingo.Value = Val("" & dbclie.Fields("domingo"))
tclieta.fechalta = Trim("" & dbclie.Fields("fechanac"))
tclieta.referencias = Trim("" & dbclie.Fields("observa"))
tclieta.referencia = Trim("" & dbclie.Fields("referencia"))
tclieta.garantia = Trim("" & dbclie.Fields("garantia"))
tclieta.flete = Trim("" & dbclie.Fields("flete"))
tclieta.moneda = Trim("" & dbclie.Fields("moneda"))
tclieta.descuento1 = Trim("" & dbclie.Fields("descuento1"))
tclieta.credito = Trim("" & dbclie.Fields("credito"))
tclieta.vendedor = Trim("" & dbclie.Fields("vendedor"))
tclieta.descuento = Trim("" & dbclie.Fields("descuento"))
tclieta.diapago = Trim("" & dbclie.Fields("diapago"))
tclieta.fpago = Trim("" & dbclie.Fields("fpago"))
tclieta.cuenta = Trim("" & dbclie.Fields("cuenta"))
tclieta.codigo = Trim("" & dbclie.Fields("codigo"))
tclieta.codigo1 = Trim("" & dbclie.Fields("extranjeria"))
tclieta.nombre = Trim("" & dbclie.Fields("nombre"))
tclieta.nombrec = Trim("" & dbclie.Fields("nombrec"))
tclieta.contacto = Trim("" & dbclie.Fields("contacto"))
tclieta.direccion = Trim("" & dbclie.Fields("direccion"))
tclieta.dpto = Trim("" & dbclie.Fields("dpto"))
tclieta.distrito = Trim("" & dbclie.Fields("distrito"))
tclieta.telefono = Trim("" & dbclie.Fields("telefono"))
tclieta.telefono1 = Trim("" & dbclie.Fields("telefono1"))
tclieta.telefono2 = Trim("" & dbclie.Fields("telefono2"))
tclieta.correo = Trim("" & dbclie.Fields("correo"))
tclieta.estado = Trim("" & dbclie.Fields("estado"))
tclieta.codigo.Enabled = False

tclieta.Show 1
Exit Sub
cmd3_err:
MsgBox "Seleccione unh registro ", 48, "Aviso"
Exit Sub

End Sub
Sub menu_excell()
excel_paso
End Sub
Sub excel_paso()
Dim sdx As String
On Error GoTo cmd81_err
sdx = "" & dbclie("codigo")
conteo_excell
Exit Sub
cmd81_err:
MsgBox "Elegir un dato ", 48, "Aviso"
Exit Sub

End Sub

Sub conteo_excell()
 Dim v, h As Long
 
    Dim Heading(9) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
    On Error GoTo cmd1561212_err
   'Data1.Refresh
   
   
    Heading(1) = "Nombre"
    Heading(2) = "Codigo"
    Heading(3) = "barras"
    Heading(4) = "Dni"
    Heading(5) = "Ruc"
    Heading(6) = "Extranjeria"
    Heading(7) = "Direccion"
    Heading(8) = "Telefono"
    
   
    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excelcli(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    
v = 5
h = 1
dbclie.MoveFirst

     Do
            If dbclie.EOF Then Exit Do
            objExcel.ActiveSheet.Cells(v, h) = "'" & dbclie.Fields("Nombre")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & dbclie.Fields("codigo")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & dbclie.Fields("Barras")
            objExcel.ActiveSheet.Cells(v, h + 3) = "'" & dbclie.Fields("Dni")
            objExcel.ActiveSheet.Cells(v, h + 4) = "'" & dbclie.Fields("Ruc")
            objExcel.ActiveSheet.Cells(v, h + 5) = "'" & dbclie.Fields("Extranjeria")
            objExcel.ActiveSheet.Cells(v, h + 6) = "'" & dbclie.Fields("Direccion")
            objExcel.ActiveSheet.Cells(v, h + 7) = "'" & dbclie.Fields("Telefono")
            v = v + 1
            dbclie.MoveNext
     Loop
Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
Exit Sub
cmd1561212_err:
MsgBox "Error en exportacion excell " + error$, 48, "Aviso"
Exit Sub





End Sub
Public Function Formato_Excelcli(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 8)).Font.Bold = True
        
    For i = 1 To Num_Campos Step 1
        .Cells(3, i) = Nombre_Campos(i)
    Next i
        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
       
        .Columns("A").ColumnWidth = 35
        .Columns("B").ColumnWidth = 11
        .Columns("C").ColumnWidth = 11
        .Columns("D").ColumnWidth = 11
        .Columns("E").ColumnWidth = 11
        .Columns("F").ColumnWidth = 25
        .Columns("G").ColumnWidth = 10
        .Columns("H").ColumnWidth = 10
    
   
    
End With
End Function




