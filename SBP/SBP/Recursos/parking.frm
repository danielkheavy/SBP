VERSION 5.00
Begin VB.Form parking 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Parqueo"
   ClientHeight    =   7485
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   13440
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Parametro"
      Height          =   1485
      Left            =   4770
      TabIndex        =   23
      Top             =   2490
      Visible         =   0   'False
      Width           =   5505
      Begin VB.CommandButton Command11 
         Caption         =   "Salir"
         Height          =   375
         Left            =   3825
         TabIndex        =   27
         Top             =   750
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Graba"
         Height          =   375
         Left            =   3810
         TabIndex        =   26
         Top             =   285
         Width           =   1335
      End
      Begin VB.TextBox prehora 
         Height          =   495
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio x Hora"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   13380
      TabIndex        =   9
      Top             =   0
      Width           =   13440
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PARAMETRO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "    CIERRE     CAJA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CUADRE PARCIAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COPIA   FACTURAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ANULA   FACTURAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SALIDA VEHICULOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENTRADA VEHICULOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "EGRESO DINERO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "INGRESO DINERO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   8160
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   3795
      TabIndex        =   4
      Top             =   3840
      Width           =   3855
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PRECIO X HORA S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   600
         Width           =   2415
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Left            =   0
         Picture         =   "parking.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   2640
      Width           =   3855
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Left            =   0
         Picture         =   "parking.frx":0D1A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   915
      End
      Begin VB.Label hora 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   3
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   1440
      Width           =   3855
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   -360
         Picture         =   "parking.frx":1DB6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   915
      End
      Begin VB.Label fecha 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   1
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turno"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cajero"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   5895
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   9255
   End
   Begin VB.Label turno 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label caja 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label cajero 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Menu fdoo23 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "parking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If Frame1.Visible = True Then Exit Sub
    PARKEs.caja = caja
    PARKEs.turno = turno
    PARKEs.cajero = cajero
    PARKEs.valor = Label2

    PARKEs.Show 1

End Sub

Private Sub Command10_Click()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("prehora") = Val(prehora)
        mytablex.Update

    End If

    mytablex.Close

End Sub

Private Sub Command11_Click()
    fdoo23_Click

End Sub

Private Sub Command2_Click()

    If Frame1.Visible = True Then Exit Sub
    'PARKEE.Caption = "ENTRADA DE VEHICULO"
    PARKEE.Show 1

End Sub

Private Sub Command3_Click()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    If Frame1.Visible = True Then Exit Sub

    If Len(Trim("" & mytable11.Fields("tipore"))) = 0 Then
        MsgBox "No se ha definido el tipo recibo en parametros caja ", 48, "Aviso"
        Exit Sub

    End If

    mytablex.Open "SELECT * FROM tipo where  tipo='" & "" & mytable11.Fields("tipore") & "' and tipodoc='V'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        MsgBox "No existe tipo recibo definido en parametros Caja ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    mytablex.Close

    gofpago = "fpagov"
    found = copiar_recibos()

    If found = 0 Then
        MsgBox "Error al copiar archivo temporal", 24, "Aviso"
        Exit Sub

    End If

    fgusuario = "_l" & gusuario
    found = copiar_tmpfpagoR()

    If found = 0 Then
        MsgBox "No se puede copiar temporal tmpfpagor", 48, "Aviso"
        Exit Sub

    End If

    'found = copiar_recibos()
    'If found = 0 Then
    '   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
    '   End
    '   Exit Sub
    'End If
    'explreci.afecta = "P"  'proveedor
    'explreci.acu = "V"
    trecaja.tipo = "" & mytable11.Fields("tipore")
    trecaja.tipo.Enabled = False
    trecaja.pagocash.Visible = True
    trecaja.pagocash.Value = 1

    fgusuario = "_r" & gusuario
    trecaja.Combo2.Enabled = True
    trecaja.xcuentaco = "cuentaC"
    trecaja.XCUENTACO1 = "cuentaCd"
    trecaja.tipoclie = "C"

    trecaja.Caption = "EGRESO DINERO"
    trecaja.local1 = "" & mytable11.Fields("local")
    trecaja.serie = "" & mytable11.Fields("seriere")
    'trecaja.local1.Enabled = False
    trecaja.afecta = "C"
    trecaja.acu = "V"
    trecaja.cajero = cajero
    trecaja.caja = caja
    trecaja.turno = turno
    trecaja.fecha = dia
    trecaja.dia = dia
    trecaja.ch89343.Visible = True
    trecaja.d7823.Visible = True

    trecaja.fecha.Enabled = False
    trecaja.Show 1

End Sub

Private Sub Command4_Click()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    If Frame1.Visible = True Then Exit Sub

    If Len(Trim("" & mytable11.Fields("tipori"))) = 0 Then
        MsgBox "No se ha definido el tipo recibo en parametros caja ", 48, "Aviso"
        Exit Sub

    End If

    mytablex.Open "SELECT * FROM tipo where  tipo='" & "" & mytable11.Fields("tipori") & "' and tipodoc='W'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        MsgBox "No existe tipo recibo definido en parametros Caja ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    mytablex.Close
    found = copiar_recibos()

    If found = 0 Then
        MsgBox "Error al copiar archivo temporal", 24, "Aviso"
        Exit Sub

    End If

    fgusuario = "_l" & gusuario
    found = copiar_tmpfpagoR()

    If found = 0 Then
        MsgBox "No se puede copiar temporal tmpfpagor", 48, "Aviso"
        Exit Sub

    End If

    '

    fgusuario = "_r" & gusuario
    trecaja.pagocash.Visible = True
    trecaja.pagocash.Value = 1
    trecaja.Combo2.Enabled = True

    trecaja.xcuentaco = "cuentac"
    trecaja.XCUENTACO1 = "cuentacd"
    trecaja.tipoclie = "C"
    trecaja.tipo = "" & mytable11.Fields("tipori")
    trecaja.tipo.Enabled = False
    trecaja.Caption = "INGRESO DINERO"
    trecaja.afecta = "C"
    trecaja.local1 = Trim("" & mytable11.Fields("local"))
    trecaja.serie = "" & mytable11.Fields("serieri")
    trecaja.acu = "W"
    trecaja.cajero = cajero
    trecaja.caja = caja
    trecaja.turno = turno
    trecaja.fecha = dia
    trecaja.dia = dia
    trecaja.fecha.Enabled = False
    trecaja.ch89343.Visible = True
    trecaja.d7823.Visible = True
    trecaja.Show 1

End Sub

Private Sub Command5_Click()

    'Dim rrlocal11 As String
    'Dim rrtipo As String
    'Dim rrserie As String
    'Dim rrnumero As String
    If Frame1.Visible = True Then Exit Sub

    Dim found As Integer

    flag_clave1 = 0
    tconcla.X = "ANULA"  '
    tconcla.Show 1

    If flag_clave1 <> 1 Then  'si es descongela
        Exit Sub

    End If

    rrlocal11 = ""
    rrtipo = ""
    rrserie = ""
    rrnumero = ""
    mmenua.Caption = "ANULA"
    mmenua.Show 1

    If Len(rrlocal11) = 0 Then Exit Sub
    If Len(rrtipo) = 0 Then Exit Sub
    If Len(rrnumero) = 0 Then Exit Sub
    found = valida_otros()

    If found = 0 Then
        MsgBox "No existe Documento ", 48, "Aviso"
        Exit Sub

    End If

    anularr

End Sub

Private Sub Command6_Click()

    If Frame1.Visible = True Then Exit Sub

    Dim found As Integer

    flag_clave1 = 0
    tconcla.X = "COPIA"  '
    tconcla.Show 1

    If flag_clave1 <> 1 Then  'si es descongela
        'DBGrid2.SetFocus
        Exit Sub

    End If

    rrlocal11 = ""
    rrtipo = ""
    rrserie = ""
    rrnumero = ""
    mmenua.Caption = "COPIA"
    mmenua.Show 1

    If Len(rrlocal11) = 0 Then Exit Sub
    If Len(rrtipo) = 0 Then Exit Sub
    If Len(rrnumero) = 0 Then Exit Sub
    found = valida_otros()

    If found = 0 Then
        MsgBox "No existe Documento ", 48, "Aviso"
        Exit Sub

    End If

    proceso_impresion11 rrtipo, rrserie, rrnumero, 1, "1"

End Sub

Private Sub Command7_Click()

    If Frame1.Visible = True Then Exit Sub
    flag_clave1 = 0
    tconcla.X = "CUADRE"
    tconcla.Show 1

    If flag_clave1 <> 1 Then  'si es descongela
        'dbgrid2.SetFocus
        Exit Sub

    End If

    opcion2 = "1"
    opcion1 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True

    usuariopos = gusuario
    tcuadrc1.cajero = "" & cajero
    tcuadrc1.caja = "" & caja
    tcuadrc1.turno = "" & turno
    tcuadrc1.fechai = "" & dia
    tcuadrc1.fechaf = "" & dia
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.Caption = "CUADRE PARCIAL DEL DIA"
    tcuadrc1.Show 1

End Sub

Private Sub Command8_Click()

    Dim sw As Integer

    If Frame1.Visible = True Then Exit Sub

    flag_clave1 = 0
    tconcla.X = "CIERRE"
    tconcla.Show 1

    If flag_clave1 = 0 Then  'si es descongela
        'Label27_Click
        Exit Sub

    End If
    
    opcion1 = "5"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.cajero = usuariopos
    tcuadrc1.caja = caja
    tcuadrc1.turno = turno
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = dia
    tcuadrc1.fechaf = dia
    tcuadrc1.Caption = "CIERRE DEL DIA"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

End Sub

Private Sub Command9_Click()

    Dim mytablex As New ADODB.Recordset

    If Frame1.Visible = True Then Exit Sub
    prehora = ""
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        prehora = "" & mytablex.Fields("prehora")

    End If

    mytablex.Close
    Frame1.Visible = True
    prehora.SetFocus

End Sub

Private Sub fdoo23_Click()

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    parking.Hide
    Unload parking

End Sub

Private Sub Form_Load()
    carga_precio
    carga_imagen

End Sub

Sub carga_imagen()

    On Error GoTo cmd9000_err

    image5.Picture = LoadPicture(globalpath & "\ico\parkeo.jpg")
    Exit Sub
cmd9000_err:

End Sub

Private Sub Timer1_Timer()
    fecha = Format(Now, "dd/mm/yyyy")
    hora = Format(Now, "hh:mm:ss")

End Sub

Sub carga_precio()

    Dim mytablex As New ADODB.Recordset

    Label2 = ""
    'Set mytablex = mydbxglo.OpenTable("parame")
    'mytablex.Index = "codigo"
    'mytablex.Seek "=", "01"
    'If Not mytablex.NoMatch Then
    '   Label2 = "" & mytablex.Fields("prehora")
    'End If
    'mytablex.Close
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Label2 = "" & mytablex.Fields("prehora")

    End If

    mytablex.Close

End Sub

Function valida_otros()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from factura where local='" & rrlocal1 & "' and tipo='" & rrtipo & "' and serie='" & rrserie & "' and numero='" & rrnumero & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_otros = 1

    End If

    mytablex.Close

End Function

Sub anularr()

    Dim found As Integer

    found = proceso_anular(rrtipo, rrserie, rrnumero)

    If found = 1 Then
        proceso_impresion11 rrtipo, rrserie, rrnumero, 0, ""

    End If

    'DBGrid2.SetFocus
End Sub

Function proceso_anular(ytipo As String, yserie As String, ynumero As String)

    Dim sw       As Integer

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM " & gocabeza & " where  local='" & Trim("" & selocal.Text) & "' and tipo='" & ytipo & "' and serie='" & yserie & "' and numero='" & ynumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        'mytablex.Edit
        mytablex.Fields("estado") = "1"
        mytablex.Update

    End If

    mytablex.Close
    sw = 0

    If mytablex.State = 1 Then mytablex.Close
    'MsgBox godetalle
    mytablex.Open "SELECT * FROM " & godetalle & " where  local='" & Trim("" & selocal.Text) & "' and tipo='" & ytipo & "' and serie='" & yserie & "' and numero='" & ynumero & "'", cn, adOpenStatic, adLockOptimistic  'adOpenDynamic

    If mytablex.RecordCount > 0 Then 'si existe
        'found = descarga_saldo("" & selocal.Text, mytablex, ytipo, yserie, ynumero, 1, 1)
        mytablex.Close
        cn.Execute ("update " & godetalle & " set estado='1'" & " where  local='" & Trim("" & selocal.Text) & "' and tipo='" & ytipo & "' and serie='" & yserie & "' and numero='" & ynumero & "'")
        sw = 1

    End If

    If sw = 0 Then
        mytablex.Close

    End If

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM " & gofpago & " where  local='" & Trim("" & selocal.Text) & "' and tipo='" & ytipo & "' and serie='" & yserie & "' and numero='" & ynumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        Do

            If mytablex.EOF Then Exit Do
            'mytablex.Edit
            mytablex.Fields("estado") = "1"
            mytablex.Update

            If "" & mytablex.Fields("acufp") = "V" Then

                'graba_acumulado_clientes "" & mytablex.Fields("codigo"), -1, Val("" & mytablex.Fields("recibe"))
            End If

            'found = borra_credito(ytipo, yserie, ynumero)
            'If "" & mytablex.Fields("acufp") = "I" Then
            '  found = anula_tmpcta(mytablex)
            'End If
            'desgraba_deposito mytablex
      
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    cn.Execute ("DELETE FROM  cuentap where local='" & Trim("" & selocal.Text) & "' and tipo='" & Trim("" & ytipo) & "' and serie='" & Trim("" & yserie) & "' and numero='" & Trim("" & ynumero) & "' and cuota='1'")

    'reversa_guia_mensual Trim("" & selocal.Text), ytipo, yserie, ynumero
    proceso_anular = 1

End Function

Sub proceso_impresion11(bxtipo As String, _
                        bxserie As String, _
                        bxnumero As String, _
                        sw As Integer, _
                        ascopia As String)

    Dim found    As Integer

    Dim archivot As String

    On Error GoTo cmd6_err:

    'MsgBox ""
    cerrar_archivo

    If sw = 0 Then   'si es posible
        found = abre_puerto("" & mytable11.Fields("capuerto"), Val("" & mytable11.Fields("catipo")), "" & mytable11.Fields("gavetacola"))

    End If

    'verificamos si es puerto LPT para no hacer formato impresion
    found = control_impresion(bxtipo, 10)

    If found = 10 And sw <> 2 Then
        Exit Sub

    End If

    'MsgBox "proceso impresion"
    factura_formatox Trim("" & mytable11.Fields("local")), "" & bxtipo, "" & bxserie, "" & bxnumero, ascopia, sw
    cerrar_archivo
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = control_impresion(bxtipo, sw)
    'genver.File = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$, 48, "Aviso"
    Exit Sub

End Sub

Function control_impresion(bxtipo As String, psw As Integer)

    Dim found      As Integer

    Dim sFile      As String

    Dim mytablex   As New ADODB.Recordset

    Dim sw         As String

    Dim xcolax     As String

    Dim xxpuerto   As String

    Dim oldprinter As String

    On Error GoTo cmd67111_err

    sw = ""
    xcolax = ""
    xxpuerto = "X_"

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM tipo where  tipo='" & bxtipo & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe

        Select Case "" & mytablex.Fields("tipodoc")

            Case "A"
                xxpuerto = "" & mytable11.Fields("puertobm")
                sw = "" & mytable11.Fields("ibm")
                xcolax = "" & mytable11.Fields("cbm")

            Case "B"
                xxpuerto = "" & mytable11.Fields("puertofm")
                sw = "" & mytable11.Fields("ifm")
                xcolax = "" & mytable11.Fields("cfm")

            Case "C"
                xxpuerto = "" & mytable11.Fields("puertotb")
                sw = "" & mytable11.Fields("itb")
                xcolax = "" & mytable11.Fields("ctb")

            Case "D"
                xxpuerto = "" & mytable11.Fields("puertotf")
                sw = "" & mytable11.Fields("itf")
                xcolax = "" & mytable11.Fields("ctf")

            Case "G"
                xxpuerto = "" & mytable11.Fields("puertonv")
                sw = "" & mytable11.Fields("inv")
                xcolax = "" & mytable11.Fields("cnv")

            Case "H"
                xxpuerto = "" & mytable11.Fields("puertope")
                sw = "" & mytable11.Fields("ipe")
                xcolax = "" & mytable11.Fields("cpe")

            Case "I"  'pedidos
       
                xxpuerto = "" & mytable11.Fields("puertope")
                sw = "" & mytable11.Fields("ipe")
                xcolax = "" & mytable11.Fields("cpro")
       
            Case "T"
                xxpuerto = "" & mytable11.Fields("puertoot")
                sw = "" & mytable11.Fields("iot")
                xcolax = "" & mytable11.Fields("cpro")

            Case "1"
                xxpuerto = "" & mytable11.Fields("puertoexo")
                sw = "" & mytable11.Fields("iexo")
                xcolax = "" & mytable11.Fields("cexo")

        End Select

    End If

    mytablex.Close

    If psw = 10 Then  'solo es para ver si es LPT
        control_impresion = 11

        If xxpuerto = "LPT" Then
            control_impresion = 10

        End If

        Exit Function

    End If

    'found = abre_puerto("" & mytable11.Fields("capuerto"), Val("" & mytable11.Fields("catipo")))
    'ahora validamos los parametros de impresion

    If psw = 2 Then  'si  es orden de despacho
   
        If "" & mytable11.Fields("odcola") = "S" Then
      
            oldprinter = Printer.DeviceName
            selecciona_impresoras ("" & mytable11.Fields("odpuerto"))
            sFile = globaldir & "\temporal\" & gusuario & ".txt"
            found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
            selecciona_impresoras (oldprinter)

        End If

        If "" & mytable11.Fields("odcola") <> "S" Then
            'MsgBox "" & mytable11.Fields("odpuerto")
            found = star_sp342("" & mytable11.Fields("odpuerto"), 0)
            found = corte_papel("" & mytable11.Fields("odpuerto"), Val("" & mytable11.Fields("catipo")))

        End If

        control_impresion = found
        Exit Function

    End If

    If sw = "S" Then
        If MsgBox("Desea Imprimir", 1 + 256, "Aviso") <> 1 Then
            control_impresion = 1
            Exit Function

        End If

    End If

    If xcolax = "S" Then
        oldprinter = Printer.DeviceName
        selecciona_impresoras (xxpuerto)
        sFile = globaldir & "\temporal\" & gusuario & ".txt"
        found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
        selecciona_impresoras (oldprinter)

    End If

    If xcolax <> "S" Then
        found = star_sp342(xxpuerto, 0)
        found = corte_papel(xxpuerto, Val("" & mytable11.Fields("catipo")))

    End If

    control_impresion = found
    Exit Function
cmd67111_err:
    MsgBox "Aviso en control impresion " + error$, 48, "Aviso"
    Exit Function

End Function

Sub factura_formatox(bxlocal As String, _
                     bxtipo As String, _
                     bxserie As String, _
                     bxnumero As String, _
                     ascopia As String, _
                     psw As Integer)

    Dim vacu            As String

    Dim mytablex        As New ADODB.Recordset

    Dim mytabley        As New ADODB.Recordset

    Dim mytablez        As New ADODB.Recordset

    Dim found           As Integer

    Dim nro_lineas      As Integer

    Dim contando        As Integer

    Dim faltante        As Integer

    Dim I               As Integer

    Dim archivo_formato As String

    On Error GoTo cmd450009_err

    vacu = ""
    'MsgBox "QU"
       
    nro_lineas = busca_tipo_lineas(bxtipo)
    'MsgBox ""
    'If nro_lineas <= 0 Then
    '   nro_lineas = 10
    'End If
    'MsgBox ""
    contando = 0
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = borra_nombre("" & FileName)
       
    If psw = 2 Then 'si es de orden
        archivo_formato = "orden"
    Else
        'MsgBox bxtipo
        archivo_formato = busca_archivo_formato(bxtipo)

        If Len(archivo_formato) = 0 Then
            MsgBox "No existe archivo formato ", 48, "Aviso"
            'MsgBox ""
            Exit Sub

        End If

    End If

    'cabeza
    'proceso_formatos(archivo_formato , mydbx , mytablex , ubicacioni , ubicacionf , basedatos , indice , tipo , numero , ascopia , contando )
    mytablex.Open "SELECT * FROM " & gocabeza & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si existe
        mytablex.Close
        Exit Sub

    End If

    'MsgBox ""
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "{", "}", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "{", "}", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
        
    vacu = "" & mytablex.Fields("acu")
    'MsgBox ""
    '
    'detalle
    flag_contando = 0
    mytabley.Open "SELECT * FROM " & godetalle & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytabley.RecordCount > 0 Then 'si existe
        Do

            If mytabley.EOF Then Exit Do
            If "" & mytabley.Fields("dua") <> "R" Then
                flag_contando = contando + 1
                'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                'found = proceso_formatos(archivo_formato, mytabley, "/", "\", godetalle, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                found = proceso_formatos(archivo_formato, mytabley, "/", "\", godetalle, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                  
                contando = contando + 1

            End If
          
            mytabley.MoveNext
        Loop

    End If

    'mytabley.Close
    '
    If nro_lineas > 0 Then

        'If bxtipo <> "1" And bxtipo <> "2" And bxtipo <> "5" And bxtipo <> "7" Then
        If contando < nro_lineas Then

            For I = contando To nro_lineas
                Open FileName For Append As #1
                found = formateaa("", 1, 2, 0)
                Close #1
            Next I

        End If

    End If

    '----- SUBTOTAL
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "$", "?", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "$", "?", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
             
    mytablez.Open "SELECT * FROM " & gofpago & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablez.RecordCount > 0 Then 'si existe
        Do

            If mytablez.EOF Then Exit Do
            'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            'found = proceso_formatos(archivo_formato, mytablez, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            found = proceso_formatos(archivo_formato, mytablez, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            mytablez.MoveNext
        Loop

    End If

    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "^", "&", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "^", "&", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
           
    mytablex.Close
    mytabley.Close
    mytablez.Close
    Exit Sub
cmd450009_err:
    MsgBox "Aviso Proceso en Compilando " & error$, 24, "Aviso"
    'mytablex.Close
    Exit Sub

End Sub

Function busca_archivo_formato(bxtipo As String) As String

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    'MsgBox bxtipo
    mytablex.Open "SELECT * FROM tipo where  tipo='" & bxtipo & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe

        Select Case "" & mytablex.Fields("tipodoc")

            Case "Z" 'si es traslado
                busca_archivo_formato = "" & mytablex.Fields("archivo")

            Case "A"
                busca_archivo_formato = "" & mytable11.Fields("archivobm")

            Case "B"
                busca_archivo_formato = "" & mytable11.Fields("archivofm")

            Case "C"
                busca_archivo_formato = "" & mytable11.Fields("archivotb")

            Case "1"
                busca_archivo_formato = "" & mytable11.Fields("archivoexo")

            Case "D"
                busca_archivo_formato = "" & mytable11.Fields("archivotf")

            Case "G"
                busca_archivo_formato = "" & mytable11.Fields("archivonv")

            Case "H"
                busca_archivo_formato = "" & mytable11.Fields("archivope")

            Case "T"
                busca_archivo_formato = "" & mytable11.Fields("archivoot")

            Case "I"
                busca_archivo_formato = "" & mytable11.Fields("archivope")

                'MsgBox ""
        End Select

        'MsgBox ""
    End If

    mytablex.Close
 
End Function

Function busca_tipo_lineas(buf As String)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM tipo  where  tipo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        busca_tipo_lineas = Val("" & mytablex.Fields("nrolineas"))

        'MsgBox ""
    End If

    mytablex.Close

End Function

