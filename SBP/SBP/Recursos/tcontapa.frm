VERSION 5.00
Begin VB.Form tcontapa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametro de cuentas  "
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Height          =   855
      Left            =   7320
      TabIndex        =   49
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox ctapagocaja2 
      Height          =   375
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   43
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox ctapagocaja1 
      Height          =   375
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   42
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox ctacajabancoorigen2 
      Height          =   375
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   40
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox ctacajabancoorigen1 
      Height          =   375
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   39
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox ctacajabanco 
      Height          =   375
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   38
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox ctahonorigen 
      Height          =   375
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   33
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox ctacoborigen 
      Height          =   375
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   32
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox ctacomorigen 
      Height          =   375
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   31
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox ctahondolar 
      Height          =   375
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   30
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox ctacobdolar 
      Height          =   375
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   29
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox ctacomdolar 
      Height          =   375
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   28
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox ctahonsoles 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   27
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox ctacobsoles 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   25
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox ctacomsoles 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   23
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox renta 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   21
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox honorario 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   19
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox igv 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   17
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox resultado 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   15
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox percepcion 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox igvhonorario 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox igvrenta 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox igvventa 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox igvcompra 
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label26 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cobranzas"
      Height          =   375
      Left            =   7320
      TabIndex        =   48
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label Label25 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pagos"
      Height          =   375
      Left            =   7320
      TabIndex        =   47
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label24 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Compras"
      Height          =   375
      Left            =   7320
      TabIndex        =   46
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ventas"
      Height          =   375
      Left            =   7320
      TabIndex        =   45
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Honorarios"
      Height          =   375
      Left            =   7320
      TabIndex        =   44
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label21 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pago con Caja"
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ctas Caja banco"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Origen"
      Height          =   375
      Left            =   5520
      TabIndex        =   36
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dolares"
      Height          =   375
      Left            =   3840
      TabIndex        =   35
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Soles"
      Height          =   375
      Left            =   2160
      TabIndex        =   34
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Por Pagar"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label15 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Por Cobrar"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Por Pagar"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%Impuesto Renta"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%Impuesto Honorarios"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%Impuesto IGV"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resultado de Ejercicio"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta Percepcion"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Igv-"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta Impuesto Honorario"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta Impuesto ala Renta"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Igv-Cuenta Propia"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Igv-Cuenta Propia"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta Impuesto Venta"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuenta Impuesto Compra"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu flkio45 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcontapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    grabar

End Sub

Private Sub flkio45_Click()
    tcontapa.Hide
    Unload tcontapa

End Sub

Private Sub Form_Load()
    cargar

End Sub

Sub inicializa()
    igvcompra = ""
    igvventa = ""
    igvrenta = ""
    igvhonorario = ""
    percepcion = ""
    resultado = ""
    ctacomsoles = ""
    ctacomdolar = ""
    ctacomorigen = ""
    ctacobsoles = ""
    ctacobdolar = ""
    ctacoborigen = ""
    ctahonsoles = ""
    ctahondolar = ""
    ctahonorigen = ""
    ctacajabanco = ""
    ctacajabancoorigen1 = ""
    ctacajabancoorigen2 = ""
    ctapagocaja1 = ""
    ctapagocaja2 = ""
    igv = ""
    honorario = ""
    renta = ""

End Sub

Sub pone_registro(mytablex As ADODB.Recordset)
    igv = Trim("" & mytablex.Fields("igv"))
    honorario = Trim("" & mytablex.Fields("honorario"))
    renta = Trim("" & mytablex.Fields("renta"))

    igvcompra = Trim("" & mytablex.Fields("igvcompra"))
    igvventa = Trim("" & mytablex.Fields("igvventa"))
    igvrenta = Trim("" & mytablex.Fields("igvrenta"))
    igvhonorario = Trim("" & mytablex.Fields("igvhonorario"))
    percepcion = Trim("" & mytablex.Fields("percepcion"))
    resultado = Trim("" & mytablex.Fields("resultado"))
    ctacomsoles = Trim("" & mytablex.Fields("ctacomsoles"))
    ctacomdolar = Trim("" & mytablex.Fields("ctacomdolar"))
    ctacomorigen = Trim("" & mytablex.Fields("ctacomorigen"))
    ctacobsoles = Trim("" & mytablex.Fields("ctacobsoles"))
    ctacobdolar = Trim("" & mytablex.Fields("ctacobdolar"))
    ctacoborigen = Trim("" & mytablex.Fields("ctacoborigen"))
    ctahonsoles = Trim("" & mytablex.Fields("ctahonsoles"))
    ctahondolar = Trim("" & mytablex.Fields("ctahondolar"))
    ctahonorigen = Trim("" & mytablex.Fields("ctahonorigen"))
    ctacajabanco = Trim("" & mytablex.Fields("ctacajabanco"))
    ctacajabancoorigen1 = Trim("" & mytablex.Fields("ctacajabancoorigen1"))
    ctacajabancoorigen2 = Trim("" & mytablex.Fields("ctacajabancoorigen2"))
    ctapagocaja1 = Trim("" & mytablex.Fields("ctapagocaja1"))
    ctapagocaja2 = Trim("" & mytablex.Fields("ctapagocaja2"))

End Sub

Sub grabando(mytablex As ADODB.Recordset)
    mytablex.Fields("igv") = Trim(igv)
    mytablex.Fields("honorario") = Trim(honorario)
    mytablex.Fields("renta") = Trim(renta)

    mytablex.Fields("igvcompra") = Trim(igvcompra)
    mytablex.Fields("igvventa") = Trim(igvventa)
    mytablex.Fields("igvrenta") = Trim(igvrenta)
    mytablex.Fields("igvhonorario") = Trim(igvhonorario)
    mytablex.Fields("percepcion") = Trim(percepcion)
    mytablex.Fields("resultado") = Trim(resultado)
    mytablex.Fields("ctacomsoles") = Trim(ctacomsoles)
    mytablex.Fields("ctacomdolar") = Trim(ctacomdolar)
    mytablex.Fields("ctacomorigen") = Trim(ctacomorigen)
    mytablex.Fields("ctacobsoles") = Trim(ctacobsoles)
    mytablex.Fields("ctacobdolar") = Trim(ctacobdolar)
    mytablex.Fields("ctacoborigen") = Trim(ctacoborigen)
    mytablex.Fields("ctahonsoles") = Trim(ctahonsoles)
    mytablex.Fields("ctahondolar") = Trim(ctahondolar)
    mytablex.Fields("ctahonorigen") = Trim(ctahonorigen)
    mytablex.Fields("ctacajabanco") = Trim(ctacajabanco)
    mytablex.Fields("ctacajabancoorigen1") = Trim(ctacajabancoorigen1)
    mytablex.Fields("ctacajabancoorigen2") = Trim(ctacajabancoorigen2)
    mytablex.Fields("ctapagocaja1") = Trim(ctapagocaja1)
    mytablex.Fields("ctapagocaja2") = Trim(ctapagocaja2)

End Sub

Private Sub grba1_Click()

End Sub

Sub cargar()

    Dim mytablex As New ADODB.Recordset

    inicializa
    mytablex.Open "select * from cuentasparametro", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        pone_registro mytablex

    End If

    mytablex.Close

End Sub

Sub grabar()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from cuentasparametro", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        grabando mytablex
        mytablex.Update
    Else
        mytablex.AddNew
        grabando mytablex
        mytablex.Update

    End If

    mytablex.Close

End Sub
