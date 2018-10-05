VERSION 5.00
Begin VB.Form tmenucli 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Osi"
   ClientHeight    =   8955
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   14070
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Acceso al Sistema"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   13935
      Begin VB.TextBox gempresa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   7680
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "1"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox clave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   7680
         MaxLength       =   11
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   1680
         Width           =   1575
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4920
         ScaleHeight     =   555
         ScaleWidth      =   8715
         TabIndex        =   20
         Top             =   480
         Width           =   8775
         Begin VB.Label Label2 
            BackColor       =   &H000000FF&
            Caption         =   "CONTROL DE ACCESO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2400
            TabIndex        =   21
            Top             =   120
            Width           =   3495
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   5040
         ScaleHeight     =   6435
         ScaleWidth      =   8595
         TabIndex        =   18
         Top             =   2280
         Width           =   8655
         Begin VB.Label Denomina 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2265
            Left            =   120
            TabIndex        =   19
            Top             =   4080
            Width           =   8400
         End
      End
      Begin VB.Image Image16 
         Height          =   945
         Left            =   10200
         Picture         =   "tmenucli.frx":0000
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Image Image15 
         BorderStyle     =   1  'Fixed Single
         Height          =   465
         Left            =   9360
         Picture         =   "tmenucli.frx":17AA
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SEDE"
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
         Left            =   5640
         TabIndex        =   14
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CLAVE"
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
         Left            =   5640
         TabIndex        =   15
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Image Image5 
         BorderStyle     =   1  'Fixed Single
         Height          =   8325
         Left            =   0
         Picture         =   "tmenucli.frx":1AB4
         Stretch         =   -1  'True
         Top             =   480
         Width           =   4845
      End
   End
   Begin VB.CommandButton image1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Tienda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      Picture         =   "tmenucli.frx":6D63
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton btnSalir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9840
      Picture         =   "tmenucli.frx":71A7
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salir"
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton image12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Asistencia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      Picture         =   "tmenucli.frx":75E9
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton image8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&ControlIngreso"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9840
      Picture         =   "tmenucli.frx":92DB
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton image4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Diagnostico"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      MaskColor       =   &H8000000E&
      Picture         =   "tmenucli.frx":AA39
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton image6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&CuentasxCobrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9840
      Picture         =   "tmenucli.frx":C8C3
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton image3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Consulta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      Picture         =   "tmenucli.frx":F40D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton image2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Productos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9840
      Picture         =   "tmenucli.frx":111CB
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Tratamiento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      Picture         =   "tmenucli.frx":1384D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&RecibosCaja"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9840
      Picture         =   "tmenucli.frx":16397
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   14010
      TabIndex        =   2
      Top             =   0
      Width           =   14070
      Begin VB.Label nusuario 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario Actual"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Image Image1b 
      BorderStyle     =   1  'Fixed Single
      Height          =   7125
      Left            =   120
      Picture         =   "tmenucli.frx":166A1
      Stretch         =   -1  'True
      Top             =   960
      Width           =   7845
   End
   Begin VB.Menu dlo2211 
      Caption         =   "&Tablas"
      Begin VB.Menu se223 
         Caption         =   "&0.Sede"
      End
      Begin VB.Menu cli81 
         Caption         =   "&1.Clinicas"
      End
      Begin VB.Menu em1 
         Caption         =   "&2.Empresas"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu dki22 
         Caption         =   "&3.Seguros"
      End
      Begin VB.Menu d78re 
         Caption         =   "&4.Referencias"
      End
      Begin VB.Menu dk221 
         Caption         =   "&5.Grupos"
      End
      Begin VB.Menu dkser4 
         Caption         =   "&6.Productos Servicios"
      End
      Begin VB.Menu l9911 
         Caption         =   "-"
      End
      Begin VB.Menu dkicli 
         Caption         =   "&7.Clientes/Empresas"
      End
      Begin VB.Menu mei911 
         Caption         =   "-"
      End
      Begin VB.Menu dkme81 
         Caption         =   "&8.Medicos"
      End
      Begin VB.Menu cmo992 
         Caption         =   "-"
      End
      Begin VB.Menu dlo222 
         Caption         =   "&A.Personal"
      End
      Begin VB.Menu buti8 
         Caption         =   "&B.Tipo Afiliado"
      End
      Begin VB.Menu tieau7 
         Caption         =   "&C.Tipo de Autorizacion"
      End
      Begin VB.Menu c433 
         Caption         =   "&D.Tipo de Consulta"
      End
      Begin VB.Menu dem44 
         Caption         =   "&E.Enfermedades"
      End
      Begin VB.Menu ff44 
         Caption         =   "-"
      End
      Begin VB.Menu dj784 
         Caption         =   "&F.Tipo Documento"
      End
      Begin VB.Menu djfpa5 
         Caption         =   "&G.Forma Pago"
      End
      Begin VB.Menu ff884 
         Caption         =   "-"
      End
      Begin VB.Menu xcia1 
         Caption         =   "&H.Caja"
      End
      Begin VB.Menu jntur1 
         Caption         =   "&I.Turno"
      End
      Begin VB.Menu dju7343 
         Caption         =   "&J.Transportista"
      End
      Begin VB.Menu djj7733 
         Caption         =   "&K.Almacenes"
      End
   End
   Begin VB.Menu dlo22 
      Caption         =   "&HistoriaClinica"
      Visible         =   0   'False
      Begin VB.Menu fi81 
         Caption         =   "&1.Fichas Cliente"
      End
      Begin VB.Menu f77 
         Caption         =   "-"
      End
      Begin VB.Menu kfdi11 
         Caption         =   "&2.Consultas/Atenciiones"
      End
      Begin VB.Menu f566 
         Caption         =   "-"
      End
      Begin VB.Menu dki33 
         Caption         =   "&3.Diagnostico"
      End
      Begin VB.Menu f891 
         Caption         =   "-"
      End
      Begin VB.Menu trat33 
         Caption         =   "&4.Tratamiento"
      End
      Begin VB.Menu as909 
         Caption         =   "-"
      End
      Begin VB.Menu dk8373 
         Caption         =   "&5.Asistencia"
      End
   End
   Begin VB.Menu djy6631 
      Caption         =   "&Tienda"
      Begin VB.Menu Cashre4 
         Caption         =   "&1.Caja registradora"
      End
   End
   Begin VB.Menu dkfac5 
      Caption         =   "&Ventas"
      Begin VB.Menu gnfac1 
         Caption         =   "&1.Facturacion"
      End
      Begin VB.Menu dfj33 
         Caption         =   "&2.Recibo Ingreso"
      End
      Begin VB.Menu fdk33 
         Caption         =   "&3.Recibo Egreso"
      End
      Begin VB.Menu fk883334 
         Caption         =   "&4.Cuadre de Caja"
      End
   End
   Begin VB.Menu dfkk22 
      Caption         =   "&Reportes"
      Visible         =   0   'False
   End
   Begin VB.Menu flo323 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tmenucli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub buti8_Click()
ttipoafi.Show 1
End Sub

Private Sub c433_Click()
ttipocon.Show 1
End Sub

Private Sub Cashre4_Click()
menucaja.Show 1
End Sub

Private Sub clave_KeyPress(KeyAscii As Integer)
Dim rsexiste As New ADODB.Recordset
Dim gsede As New ADODB.Recordset
If KeyAscii <> 13 Then Exit Sub
   
   If Len(clave) = 0 Then
      clave.SetFocus
      Exit Sub
   End If
   clave = UCase(clave)
   rsexiste.Open "SELECT sede,Nombre FROM sede where sede='" & Trim(gempresa) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount = 0 Then  'si existe
      MsgBox "No existe Sede ", 48, "Aviso"
      gempresa.SetFocus
      Exit Sub
   End If
   ngsede1 = Trim("" & rsexiste.Fields("nombre").Value)
   gsede.Open "SELECT personal,Nombre FROM personal where clave='" & Trim(clave) & "'", cn, adOpenKeyset, adLockOptimistic
   If gsede.RecordCount = 0 Then  'si existe
      MsgBox "No existe Persona Autorizada ", 48, "Aviso"
      clave.SetFocus
      Exit Sub
   End If
   If Not gsede.EOF Then
      nusuario = "" & gsede.Fields("nombre")
      dgusuario = Trim("" & gsede.Fields("personal"))
   End If
   dlo2211.Visible = True
   dlo22.Visible = True
   dfkk22.Visible = True
   Frame2.Visible = False
   dkfac5.Visible = True
   gsede1 = "" & gempresa
   
   globaldat = "\osi"
   gsede.Close
   rsexiste.Close
   
End Sub

Private Sub cli81_Click()
tclinica.Show 1
End Sub

Private Sub Command1_Click()
ttratame.Show 1
End Sub

Private Sub d78re_Click()
treferen.Show 1
End Sub

Private Sub dem44_Click()
tenferme.Show 1
End Sub

Private Sub dfj33_Click()
trecibo.Show 1
End Sub

Private Sub dj784_Click()
ttipodoc.Show 1
End Sub

Private Sub djfpa5_Click()
tfpago.Show 1
End Sub

Private Sub djj7733_Click()
tbodega.Show 1
End Sub

Private Sub dju7343_Click()
ttranspo.Show 1
End Sub

Private Sub dk221_Click()
tgrupo.Show 1
End Sub

Private Sub dk672_Click()

End Sub

Private Sub dk333_Click()

End Sub

Private Sub dk8373_Click()
tasiste.Show 1
End Sub

Private Sub dki22_Click()
tseguro.Show 1
End Sub

Private Sub dki33_Click()
'tconsult.ahyy1.Visible = False
'tconsult.dmi22.Visible = False
'tconsult.dfj8221.Visible = False
'tconsult.dk281.Visible = False
tdiagnos.Show 1
End Sub

Private Sub dkicli_Click()
tficha.Show 1
End Sub

Private Sub dkico3_Click()

End Sub

Private Sub dkme81_Click()
tmedico.Show 1
End Sub

Private Sub dkser4_Click()
tproduct.Show 1
End Sub

Private Sub dlo222_Click()
tpersona.Show 1
End Sub

Private Sub em1_Click()
tempresa.Show 1
End Sub

Private Sub fi81_Click()
tficha.Show 1
End Sub

Private Sub fk8833_Click()
End Sub

Private Sub flo323_Click()
tmenucli.Hide
Unload tmenucli
End Sub

Private Sub Form_Load()
Dim found As Integer
found = conectar()
If found = 0 Then
   MsgBox "Error de Conexion Sql Server ", 48, "Aviso"
   End
   Exit Sub
End If

End Sub

Private Sub gempresa_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
clave.SetFocus
End Sub

Private Sub gnfac1_Click()
tfactura.Show 1
End Sub

Private Sub image10_Click()
gnfac1_Click
End Sub

Private Sub image1_Click()
menucaja.Show 1
End Sub

Private Sub image12_Click()
tasiste.Show 1
End Sub

Private Sub Image15_Click()
clave_KeyPress 13
End Sub

Private Sub image2_Click()
tproduct.Show 1
End Sub

Private Sub image3_Click()
tconsult.dii33.Visible = False
tconsult.Show 1
End Sub

Private Sub image4_Click()
tdiagnos.Show 1

End Sub

Private Sub jntur1_Click()
tturno.Show 1
End Sub

Private Sub kfdi11_Click()
tconsult.dii33.Visible = False
tconsult.Show 1
End Sub

Private Sub se223_Click()
tsede.Show 1
End Sub

Private Sub tieau7_Click()
ttipoaut.Show 1
End Sub

Private Sub trat33_Click()
ttratame.Show 1
End Sub

Private Sub xcia1_Click()
tcaja.Show 1
End Sub
