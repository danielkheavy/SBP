VERSION 5.00
Begin VB.Form trepevta 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Estadisticas de Ventas"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   10260
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      Begin VB.CommandButton btnSalir 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9000
         Picture         =   "trepevta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Menu fdlk44 
      Caption         =   "&Reportes"
      Begin VB.Menu ny44 
         Caption         =   "&1.Venta"
      End
      Begin VB.Menu dflo344 
         Caption         =   "&2.Ventas Mensuales"
      End
      Begin VB.Menu dki444 
         Caption         =   "&3.Ranking Productos"
      End
      Begin VB.Menu fo444 
         Caption         =   "&4.Graficos"
      End
   End
   Begin VB.Menu fdlo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trepevta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnsalir_Click()
fdlo44_Click
End Sub

Private Sub dj88_Click()

End Sub

Private Sub dflo344_Click()
opcion2 = "12"   'analisis de ventas
cgusuario = "FACTURA"
dgusuariog = "DETALLE"
'repdocum.Label18.Visible = False
'repdocum.Combo1.Visible = False
repdocum.vdetalle.Enabled = False
repdocum.vfpago.Enabled = False
repdocum.acu = "V"
repdocum.Show 1

End Sub

Private Sub dki444_Click()

opcion2 = "2"
repraped.Label12.Visible = True
repraped.orden.Visible = True
repraped.acu = "V" 'PEDIDO
repraped.xdata = "DETALLE"
repraped.Show 1

End Sub

Private Sub fdlo44_Click()
trepevta.Hide
Unload trepevta
End Sub

Private Sub fo444_Click()
FrmChart.acu = "V"
FrmChart.Show 1

End Sub

Private Sub ny44_Click()
opcion2 = "10"   'analisis de ventas
cgusuario = "FACTURA"
dgusuariog = "DETALLE"
'repdocum.Label18.Visible = False
'repdocum.Combo1.Visible = False
repdocum.vdetalle.Enabled = False
repdocum.vfpago.Enabled = False
repdocum.acu = "V"
repdocum.Show 1

End Sub
