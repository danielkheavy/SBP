VERSION 5.00
Begin VB.Form trepvta 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menu Reportes Ventas Ticket"
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
         Picture         =   "trepvta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Menu fdlk44 
      Caption         =   "&Reportes"
      Begin VB.Menu dj88 
         Caption         =   "&1.Cuadres de Caja"
      End
      Begin VB.Menu uni844 
         Caption         =   "&2.Unidades vendidas"
      End
      Begin VB.Menu dki8 
         Caption         =   "&3.Documentos Emitidos"
      End
      Begin VB.Menu fki44 
         Caption         =   "&4.Productos vs Documentos"
      End
      Begin VB.Menu fk8i33 
         Caption         =   "&5.Recibos Ingreso /egreso/seccion"
      End
      Begin VB.Menu dfk833 
         Caption         =   "&6.Copia Cierre de Cajas"
      End
   End
   Begin VB.Menu fdlo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trepvta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalir_Click()
fdlo44_Click
End Sub

Private Sub dfk833_Click()
   
Dim sw As Integer
Dim found As Integer

    opcion1 = "5"
    opcion2 = "1"
    opcion3 = "2"
    
    
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    
    usuariopos = gusuario
    tcuadrc1.tipoexterno.Visible = True
    tcuadrc1.numcuadre.Visible = True
    'tcuadrc1.flagdiario = "1"
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.Caption = "COPIA CIERRE DEL DIA"
    tcuadrc1.Show 1

End Sub

Private Sub dj88_Click()
Dim sw As Integer
Dim found As Integer


    
    opcion2 = "1"
    opcion1 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True

    usuariopos = gusuario
    tcuadrc1.flagdiario = "1"
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.Caption = "CUADRE PARCIAL DEL DIA"
    tcuadrc1.Show 1

End Sub

Private Sub Image1_Click()

End Sub

Private Sub dki8_Click()
Dim sw As Integer

    
    opcion1 = "2"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.flagdiario = "1"

    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.Caption = "DOCUMENTOS EMITIDOS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

End Sub

Private Sub fdlo44_Click()
trepvta.Hide
Unload trepvta
End Sub

Private Sub fk8i33_Click()
Dim sw As Integer

    
    opcion1 = "20"
    opcion2 = "2"
    opcion3 = ""
    tcuadrc1.flagdiario = "1"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.Caption = "DOCUMENTOS EMITIDOS PERIODICO"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1
    tcuadrc1.flagdiario = ""
    'Set mydbxglo = OpenDatabase(globaldir, False, False, "foxpro 2.5;")

End Sub

Private Sub fki44_Click()
Dim sw As Integer

    
    opcion1 = "4"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.flagdiario = "1"

    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.Caption = "PRODUCTOS VS DOCUMENTOS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

End Sub

Private Sub uni844_Click()
Dim sw As Integer
    
    opcion1 = "3"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.flagdiario = "1"
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.Caption = "UNIDADES VENDIDAS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

End Sub
