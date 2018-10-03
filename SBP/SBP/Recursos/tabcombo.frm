VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form tabcombo 
   BackColor       =   &H00808080&
   Caption         =   "Tabla de Combos"
   ClientHeight    =   9555
   ClientLeft      =   165
   ClientTop       =   15
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   13935
      Begin VB.ComboBox Combo3 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox bufferx 
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   6735
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   11880
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   22
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin ChamaleonButton.ChameleonBtn ChaAceptar 
         Height          =   930
         Left            =   12480
         TabIndex        =   31
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1640
         BTYPE           =   5
         TX              =   "Aceptar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   255
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "tabcombo.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChaCerrar 
         Height          =   750
         Left            =   12480
         TabIndex        =   32
         Top             =   2400
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1323
         BTYPE           =   5
         TX              =   "Cerrar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421504
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "tabcombo.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn Command3 
         Height          =   495
         Left            =   10200
         TabIndex        =   30
         Top             =   480
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "Buscar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421504
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "tabcombo.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Frame2"
      Height          =   6615
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   12495
      Begin VB.CommandButton cmdBuscarProducto 
         Height          =   420
         Left            =   4200
         Picture         =   "tabcombo.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Buscar producto"
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox Cantidad 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   27
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.TextBox gcombina 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   19
         Top             =   1440
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox producto 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox descripcio 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         MaxLength       =   60
         TabIndex        =   14
         Top             =   600
         Width           =   6015
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   8400
         Picture         =   "tabcombo.frx":0802
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir todo"
         Top             =   3120
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   8400
         Picture         =   "tabcombo.frx":10CC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2160
         Width           =   1470
      End
      Begin VB.Label LblCantidad 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12435
      TabIndex        =   2
      Top             =   0
      Width           =   12495
      Begin VB.CommandButton cmdSave 
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
         Picture         =   "tabcombo.frx":1996
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
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
         Height          =   375
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Buscar"
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
         Left            =   10560
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0E0E0&
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
         Picture         =   "tabcombo.frx":2BA8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Borrar registro"
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
         Left            =   2760
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tabcombo.frx":3DBA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
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
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tabcombo.frx":4FCC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAddEntry 
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
         Picture         =   "tabcombo.frx":61DE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label xdescripcio 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label xproducto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   18
         Top             =   720
         Width           =   105
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   13996
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Descripciop"
            Caption         =   "Descripcio"
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
            DataField       =   "Productop"
            Caption         =   "Producto"
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
            DataField       =   "Cantidad"
            Caption         =   "Cantidad"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   5940.284
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu f8443 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu fjh433 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
      Begin VB.Menu dk9893 
         Caption         =   "&0.GENERAL"
      End
      Begin VB.Menu mnuArchivoArray 
         Caption         =   "Novisible"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tabcombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txempre As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    producto.Enabled = True
    producto = ""
    producto.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    buf = txempre.Fields("producto")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txempre.Fields("producto"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txempre.Delete
    Command1_Click

    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command1_Click

End Sub

Private Sub bufferx_Change()
    'Combos Combinaciones 07/07/2018
    ejecutax 1

    'Combos Combinaciones 07/07/2018
End Sub

'Combos Combinaciones 07/07/2018
Private Sub ChaAceptar_Click()

    If opcion1 = "1" Then
     
        Dim rbusca As New ADODB.Recordset

        rbusca.Open "select * from combina where productop='" & Trim(DBGrid2.columns(1)) & "' and producto='" & xproducto & "'", cn, adOpenStatic, adLockOptimistic
  
        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe producto ", 48, "Aviso"
            Exit Sub

        End If
   
        producto = Trim(DBGrid2.columns(1))
        descripcio = Trim(DBGrid2.columns(0))
        Frame4.Visible = False
        Frame4.Enabled = False
        producto.SetFocus
        producto_KeyPress 13

    End If

End Sub

'Combos Combinaciones 07/07/2018

'Combos Combinaciones 07/07/2018
Private Sub ChaCERRAR_Click()
    Frame4.Visible = False
    producto.SetFocus

End Sub

'Combos Combinaciones 07/07/2018

Private Sub cmdAddEntry_Click()
    ajdu1_Click

End Sub

'Combos Combinaciones 07/07/2018
Private Sub cmdBuscarProducto_Click()
    consulta_producto

End Sub

'Combos Combinaciones 07/07/2018

Private Sub cmdCerrar_Click()
    dlo132_Click

End Sub

Private Sub cmdDelete_Click()
    bo712_Click

End Sub

Private Sub cmdExit_Click()
    dlo132_Click

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdGuardar_Click()

    Dim found As Integer

    found = grabar()

End Sub

Private Sub cmdPrint_Click()

    'djuer1_Click
End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub Combo2_Click()

    If Combo2 <> "%" Then
        gcombina = Trim(Combo2.Text)

    End If

End Sub

Sub ejecutax(sw As Integer)

    Dim buf      As String

    Dim indx     As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd34_err

    If opcion1 = "1" Then
        If Len(bufferx) = 0 Then
            buf = "select Descripcio,Producto,Unidad,Factor,Costou from producto "
        Else
            buf = "select Descripcio,Producto,Unidad,Factor,Costou from producto where "
      
            'Combos Combinaciones 07/07/2018
            buf = buf & "" & Combo3 & " like '%" & bufferx & "%'"

            'Combos Combinaciones 07/07/2018
        End If

    End If

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = mytablex

    If mytablex.RecordCount = 0 Then
        bufferx = ""
        mytablex.Close
        bufferx.SetFocus
        Exit Sub

    End If

    If sw = 1 Then
        DBGrid2.SetFocus

    End If

    Exit Sub
cmd34_err:
    bufferx = ""
    Exit Sub

End Sub

Private Sub Command3_Click()
    ejecutax 1

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
   
            'Combos Combinaciones 07/07/2018
   
            Dim rbusca As New ADODB.Recordset

            rbusca.Open "select * from combina where productop='" & Trim(DBGrid2.columns(1)) & "' and producto='" & xproducto & "'", cn, adOpenStatic, adLockOptimistic
  
            If rbusca.RecordCount > 0 Then
                rbusca.Close
                MsgBox "Ya existe producto ", 48, "Aviso"
                Exit Sub

            End If

            'Combos Combinaciones 07/07/2018
   
            producto = Trim(DBGrid2.columns(1))
            descripcio = Trim(DBGrid2.columns(0))
            Frame4.Visible = False
            Frame4.Enabled = False
            producto.SetFocus
            producto_KeyPress 13

        End If

    End If

End Sub

'Combos Combinaciones 07/07/2018
Private Sub dbgrid2_KeyPress(KeyAscii As Integer)

    Dim buf  As String

    Dim buf2 As String

    If KeyAscii <> 13 And KeyAscii <> 27 Then
        If KeyAscii = 8 Then
            If Len(bufferx) > 0 Then
                buf = Mid$(bufferx, 1, Len(bufferx) - 1)
                bufferx = buf
                KeyAscii = 0
            Else
                KeyAscii = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyAscii)

        If Chr(KeyAscii) = "*" Then
            buf = ""
            bufferx = buf

        End If

        If KeyAscii <> 13 Then
            bufferx = bufferx + buf

        End If

        buf = bufferx
        ejecutax 0
         
    End If

End Sub

'Combos Combinaciones 07/07/2018

Private Sub descripcio_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    'Combos Combinaciones 07/07/2018
    'gcombina.SetFocus
    cantidad.SetFocus
    'Combos Combinaciones 07/07/2018

End Sub

Private Sub dk9893_Click()

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "combina"
    reporgen.Show 1

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\productoesproducto.rpt", "")
End Sub

Private Sub gcombina_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub producto_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(producto) = 0 Then Exit Sub

    'Combos Combinaciones 07/07/2018
    'descripcio.SetFocus
    cantidad.SetFocus

    'Combos Combinaciones 07/07/2018
End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    If opcion1 = "1" Then  'bodega
        If Len(buffer) = 0 Then
            cad = "SELECT * from combina where producto='" & xproducto & "' order by gcombina"

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT *  from combina   where  producto='" & xproducto & "' and " & Combo1 & " like '" & buffer & "%' order by gcombina"

        End If

        If txempre.State = 1 Then txempre.Close
        txempre.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbGrid1.DataSource = txempre
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

        If txempre.RecordCount > 0 Then
            dbGrid1.SetFocus

        End If

    End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'producto = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'producto.SetFocus
        'producto_KeyPress 13
    End If

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

    Dim buf  As String

    Dim buf2 As String

    If KeyAscii <> 13 And KeyAscii <> 27 Then
        If KeyAscii = 8 Then
            If Len(buffer) > 0 Then
                buf = Mid$(buffer, 1, Len(buffer) - 1)
                buffer = buf
                KeyAscii = 0
            Else
                KeyAscii = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyAscii)

        If Chr(KeyAscii) = "*" Then
            buf = ""
            buffer = buf

        End If

        If KeyAscii <> 13 Then
            buffer = buffer + buf

        End If

        buf = buffer
        ejecuta 0
         
    End If

End Sub

Private Sub dlo132_Click()

    If Frame4.Visible = True Then
        Frame4.Visible = False
        Exit Sub

    End If

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    tabcombo.Hide
    Unload tabcombo

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = txempre.Fields("producto")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Modifica"
    cmdGuardar.Enabled = True
    pone_registro
    habilita 1
    producto.Enabled = False

    'Combos Combinaciones 07/07/2018
    cantidad.SetFocus
    'gcombina.SetFocus
    'Combos Combinaciones 07/07/2018

    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = txempre.Fields("producto")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Zoom"
    cmdGuardar.Enabled = False
    pone_registro
    habilita 1
    producto.Enabled = False
    gcombina.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    'agregar_menus
    carga_gcombina
    Command1_Click

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "gcombina"
    Combo1.ListIndex = 0

End Sub

Sub inicializa()
    gcombina = ""
    producto = ""
    descripcio = ""

    'Combos Combinaciones 07/07/2018
    cantidad = ""
    'Combos Combinaciones 07/07/2018

End Sub

Sub pone_registro()
    producto = Trim("" & txempre.Fields("productop"))
    descripcio = Trim("" & txempre.Fields("descripciop"))
    gcombina = Trim("" & txempre.Fields("gcombina"))

    'Combos Combinaciones 07/07/2018
    cantidad = Trim("" & txempre.Fields("cantidad"))
    'Combos Combinaciones 07/07/2018

End Sub

Sub grabando()
    txempre.Fields("producto") = Trim(xproducto)
    txempre.Fields("gcombina") = Trim(gcombina)
    txempre.Fields("productop") = Trim(producto)
    txempre.Fields("descripciop") = Trim(descripcio)

    'Combos Combinaciones 07/07/2018
    txempre.Fields("cantidad") = Trim(cantidad)

    'Combos Combinaciones 07/07/2018
End Sub

Private Sub grba1_Click()

End Sub

Function grabar()

    Dim found  As Integer

    Dim rbusca As New ADODB.Recordset

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos", 48, "Aviso"
        Exit Function

    End If

    If Frame2.Caption = "Nuevo" Then
        If Len(producto) = 0 Then
            producto.SetFocus
            Exit Function

        End If
   
        'Combos Combinaciones 07/07/2018
        'rbusca.Open "select * from combina where producto='" & producto & "' and gcombina='" & gcombina & "'", cn, adOpenStatic, adLockOptimistic
        rbusca.Open "select * from combina where productop='" & producto & "' and producto='" & xproducto & "'", cn, adOpenStatic, adLockOptimistic
        'Combos Combinaciones 07/07/2018
   
        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe producto ", 48, "Aviso"
            Exit Function

        End If

        txempre.AddNew
        txempre.Fields("producto") = producto
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        txempre.Fields("producto") = producto
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    If Len(producto) = 0 Then
        producto.SetFocus
        Exit Function

    End If

    'Combos Combinaciones 07/07/2018
    'If Len(gcombina) = 0 Then
    '   gcombina.SetFocus
    '   Exit Function
    'End If
    If Len(cantidad) = 0 Then
        cantidad.SetFocus
        Exit Function

    End If

    'Combos Combinaciones 07/07/2018

    valida = 1

End Function

Sub habilita(sw As Integer)

    If sw = 0 Then

        ajdu1.Enabled = True
        f8443.Enabled = True
        bo712.Enabled = True
        fjh433.Enabled = True
        djuer1.Enabled = True
        djuer1.Enabled = True
        Picture1.Enabled = True
        dbGrid1.Enabled = True
            
    End If

    If sw = 1 Then

        ajdu1.Enabled = False
        f8443.Enabled = False
        bo712.Enabled = False
        fjh433.Enabled = False
        djuer1.Enabled = False
        djuer1.Enabled = False
        Picture1.Enabled = False
        dbGrid1.Enabled = False
        dbGrid1.Enabled = False
           
    End If
      
End Sub

Sub agregar_menus()

    Dim I As Integer

    For I = 1 To mnuArchivoArray.count - 1
        Unload mnuArchivoArray(I)
    Next
     
    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from archivo where menu='producto' and   estado='S'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        Agregarm "" & mytablex.Fields("descripcio"), mnuArchivoArray
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub Agregarm(TextoDeMenu As String, QueMenu As Object)

    Dim indice As Integer

    'MsgBox QueMenu.count
    indice = QueMenu.count

    Load QueMenu(indice)

    QueMenu(indice).Caption = TextoDeMenu
    QueMenu(indice).Visible = True

End Sub

Sub mnuarchivoarray_click(Index As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    buf = mnuArchivoArray(Index).Caption
    mytablex.Open "select * from archivo where menu='producto' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close

    End If

    'busca el reporte
    buf = mytablex.Fields("archivo")
    mytablex.Close
    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub

Sub carga_gcombina()

    Dim mytablex As New ADODB.Recordset

    Combo2.Clear
    Combo2.AddItem "%"
    mytablex.Open "select * from gcombina", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        Combo2.AddItem "" & mytablex.Fields("gcombina")
        mytablex.MoveNext
    Loop
    mytablex.Close
    Combo2.ListIndex = 0

End Sub

Private Sub producto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_producto

    End If

End Sub

Sub consulta_producto()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "Producto"
    Combo3.ListIndex = 1
    opcion1 = "1"
    Frame4.Visible = True
    Frame4.Enabled = True
    bufferx = ""
    ejecutax 1
    'dbgrid2.SetFocus

    'Combos Combinaciones 07/07/2018
    bufferx.SetFocus
    'Combos Combinaciones 07/07/2018

End Sub
