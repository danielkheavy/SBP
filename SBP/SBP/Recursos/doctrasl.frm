VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form doctrasl 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traslado Locales y Almacenes"
   ClientHeight    =   8595
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   13695
   StartUpPosition =   1  'CenterOwner
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
      Height          =   8535
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   13575
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "doctrasl.frx":0000
         Height          =   7815
         Left            =   120
         OleObjectBlob   =   "doctrasl.frx":0014
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   600
         Width           =   13335
      End
      Begin VB.Label buffer 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   120
         TabIndex        =   49
         Top             =   240
         Width           =   75
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Lista Precios"
      Height          =   3735
      Left            =   2520
      TabIndex        =   44
      Top             =   3720
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton Command8 
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
         Left            =   7440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "doctrasl.frx":09DF
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Borrar registro"
         Top             =   360
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Height          =   3255
         Left            =   120
         OleObjectBlob   =   "doctrasl.frx":1BF1
         TabIndex        =   46
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   13680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Detalle"
      Height          =   5295
      Left            =   120
      TabIndex        =   28
      Top             =   3360
      Width           =   13455
      Begin VB.CommandButton Command2 
         Caption         =   "Ir-Cabecerea"
         Height          =   495
         Left            =   11880
         TabIndex        =   43
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox producto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         MaxLength       =   10
         TabIndex        =   31
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox cantidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8640
         MaxLength       =   10
         TabIndex        =   29
         Top             =   600
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "doctrasl.frx":2C54
         Height          =   3975
         Left            =   120
         OleObjectBlob   =   "doctrasl.frx":2C68
         TabIndex        =   30
         Top             =   960
         Width           =   11655
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor"
         Height          =   255
         Left            =   7560
         TabIndex        =   41
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label descripcio 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   40
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   38
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8640
         TabIndex        =   37
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Paq"
         Height          =   255
         Left            =   6360
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label unidad 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6360
         TabIndex        =   35
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label factor 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7560
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SaldoUltimo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10200
         TabIndex        =   33
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label saldo 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10200
         TabIndex        =   32
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Cabecera"
      Height          =   2415
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   13455
      Begin VB.TextBox fecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   9720
         MaxLength       =   10
         TabIndex        =   50
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ir-Detalle"
         Height          =   495
         Left            =   11880
         TabIndex        =   42
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox serie 
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
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox tipo 
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
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox bodega 
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
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1920
         Width           =   2655
      End
      Begin VB.ComboBox local1 
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
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ComboBox bodega2 
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
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1920
         Width           =   2655
      End
      Begin VB.ComboBox local2 
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
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox vendedor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox numero 
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
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   11
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   51
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alm.Inicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Alm.Dest."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   24
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   23
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lugar Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lugar Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3960
         TabIndex        =   21
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label aksw 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4920
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label nvendedor 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3480
         TabIndex        =   18
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie/Numero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   13635
      TabIndex        =   8
      Top             =   0
      Width           =   13695
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
         Picture         =   "doctrasl.frx":3B3F
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
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
         Picture         =   "doctrasl.frx":4D51
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAddEntry 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
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
         Picture         =   "doctrasl.frx":5F63
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label znumero 
         Height          =   375
         Left            =   11040
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label zserie 
         Height          =   375
         Left            =   10200
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label ztipo 
         Height          =   375
         Left            =   9480
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label bandera 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8760
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Menu dlo23 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "doctrasl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type campo_precio
    unidad As String
    factor As String
    precio As String
    costo As String
    margen As String
    stock As String
End Type

Dim campo_precios(12) As campo_precio


Private Sub bodega_Click()
   saldo_actual

End Sub

Private Sub bodega_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
local2.SetFocus

End Sub

Private Sub bodega2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fecha.SetFocus


End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
cmdSave_Click

End Sub

Private Sub clave_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
producto.SetFocus
End Sub

Private Sub cmdExit_Click()
dlo23_Click
End Sub

Private Sub cmdSave_Click()
Dim found As Integer
If Not IsNumeric(cantidad) Then
   cantidad = ""
   cantidad.SetFocus
   Exit Sub
End If
If Val(cantidad) <= 0 Then
   cantidad = ""
   cantidad.SetFocus
   Exit Sub
End If
If Val(cantidad) * Val(factor) > Val(saldo) Then
   If MsgBox("Cantidad es mayor que el sado," + Chr$(10) + Chr$(13) & "Desea de todas maneras Grabar", 1, "Aviso") <> 1 Then
      cantidad = ""
      cantidad.SetFocus
      Exit Sub
   End If
   'Exit Sub
End If
If MsgBox("Desea Grabar", 1, "Aviso") <> 1 Then Exit Sub
found = grabar()
If found = 0 Then
   MsgBox "No se pudo grabar", 48, "Aviso"
   producto.SetFocus
End If
numero_libre 1
inicializa
sql_detalle
producto.SetFocus

End Sub

Private Sub Command1_Click()
Dim found As Integer
If bodega = bodega2 Then
   MsgBox "Almacenes no deben ser iguales ", 48, "Aviso"
   bodega2.SetFocus
   Exit Sub
End If
found = valida_total()
If found = 0 Then
   bodega2.SetFocus
   Exit Sub
End If
sql_detalle
habilita 1
habilita1 0
producto.SetFocus
End Sub
Function valida_total()
Dim found As Integer
found = busca_tipo()
If found = 0 Then
   MsgBox "No existe Tipo", 48, "Aviso"
   tipo.SetFocus
   Exit Function
End If
found = busca_vendedor()
If found = 0 Then
   MsgBox "No existe Tipo", 48, "Aviso"
   vendedor.SetFocus
   Exit Function
End If
If Len(serie) = 0 Then
   serie.SetFocus
   Exit Function
End If
If Len(numero) = 0 Then
   numero.SetFocus
   Exit Function
End If
If Not IsNumeric(numero) Then
   numero.SetFocus
   Exit Function
End If

If Not IsDate(fecha) Or Len(fecha) <> 10 Then
   fecha.SetFocus
   Exit Function
End If


valida_total = 1

End Function
Sub habilita(sw As Integer)
Dim xsw
xsw = False
If sw = 0 Then
   xsw = True
End If
cmdSave.Enabled = xsw
Command1.Enabled = xsw
tipo.Enabled = xsw
serie.Enabled = xsw
numero.Enabled = xsw
vendedor.Enabled = xsw
local1.Enabled = xsw
local2.Enabled = xsw
bodega.Enabled = xsw
bodega2.Enabled = xsw

End Sub
Sub habilita1(sw As Integer)
Dim xsw
xsw = False
If sw = 0 Then
   xsw = True
End If
cmdSave.Enabled = xsw
Command2.Enabled = xsw
producto.Enabled = xsw
cantidad.Enabled = xsw
End Sub

Private Sub Command2_Click()
habilita 0
habilita1 1
bodega2.SetFocus
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode <> 13 And KeyCode <> 27 Then Exit Sub
If KeyCode = 27 Then
  dlo23_Click
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = "1" Then
      producto = DBGrid1.columns(1)
      Frame1.Visible = False
      producto.SetFocus
      producto_KeyPress 13
   End If
   If opcion1 = "2" Then
      tipo = DBGrid1.columns(1)
      Frame1.Visible = False
      tipo.SetFocus
      tipo_KeyPress 13
   End If
   If opcion1 = "3" Then
      vendedor = DBGrid1.columns(1)
      Frame1.Visible = False
      vendedor.SetFocus
      vendedor_KeyPress 13
   End If

End If

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
Dim buf As String
Dim buf2 As String
Dim found As Integer
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
         If Chr(KeyAscii) = "/" Then
            buf = ""
            buffer = buf
         End If
         If KeyAscii <> 13 Then
            buffer = buffer + buf
         End If
         KeyAscii = 0
         buf = buffer
         'MsgBox buf & " " & buffer
         found = ejecuta(0)
         If found = 0 Then
            found = ejecuta(0)
         End If
End If
End Sub

Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim buf As String
Dim found As Integer
If KeyCode <> 13 And KeyCode <> 27 Then
          'MsgBox Shift
          If KeyCode = 32 Then
             GoTo sigue9
          End If
          If KeyCode >= 48 And KeyCode <= 57 Then
             GoTo sigue9
          End If
          If KeyCode >= 65 And KeyCode <= 90 Then
             GoTo sigue9
          End If
          If KeyCode >= 97 And KeyCode <= 122 Then
             GoTo sigue9
          End If
          If KeyCode = 8 Or Chr(KeyCode) = "*" Then
             GoTo sigue9
          End If
          Exit Sub
sigue9:
          If KeyCode = 8 Then
            If Len(buffer) > 0 Then
               buf = Mid$(buffer, 1, Len(buffer) - 1)
               buffer = buf
               KeyCode = 0
               Else
               KeyCode = 0
               Exit Sub
            End If
         End If
         MsgBox KeyCode
         buf = Chr(KeyCode)
         If Chr(KeyCode) = "/" Then
            buf = ""
            buffer = buf
         End If
         If KeyCode <> 13 Then
            buffer = buffer + buf
         End If
         'MsgBox buffer
         buf = buffer
         found = ejecuta(0)
         If found = 0 Then
            found = ejecuta(1)
         End If
Exit Sub
End If

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim found As Integer
On Error GoTo cmd12_err
If KeyCode = &H2E Then  'borrar linea
   found = desgrabar_saldo()
   If found = 1 Then
      Data2.Recordset.Delete
   End If
End If
producto.SetFocus
Exit Sub
producto.SetFocus
cmd12_err:
Exit Sub

End Sub

Private Sub dlo23_Click()
If Frame1.Visible = True Then
If opcion1 = "1" Then
   Frame1.Visible = False
   producto.SetFocus
   Exit Sub
End If
If opcion1 = "2" Then
   Frame1.Visible = False
   tipo.SetFocus
   Exit Sub
End If
If opcion1 = "3" Then
   Frame1.Visible = False
   vendedor.SetFocus
   Exit Sub
End If


End If
doctrasl.Hide
Unload doctrasl
End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(fecha) = 0 Then
   fecha = Format(Now, "dd/mm/yyyy")
End If
Command1.SetFocus
End Sub

Private Sub Form_Activate()
If aksw = "" Then
   carga_inicial
   saldo_actual
End If
aksw = "1"

End Sub
Sub carga_dbgrid4(xproducto As String)
Dim i As Integer

Dim mytablex As Table
Dim mytabley As Table
Dim sw As Integer
Dim xbodega As String
Dim xsaldo As Double
Dim xbuf As String
Dim xcosto As Double
Dim xcostou As Double
Dim xfactor As Double
Dim xunidad As String
Dim xmargen As Double
On Error GoTo cmd89012_err
For i = 0 To 9
    campo_precios(i).unidad = ""
    campo_precios(i).factor = ""
    campo_precios(i).precio = ""
    campo_precios(i).costo = ""
    campo_precios(i).margen = ""
    campo_precios(i).stock = ""
Next i
'MsgBox "hoal"
xcostou = 0
xunidad = "UND"
xfactor = 1
xbodega = extra_loquesea(bodega)
xsaldo = 0
xcosto = 0
sw = 0
Set mytabley = mydbxglo.OpenTable("almacen")
mytabley.Index = "almacen"
mytabley.Seek "=", local1, xproducto, xbodega
If Not mytabley.NoMatch Then
   xsaldo = Val("" & mytabley.Fields("saldo"))
End If
mytabley.Close
'MsgBox "h"

Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "producto"
mytablex.Seek "=", xproducto
If Not mytablex.NoMatch Then
   xcostou = Val("" & mytablex.Fields("costou"))
   xfactor = Val("" & mytablex.Fields("factor"))
   xunidad = "" & mytablex.Fields("unidad")
End If
mytablex.Close
Set mytablex = mydbxglo.OpenTable("precios")
mytablex.Index = "tprecios"
   '----------------------------------------------
   mytablex.Seek "=", xproducto, "01"
   If Not mytablex.NoMatch Then
     'MsgBox "Hola"
      xcosto = xcostou
      campo_precios(0).unidad = xunidad
      campo_precios(0).factor = xfactor
      campo_precios(0).precio = "" '& mytablex.Fields("costou")
      campo_precios(0).costo = xcostou
      xbuf = calcula_saldo(xsaldo, xfactor)
      campo_precios(0).stock = "" & xbuf
      xmargen = 0
      campo_precios(0).margen = "" & xmargen
      '----------------------------------------------
      xcosto = 0
      If Val("" & mytablex.Fields("factor1")) > 0 Then
         xcosto = xcostou / xfactor
         xcosto = xcosto * Val("" & mytablex.Fields("factor1"))
      campo_precios(1).unidad = "" & mytablex.Fields("unidad1")
      campo_precios(1).factor = "" & mytablex.Fields("factor1")
      campo_precios(1).precio = "" & mytablex.Fields("pventa1")
      campo_precios(1).costo = "" & xcosto
      xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor1")))
      campo_precios(1).stock = "" & xbuf
      xmargen = 0
      If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa1")) - xcosto) * 100) / xcosto
      End If
      campo_precios(1).margen = "" & xmargen
   '--------
   End If
   '---------
   If Val("" & mytablex.Fields("factor2")) > 0 Then
   campo_precios(2).unidad = "" & mytablex.Fields("unidad2")
   campo_precios(2).factor = "" & mytablex.Fields("factor2")
   campo_precios(2).precio = "" & mytablex.Fields("pventa2")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor2")))
   campo_precios(2).stock = "" & xbuf
   xcosto = 0
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor2"))
   campo_precios(2).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((Val("" & mytablex.Fields("pventa2")) - xcosto) * 100) / xcosto
   End If
   campo_precios(2).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor3")) > 0 Then
   campo_precios(3).unidad = "" & mytablex.Fields("unidad3")
   campo_precios(3).factor = "" & mytablex.Fields("factor3")
   campo_precios(3).precio = "" & mytablex.Fields("pventa3")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor3")))
   campo_precios(3).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor3"))
   
   campo_precios(3).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa3")) - xcosto) * 100) / xcosto
         campo_precios(3).margen = "" & xmargen
   End If
   campo_precios(3).margen = "" & xmargen
   End If
   If Val("" & mytablex.Fields("factor4")) > 0 Then
   campo_precios(4).unidad = "" & mytablex.Fields("unidad4")
   campo_precios(4).factor = "" & mytablex.Fields("factor4")
   campo_precios(4).precio = "" & mytablex.Fields("pventa4")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor4")))
   campo_precios(4).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor4"))
   
   campo_precios(4).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa4")) - xcosto) * 100) / xcosto
   End If
   campo_precios(4).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor5")) > 0 Then
   campo_precios(5).unidad = "" & mytablex.Fields("unidad5")
   campo_precios(5).factor = "" & mytablex.Fields("factor5")
   campo_precios(5).precio = "" & mytablex.Fields("pventa5")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor5")))
   campo_precios(5).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor5"))
   
   campo_precios(5).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto
   End If
   campo_precios(5).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor6")) > 0 Then
   campo_precios(6).unidad = "" & mytablex.Fields("unidad6")
   campo_precios(6).factor = "" & mytablex.Fields("factor6")
   campo_precios(6).precio = "" & mytablex.Fields("pventa6")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor6")))
   campo_precios(6).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor6"))
   
   campo_precios(6).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto
         
   End If
   campo_precios(6).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor7")) > 0 Then
   campo_precios(7).unidad = "" & mytablex.Fields("unidad7")
   campo_precios(7).factor = "" & mytablex.Fields("factor7")
   campo_precios(7).precio = "" & mytablex.Fields("pventa7")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor7")))
   campo_precios(7).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor7"))
   
   campo_precios(7).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa7")) - xcosto) * 100) / xcosto
   End If
   campo_precios(7).margen = "" & xmargen
   End If
   
   
   If Val("" & mytablex.Fields("factor8")) > 0 Then
   campo_precios(8).unidad = "" & mytablex.Fields("unidad8")
   campo_precios(8).factor = "" & mytablex.Fields("factor8")
   campo_precios(8).precio = "" & mytablex.Fields("pventa8")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor8")))
   campo_precios(8).stock = "" & xbuf
   xcosto = 0
   
      xcosto = xcostou / xfactor
      xcosto = xcosto * Val("" & mytablex.Fields("factor8"))
   
   campo_precios(8).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((Val("" & mytablex.Fields("pventa8")) - xcosto) * 100) / xcosto
   End If
   campo_precios(8).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor9")) > 0 Then
   campo_precios(9).unidad = "" & mytablex.Fields("unidad9")
   campo_precios(9).factor = "" & mytablex.Fields("factor9")
   campo_precios(9).precio = "" & mytablex.Fields("pventa9")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor9")))
   campo_precios(9).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor9"))
   
   campo_precios(9).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa9")) - xcosto) * 100) / xcosto
         
   End If
   campo_precios(9).margen = "" & xmargen
   End If
   
   If Val("" & mytablex.Fields("factor10")) > 0 Then
   campo_precios(10).unidad = "" & mytablex.Fields("unidad10")
   campo_precios(10).factor = "" & mytablex.Fields("factor10")
   campo_precios(10).precio = "" & mytablex.Fields("pventa10")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor10")))
   campo_precios(10).stock = "" & xbuf
   xcosto = 0
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor10"))
   
   campo_precios(10).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa10")) - xcosto) * 100) / xcosto
   End If
   campo_precios(10).margen = "" & xmargen
   End If
   'margenes
   sw = 1
   
 End If
mytablex.Close
dbgrid4.Refresh
Frame5.Visible = True
dbgrid4.SetFocus
Exit Sub
cmd89012_err:
MsgBox "Error en carga Grid " + error$, 48, "Aviso"
Exit Sub

End Sub

Sub inicializa()
producto = ""
cantidad = ""
descripcio = ""
unidad = ""
factor = ""
End Sub
Function grabar()
Dim found As Integer
Dim mytablex As Table

Dim sw As Integer
sw = 0
Set mytablex = mydbxglo.OpenTable("almacen")
mytablex.Index = "almacen"
mytablex.Seek "=", extra_loquesea(local1), producto, extra_loquesea(bodega)
If Not mytablex.NoMatch Then
   mytablex.Edit
   mytablex.Fields("saldo") = Val("" & mytablex.Fields("saldo")) - Val("" & cantidad) * Val(factor)
   mytablex.Update
   sw = 1
End If
If mytablex.NoMatch Then
   mytablex.AddNew
   mytablex.Fields("producto") = "" & producto
   mytablex.Fields("local") = extra_loquesea(local1)
   mytablex.Fields("bodega") = extra_loquesea(bodega)
   mytablex.Fields("saldo") = -Val("" & cantidad) * Val(factor)
   mytablex.Update
   sw = 1
End If
mytablex.Seek "=", extra_loquesea(local2), producto, extra_loquesea(bodega2)
If Not mytablex.NoMatch Then
   mytablex.Edit
   mytablex.Fields("saldo") = Val("" & mytablex.Fields("saldo")) + Val("" & cantidad) * Val(factor)
   mytablex.Update
   sw = 1
End If

If mytablex.NoMatch Then
   mytablex.AddNew
   mytablex.Fields("producto") = "" & producto
   mytablex.Fields("local") = extra_loquesea(local2)
   mytablex.Fields("bodega") = extra_loquesea(bodega2)
   mytablex.Fields("saldo") = Val("" & cantidad) * Val(factor)
   mytablex.Update
   sw = 1
End If
mytablex.Close
found = graba_traslado()
found = graba_kardex()
If found = 1 Then
   sw = 1
End If
grabar = 1
End Function
Function graba_traslado()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("ctraslad")
mytablex.Index = "tfactura"
mytablex.Seek "=", extra_loquesea(local1), "Z", "Z01", numero
If mytablex.NoMatch Then
mytablex.AddNew
mytablex.Fields("adetotal") = 0
mytablex.Fields("acuenta") = 0
mytablex.Fields("retipo1") = ""
mytablex.Fields("renumero1") = ""
mytablex.Fields("renumero2") = ""
mytablex.Fields("renumero3") = ""
mytablex.Fields("retotal1") = 0
mytablex.Fields("retotal2") = 0
mytablex.Fields("retotal3") = 0
mytablex.Fields("retotal") = 0
mytablex.Fields("tflete") = 0
mytablex.Fields("zona") = ""
mytablex.Fields("nombre") = ""
mytablex.Fields("estado") = "2"
mytablex.Fields("tipoclie") = "V"
mytablex.Fields("tipo") = "Z"
mytablex.Fields("serie") = "Z01"
mytablex.Fields("numero") = numero
mytablex.Fields("codigo") = extra_loquesea(local1)
mytablex.Fields("partida") = ""
mytablex.Fields("destino") = ""
mytablex.Fields("nro_items") = 0
   mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
   mytablex.Fields("fechasunat") = Format(fecha, "dd/mm/yyyy")
   mytablex.Fields("fechae") = Format(fecha, "dd/mm/yyyy")
mytablex.Fields("moneda") = "S"
mytablex.Fields("vendedor") = vendedor
mytablex.Fields("fpago") = ""
mytablex.Fields("transporte") = ""
mytablex.Fields("paridad") = 1
mytablex.Fields("dias") = 1
mytablex.Fields("bodega") = extra_loquesea(bodega)
mytablex.Fields("localf") = extra_loquesea(local2)
mytablex.Fields("bodegaf") = extra_loquesea(bodega2)
mytablex.Fields("observa") = "FORMA RAPIDA"
mytablex.Fields("usuario") = "" & gusuario
mytablex.Fields("acu") = "Z"
mytablex.Fields("acu1") = ""
mytablex.Fields("flage") = ""
mytablex.Fields("hora") = Format(Now, "hh:MM")
mytablex.Fields("fechacrea") = Format(fecha, "dd/mm/yyyy")
mytablex.Fields("fechasunat") = Format(fecha, "dd/mm/yyyy")
mytablex.Fields("total") = 0
mytablex.Fields("descuento") = 0
mytablex.Fields("neto") = 0
mytablex.Fields("gravado") = 0
mytablex.Fields("impuesto") = 19
mytablex.Fields("subtotal") = 0
mytablex.Fields("percepcion") = 0

mytablex.Fields("tipo1") = ""
mytablex.Fields("serie1") = ""
mytablex.Fields("serie2") = ""
mytablex.Fields("serie3") = ""
mytablex.Fields("serie4") = ""
mytablex.Fields("serie5") = ""
mytablex.Fields("serie6") = ""
mytablex.Fields("serie7") = ""

mytablex.Fields("numero1") = ""
mytablex.Fields("numero2") = ""
mytablex.Fields("numero3") = ""
mytablex.Fields("numero4") = ""
mytablex.Fields("numero5") = ""
mytablex.Fields("numero6") = ""
mytablex.Fields("numero7") = ""
mytablex.Fields("local") = extra_loquesea(local1)
mytablex.Fields("localf") = extra_loquesea(local2)

mytablex.Fields("c1") = 0
mytablex.Fields("c2") = 0
mytablex.Fields("c3") = 0
mytablex.Fields("c4") = 0
mytablex.Update
End If
mytablex.Close
Set mytablex = mydbxglo.OpenTable("Dtraslad")
mytablex.Index = "tdetalle"
mytablex.AddNew
mytablex.Fields("local") = extra_loquesea(local1)
mytablex.Fields("estado") = "2"
mytablex.Fields("acu") = "Z"
mytablex.Fields("tipo") = "Z"
mytablex.Fields("serie") = "Z01"
mytablex.Fields("numero") = "" & numero
mytablex.Fields("tipoclie") = "V"
mytablex.Fields("codigo") = extra_loquesea(local2)
mytablex.Fields("vendedor") = "" & vendedor
mytablex.Fields("acu1") = ""
mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
mytablex.Fields("moneda") = "S"
mytablex.Fields("producto") = "" & producto
mytablex.Fields("descripcio") = "" & descripcio
mytablex.Fields("unidad") = unidad
mytablex.Fields("factor") = Val(factor)
mytablex.Fields("cantidad") = Val(cantidad)
mytablex.Fields("precio") = 0
mytablex.Fields("igv") = 19
mytablex.Fields("neto") = 0
mytablex.Fields("descuento") = 0
mytablex.Fields("subtotal") = 0
mytablex.Fields("impuesto") = 0
mytablex.Fields("total") = 0
mytablex.Fields("fechacrea") = Format(fecha, "dd/mm/yyyy")
mytablex.Fields("hora") = Format(Now, "hh:mm:ss")
mytablex.Fields("bodega") = extra_loquesea(bodega)
mytablex.Fields("bodegaf") = extra_loquesea(bodega2)
mytablex.Fields("deslipo") = 0
mytablex.Fields("flage") = ""
mytablex.Fields("linea") = ""
mytablex.Fields("t1") = 0
mytablex.Fields("t2") = 0
mytablex.Fields("t3") = 0
mytablex.Fields("t4") = 0
mytablex.Fields("t5") = 0
mytablex.Fields("t6") = 0
mytablex.Fields("t7") = 0
mytablex.Fields("t8") = 0
mytablex.Fields("t9") = 0
mytablex.Fields("t10") = 0
mytablex.Fields("t11") = 0
mytablex.Fields("t12") = 0
mytablex.Fields("t13") = 0
mytablex.Fields("t14") = 0
mytablex.Fields("t15") = 0
mytablex.Fields("t16") = 0
mytablex.Fields("l1") = ""
mytablex.Fields("l2") = ""
mytablex.Fields("l3") = ""
mytablex.Fields("l4") = ""
'mytablex.Fields("local") = ""
mytablex.Fields("proveedorp") = ""
mytablex.Fields("observa1") = ""
mytablex.Fields("observa2") = ""
mytablex.Fields("observa3") = ""
mytablex.Fields("observa4") = ""
mytablex.Fields("zona") = ""
mytablex.Fields("isc") = 0
mytablex.Fields("tax") = 0
mytablex.Fields("vtaneta") = 0
mytablex.Fields("tcosto") = 0
mytablex.Fields("ganancia") = 0
mytablex.Fields("comision") = 0
mytablex.Fields("usuario") = gusuario
mytablex.Fields("caja") = ""
mytablex.Fields("turno") = ""
mytablex.Fields("servicio") = ""
mytablex.Fields("comanda") = ""
mytablex.Fields("mesa") = ""
mytablex.Fields("salon") = ""
mytablex.Fields("mesero") = ""
mytablex.Update
End Function
Function graba_kardex()
On Error GoTo cmd781_err
Dim mytablez As Table
Set mytablez = mydbxglo.OpenTable("detalle")
mytablez.AddNew
'salida
mytablez.Fields("local") = extra_loquesea(local1)
mytablez.Fields("estado") = "2"
mytablez.Fields("acu") = "T"
mytablez.Fields("tipo") = "TS"
mytablez.Fields("serie") = ""
mytablez.Fields("numero") = numero
mytablez.Fields("tipoclie") = "V"
mytablez.Fields("codigo") = "" & local1
mytablez.Fields("vendedor") = "" & vendedor
mytablez.Fields("acu1") = ""
mytablez.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
mytablez.Fields("moneda") = "S"
mytablez.Fields("producto") = "" & producto
mytablez.Fields("descripcio") = "" & descripcio
mytablez.Fields("unidad") = unidad
mytablez.Fields("factor") = Val(factor)
mytablez.Fields("cantidad") = Val(cantidad)
mytablez.Fields("precio") = 0
mytablez.Fields("igv") = 19
mytablez.Fields("neto") = 0
mytablez.Fields("descuento") = 0
mytablez.Fields("subtotal") = 0
mytablez.Fields("impuesto") = 0
mytablez.Fields("total") = 0
mytablez.Fields("fechacrea") = Format(fecha, "dd/mm/yyyy")
mytablez.Fields("hora") = Format(Now, "hh:mm:ss")
mytablez.Fields("bodega") = extra_loquesea(bodega)
mytablez.Fields("bodegaf") = "" 'extra_loquesea(bodegaf)
mytablez.Fields("deslipo") = 0
mytablez.Fields("flage") = ""
mytablez.Fields("linea") = ""
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
mytablez.Fields("local") = ""
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
mytablez.Fields("usuario") = gusuario
mytablez.Fields("caja") = ""
mytablez.Fields("turno") = ""
mytablez.Fields("servicio") = ""
mytablez.Fields("comanda") = ""
mytablez.Fields("mesa") = ""
mytablez.Fields("salon") = ""
mytablez.Fields("mesero") = ""
mytablez.Update
'ENTRADA
mytablez.AddNew
mytablez.Fields("local") = extra_loquesea(local2)
mytablez.Fields("localf") = "" 'extra_loquesea(local2)"
mytablez.Fields("estado") = "2"
mytablez.Fields("acu") = "S"
mytablez.Fields("tipo") = "TE"
mytablez.Fields("serie") = ""
mytablez.Fields("numero") = "" & numero
mytablez.Fields("tipoclie") = "V"
mytablez.Fields("codigo") = "" & local2
mytablez.Fields("vendedor") = "" & vendedor

mytablez.Fields("acu1") = ""
mytablez.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
mytablez.Fields("moneda") = "S"
mytablez.Fields("producto") = "" & producto
mytablez.Fields("descripcio") = "" & descripcio
mytablez.Fields("unidad") = unidad
mytablez.Fields("factor") = Val(factor)
mytablez.Fields("cantidad") = Val(cantidad)
mytablez.Fields("precio") = 0
mytablez.Fields("igv") = 19
mytablez.Fields("neto") = 0
mytablez.Fields("descuento") = 0
mytablez.Fields("subtotal") = 0
mytablez.Fields("impuesto") = 0
mytablez.Fields("total") = 0
mytablez.Fields("fechacrea") = Format(fecha, "dd/mm/yyyy")
mytablez.Fields("hora") = Format(Now, "hh:mm:ss")
mytablez.Fields("bodega") = extra_loquesea(bodega2)
mytablez.Fields("bodegaf") = ""
mytablez.Fields("deslipo") = 0
mytablez.Fields("flage") = ""
mytablez.Fields("linea") = ""
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
mytablez.Fields("local") = ""
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
mytablez.Fields("usuario") = gusuario
mytablez.Fields("caja") = ""
mytablez.Fields("turno") = ""
mytablez.Fields("servicio") = ""
mytablez.Fields("comanda") = ""
mytablez.Fields("mesa") = ""
mytablez.Fields("salon") = ""
mytablez.Fields("mesero") = ""
mytablez.Update
mytablez.Close
graba_kardex = 1
Exit Function
cmd781_err:
MsgBox "Error " + error$, 48, "Aviso"
Exit Function
End Function
Sub carga_inicial()
Dim mytablex As Table
local1.Clear
local2.Clear
Set mytablex = mydbxglo.OpenTable("tlocal")
Do
If mytablex.EOF Then Exit Do
local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
local2.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")

mytablex.MoveNext
Loop
mytablex.Close
local1.ListIndex = 0
local2.ListIndex = 0
bodega.Clear
bodega.Clear
Set mytablex = mydbxglo.OpenTable("bodega")
Do
If mytablex.EOF Then Exit Do
bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
bodega2.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")

mytablex.MoveNext
Loop
mytablex.Close
bodega.ListIndex = 0
bodega2.ListIndex = 0
End Sub
Sub saldo_actual()
Dim mytablex As Table
saldo = ""
Set mytablex = mydbxglo.OpenTable("almacen")
mytablex.Index = "almacen"
mytablex.Seek "=", extra_loquesea(local1), producto, extra_loquesea(bodega)
If Not mytablex.NoMatch Then
   saldo = "" & mytablex.Fields("saldo")
End If
mytablex.Close
End Sub

Private Sub Form_Load()
habilita 0
habilita1 1
fecha = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Label13_Click()
If Len(producto) = 0 Or Len(descripcio) = 0 Then
   producto.SetFocus
   Exit Sub
End If

carga_dbgrid4 producto
End Sub

Private Sub local1_Click()
   saldo_actual

End Sub


Private Sub local1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
bodega.SetFocus
End Sub

Private Sub local2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
bodega2.SetFocus

End Sub

Private Sub numero_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(numero) = 0 Then
    numero_libre 0
End If
vendedor.SetFocus
End Sub
Sub numero_libre(sw As Integer)
Dim mytablex As Table
Dim sdx As Double
sdx = 0
Set mytablex = mydbxglo.OpenTable("tipo")
mytablex.Index = "tipo"
mytablex.Seek "=", "Z"
If Not mytablex.NoMatch Then
   If sw = 1 Then
      mytablex.Edit
      mytablex.Fields("numero") = numero
      mytablex.Update
   End If
   If sw = 0 Then
      sdx = Val("" & mytablex.Fields("numero")) + 1
      numero = "" & sdx
   End If
End If
mytablex.Close


End Sub
Private Sub producto_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(producto) = 0 Then Exit Sub
found = busca_producto()
If found = 0 Then
   MsgBox "No existe producto", 48, "Aviso"
   Exit Sub
End If
bodega_Click
cantidad.SetFocus

End Sub
Function busca_producto()
Dim mytablex As Table
Dim mytabley As Table
descripcio = ""
Set mytabley = mydbxglo.OpenTable("precios")
mytabley.Index = "tprecios"
Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "producto"
mytablex.Seek "=", producto
If Not mytablex.NoMatch Then
   descripcio = "" & mytablex.Fields("descripcio")
   mytabley.Seek "=", producto, "01"
   If Not mytabley.NoMatch Then
      unidad = "" & mytabley.Fields("unidad1")
      factor = "" & mytabley.Fields("factor1")
      busca_producto = 1
   End If
End If
mytablex.Close
mytabley.Close

End Function
Function busca_vendedor()
Dim mytablex As Table
Dim sdx As Double
nvendedor = ""
Set mytablex = mydbxglo.OpenTable("vendedor")
mytablex.Index = "codigo"
mytablex.Seek "=", vendedor
If Not mytablex.NoMatch Then
   nvendedor = "" & mytablex.Fields("nombre")
   busca_vendedor = 1
End If
mytablex.Close
End Function
Function ejecuta(sw As Integer)
Dim buf As String
Dim indx As Integer
On Error GoTo cmd7654_err
indx = -1
If opcion1 = "1" Then
     If Len(buffer) = 0 Then
      buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,Precios.Unidad1 as Und1,Precios.Factor1 as F,Precios.pventa1 ,Producto.Monedav as M,Producto.Familia,Producto.Subfamilia,Producto.Proveedor1 from producto left join precios on producto.producto=precios.producto  where precios.local='" & extra_loquesea(local1) & "'"
      Else
      buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,Precios.Unidad1 as Und1,Precios.Factor1 as F,Precios.Pventa1 ,Producto.Monedav as M,Producto.Familia,Producto.Subfamilia,Producto.Proveedor1 from producto  left join precios on producto.producto=precios.producto  where precios.local='" & extra_loquesea(local1) & "' and  "
      buf = buf & "" & Data1.Recordset.Fields("" & DBGrid1.columns(DBGrid1.Col).Caption).Name & " like '" & buffer & "%'"
      indx = DBGrid1.Col
      End If
End If
If opcion1 = "2" Then
     If Len(buffer) = 0 Then
      buf = "select Descripcio,Tipo from tipo where tipodoc='Z'"
      Else
      buf = "select Descripcio,Tipo from tipo  WHERE tipodoc='Z' and   "
      buf = buf & "" & Data1.Recordset.Fields("" & DBGrid1.columns(DBGrid1.Col).Caption).Name & " like '" & buffer & "%'"
      indx = DBGrid1.Col
      End If
End If
If opcion1 = "3" Then
     If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from Vendedor"
      Else
      buf = "select Nombre,Codigo from Vendedor  WHERE   "
      buf = buf & "" & Data1.Recordset.Fields("" & DBGrid1.columns(DBGrid1.Col).Caption).Name & " like '" & buffer & "%'"
      indx = DBGrid1.Col
      End If
End If


               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = buf
               Data1.Refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  'buffer.SetFocus
                  buffer = ""
                  Exit Function
               End If
               If opcion1 = "2" Then
                  DBGrid1.columns(0).Width = 5000
                  DBGrid1.columns(1).Width = 1300
               End If
               
               If opcion1 = "1" Then
               DBGrid1.columns(0).Width = 5000
               DBGrid1.columns(1).Width = 1300
               DBGrid1.columns(2).Width = 1000
               DBGrid1.columns(3).Width = 900
               DBGrid1.columns(4).Width = 500
               DBGrid1.columns(5).Width = 800
               DBGrid1.columns(6).Width = 500
               DBGrid1.columns(7).Width = 1000
               DBGrid1.columns(8).Width = 1500
               DBGrid1.columns(9).Width = 1500
               End If
               If sw = 1 Then
                  DBGrid1.SetFocus
               End If
               If indx <> -1 Then
                  DBGrid1.Col = indx
               End If
ejecuta = 1
Exit Function
cmd7654_err:
buffer = ""
MsgBox error$
Exit Function
End Function

Private Sub producto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then
Frame1.Visible = True
buffer = ""
opcion1 = "1"
ejecuta 1
End If

End Sub

Private Sub serie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
numero.SetFocus
End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = busca_tipo()
If found = 0 Then
   MsgBox "No existe Tipo", 48, "Aviso"
   tipo.SetFocus
   Exit Sub
End If
serie.SetFocus
End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then
Frame1.Visible = True
buffer = ""
opcion1 = "2"
ejecuta 1
End If

End Sub

Private Sub vendedor_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = busca_vendedor()
If found = 0 Then
   MsgBox "No existe Tipo", 48, "Aviso"
   vendedor.SetFocus
   Exit Sub
End If
local1.SetFocus
End Sub

Private Sub vendedor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then
Frame1.Visible = True
buffer = ""
opcion1 = "3"
ejecuta 1
End If

End Sub
Function busca_tipo()
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("Tipo")
mytablex.Index = "tipo"
mytablex.Seek "=", tipo
If Not mytablex.NoMatch Then
   If "" & mytablex.Fields("tipodoc") = "Z" Then
      serie = "" & mytablex.Fields("serie")
      busca_tipo = 1
   End If
End If
mytablex.Close

End Function
Sub sql_detalle()
Dim buf As String
buf = "select * from dtraslad where local='" & extra_loquesea(local1) & "'"
buf = buf & " and tipo='" & tipo & "'"
buf = buf & " and serie='" & serie & "'"
buf = buf & " and numero='" & numero & "'"
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
End Sub
Private Sub DBGrid4_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sdx As Double
If KeyCode = 27 Then
   Frame5.Visible = False
   producto.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   unidad = "" & dbgrid4.columns(0)
   factor = "" & dbgrid4.columns(1)
   Frame5.Visible = False
   cantidad.SetFocus
End If

End Sub

Private Sub DBGrid4_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim dr As Integer
Dim row_num As Integer
Dim r As Integer
Dim rows_returned As Integer
If ReadPriorRows Then
        dr = -1
    Else
        dr = 1
    End If
    If IsNull(StartLocation) Then
        If ReadPriorRows Then
           row_num = RowBuf.RowCount - 1
           'row_num = 9
        Else
           row_num = 0
        End If
    Else
        row_num = CLng(StartLocation) + dr
    End If
    rows_returned = 0
    For r = 0 To RowBuf.RowCount - 1
        If row_num < 0 Or row_num > 9 Then Exit For
        RowBuf.Value(r, 0) = campo_precios(row_num).unidad
        RowBuf.Value(r, 1) = campo_precios(row_num).factor
        RowBuf.Value(r, 2) = campo_precios(row_num).precio
        RowBuf.Value(r, 3) = campo_precios(row_num).costo
        RowBuf.Value(r, 4) = campo_precios(row_num).margen
        RowBuf.Value(r, 5) = campo_precios(row_num).stock
        RowBuf.Bookmark(r) = row_num
        row_num = row_num + dr
        rows_returned = rows_returned + 1
   Next r
   RowBuf.RowCount = rows_returned
End Sub
Function desgrabar_saldo()
Dim mytablex As Table
Dim sw As Integer
sw = 0
Set mytablex = mydbxglo.OpenTable("almacen")
mytablex.Index = "almacen"
mytablex.Seek "=", extra_loquesea(local1), "" & Data2.Recordset.Fields("producto"), extra_loquesea(bodega)
If Not mytablex.NoMatch Then
   mytablex.Edit
   mytablex.Fields("saldo") = Val("" & mytablex.Fields("saldo")) + Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("factor"))
   mytablex.Update
   sw = 1
End If
If mytablex.NoMatch Then
   mytablex.AddNew
   mytablex.Fields("producto") = "" & producto
   mytablex.Fields("local") = extra_loquesea(local1)
   mytablex.Fields("bodega") = extra_loquesea(bodega)
   mytablex.Fields("saldo") = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("factor"))
   mytablex.Update
   sw = 1
End If
mytablex.Seek "=", extra_loquesea(local2), "" & Data2.Recordset.Fields("producto"), extra_loquesea(bodega2)
If Not mytablex.NoMatch Then
   mytablex.Edit
   mytablex.Fields("saldo") = Val("" & mytablex.Fields("saldo")) - Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("factor"))
   mytablex.Update
   sw = 1
End If

If mytablex.NoMatch Then
   mytablex.AddNew
   mytablex.Fields("producto") = "" & producto
   mytablex.Fields("local") = extra_loquesea(local2)
   mytablex.Fields("bodega") = extra_loquesea(bodega2)
   mytablex.Fields("saldo") = -Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("factor"))
   mytablex.Update
   sw = 1
End If
mytablex.Close
If sw = 1 Then
   desgrabar_saldo = 1
End If
End Function

