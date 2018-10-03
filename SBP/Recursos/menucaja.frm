VERSION 5.00
Object = "{19BD1EA6-6E36-45BA-AEBD-BCF3093017CC}#11.0#0"; "GorditoButton.ocx"
Begin VB.Form menucaja 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Ingreso CAJA REGISTRADORA Y/O TERMINAL PEDIDOS"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Cambio de Paridad"
      ForeColor       =   &H00404040&
      Height          =   5610
      Left            =   1320
      TabIndex        =   32
      Top             =   5955
      Visible         =   0   'False
      Width           =   12930
      Begin VB.CommandButton Command11 
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
         Left            =   4830
         MaskColor       =   &H00E0E0E0&
         Picture         =   "menucaja.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Grabar registro"
         Top             =   1590
         Width           =   975
      End
      Begin VB.CommandButton Command10 
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
         Height          =   735
         Left            =   4800
         MaskColor       =   &H00E0E0E0&
         Picture         =   "menucaja.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Salir"
         Top             =   795
         Width           =   975
      End
      Begin VB.TextBox venta 
         Height          =   615
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   34
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox compra 
         Height          =   615
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   33
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VENTA"
         Height          =   615
         Left            =   360
         TabIndex        =   38
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COMPRA"
         Height          =   615
         Left            =   360
         TabIndex        =   37
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Mensajes del Sistema...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   6615
      Left            =   615
      TabIndex        =   29
      Top             =   5490
      Visible         =   0   'False
      Width           =   12855
      Begin VB.CommandButton Command3 
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
         Height          =   1155
         Left            =   9585
         MaskColor       =   &H00E0E0E0&
         Picture         =   "menucaja.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Salir"
         Top             =   1020
         Width           =   1230
      End
      Begin VB.Label mensaje_error 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   150
      End
   End
   Begin VB.TextBox TERMINAL 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      MaxLength       =   2
      TabIndex        =   0
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox cajero 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      MaxLength       =   11
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox turno 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   1230
      Left            =   0
      ScaleHeight     =   1170
      ScaleWidth      =   13215
      TabIndex        =   4
      Top             =   0
      Width           =   13275
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CierreCiego"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   11160
         Picture         =   "menucaja.frx":3636
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cierre de Caja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5760
         MaskColor       =   &H8000000E&
         Picture         =   "menucaja.frx":3940
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   40
         Width           =   1695
      End
      Begin VB.CommandButton image10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Tipo/Cambio"
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
         Left            =   360
         Picture         =   "menucaja.frx":57CA
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   40
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Apertura"
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
         Left            =   2160
         Picture         =   "menucaja.frx":706C
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   40
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BorraTemporal"
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
         Left            =   3960
         Picture         =   "menucaja.frx":890E
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   40
         Width           =   1695
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7560
         MaskColor       =   &H00E0E0E0&
         Picture         =   "menucaja.frx":AAC8
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Salir"
         Top             =   40
         Width           =   1575
      End
   End
   Begin GorditoButton.Boton turnoarrays 
      Height          =   555
      Index           =   0
      Left            =   4425
      TabIndex        =   49
      Top             =   3630
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   979
      PicturePosition =   0
      Caption         =   ""
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton turnoarrays 
      Height          =   570
      Index           =   1
      Left            =   5175
      TabIndex        =   50
      Top             =   3630
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      PicturePosition =   0
      Caption         =   ""
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton turnoarrays 
      Height          =   570
      Index           =   2
      Left            =   5895
      TabIndex        =   51
      Top             =   3630
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      PicturePosition =   0
      Caption         =   ""
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton turnoarrays 
      Height          =   570
      Index           =   3
      Left            =   6600
      TabIndex        =   52
      Top             =   3630
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      PicturePosition =   0
      Caption         =   ""
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton turnoarrays 
      Height          =   570
      Index           =   4
      Left            =   7335
      TabIndex        =   53
      Top             =   3630
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      PicturePosition =   0
      Caption         =   ""
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton turnoarrays 
      Height          =   570
      Index           =   5
      Left            =   8070
      TabIndex        =   54
      Top             =   3630
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      PicturePosition =   0
      Caption         =   ""
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin GorditoButton.Boton turnoarrays 
      Height          =   570
      Index           =   6
      Left            =   8790
      TabIndex        =   55
      Top             =   3630
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1005
      PicturePosition =   0
      Caption         =   ""
      BackColor       =   4210752
      ResalteColor    =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin VB.Label turnoarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   4
      Left            =   9285
      TabIndex        =   47
      Top             =   5715
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label turnoarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   8565
      TabIndex        =   46
      Top             =   5715
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label oempresa 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   44
      Top             =   2160
      Width           =   7095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EMPRESA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   43
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label turnoarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   7845
      TabIndex        =   28
      Top             =   5715
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label turnoarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   7125
      TabIndex        =   27
      Top             =   5715
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label turnoarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   0
      Left            =   6405
      TabIndex        =   26
      Top             =   5715
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   13
      Left            =   8760
      TabIndex        =   25
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   12
      Left            =   8040
      TabIndex        =   24
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   11
      Left            =   7320
      TabIndex        =   23
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   10
      Left            =   6600
      TabIndex        =   22
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   9
      Left            =   5880
      TabIndex        =   21
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   8
      Left            =   5160
      TabIndex        =   20
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   7
      Left            =   4440
      TabIndex        =   19
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   6
      Left            =   8760
      TabIndex        =   18
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   5
      Left            =   8040
      TabIndex        =   17
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   4
      Left            =   7320
      TabIndex        =   16
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   6600
      TabIndex        =   15
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   5880
      TabIndex        =   14
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   5160
      TabIndex        =   13
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label cajarray 
      BackColor       =   &H00C0C0C0&
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
      Index           =   0
      Left            =   4440
      TabIndex        =   12
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label tipoterminal 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4440
      TabIndex        =   11
      Top             =   8760
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label xactivo 
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   8760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label acu 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4800
      TabIndex        =   7
      Top             =   8760
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TURNO"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label vendedor 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   4440
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3735
      Left            =   9600
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   3180
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TERMINAL"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CONTROL DE ACCESO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   15
      TabIndex        =   2
      Top             =   1560
      Width           =   9600
   End
   Begin VB.Menu keui121 
      Caption         =   "&Menu"
      Begin VB.Menu lso23023 
         Caption         =   "&1.Apertura del Dia"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu lfohyee 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "menucaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'LEONARDO: HACER SONAR UN SONIDO
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long

Dim iResultadoSound As Variant

Private Sub clave_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Image1_Click

End Sub

Private Sub cajarray_Click(Index As Integer)

    If Len(cajarray(Index).Caption) = 0 Then Exit Sub

    'If cajarray(Index).Caption = "00" Then Exit Sub
    If terminal.Enabled = True Then
        terminal = cajarray(Index).Caption

    End If

End Sub

Private Sub cmdExit_Click()
    lfohyee_Click

End Sub

Private Sub Command1_Click()

    Dim mytablex As New ADODB.Recordset

    If Len(terminal) = 0 Then
        MsgBox "Digite un numero de Caja o Terminal", 48, "Aviso"

        If terminal.Enabled = True Then
            terminal.SetFocus

        End If

        Exit Sub

    End If

    mytablex.Open "SELECT * FROM parameca where caja='" & terminal & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    If "" & mytablex.Fields("terminal") = "T" Then
        MsgBox "No es numero de Caja ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    mytablex.Close
    tcajaper.buffer = terminal
    tcajaper.Show 1
    Exit Sub

    'tapertur.CAJA = terminal
    'tapertur.cajero = cajero
    'tapertur.Show 1
    'If terminal.Enabled = True Then
    '   terminal.SetFocus
    '   Exit Sub
    'End If
    'If turno.Enabled = True Then
    '   turno.SetFocus
    '   Exit Sub
    'End If
End Sub

Private Sub Command10_Click()
    Frame2.Visible = False

End Sub

Private Sub Command11_Click()

    Dim found As Integer

    If Val(compra) = 0 Then
        compra.SetFocus
        Exit Sub

    End If

    If Val(venta) = 0 Then
        venta.SetFocus
        Exit Sub

    End If

    found = busca_parame1(1)

    If found = 0 Then Exit Sub
    Frame2.Visible = False

End Sub

Private Sub Command12_Click()

End Sub

Private Sub Command2_Click()
    tcajacie.Show 1

End Sub

Private Sub Command3_Click()
    lfohyee_Click

End Sub

Private Sub dia_Click()

End Sub

Private Sub Command7_Click()

    Dim found As Integer

    If Len(terminal) = 0 Then
        MsgBox "Digite un numero de Caja o Terminal", 48, "Aviso"

        If terminal.Enabled = True Then
            terminal.SetFocus

        End If

        Exit Sub

    End If

    found = copiar_deliveri("" & terminal) 'el numero de terminal

    If found = 0 Then
        MsgBox "Error al copiar archivo temporal,Esta en Uso ", 24, "Aviso"
        End
        Exit Sub

    End If

    MsgBox "Proceso Realizado " + terminal, 58, "Aviso"

    If terminal.Enabled = True Then
        terminal.SetFocus

    End If

    Exit Sub

End Sub

Private Sub Command8_Click()

    Dim bfecha   As String

    Dim mytablex As New ADODB.Recordset

    If Len(terminal) = 0 Then
        MsgBox "Ingreso Caja ", 48, "Aviso"
        Exit Sub

    End If

    flag_clave1 = 0
    tconcla.X = "CIERRE"
    tconcla.Show 1

    If flag_clave1 = 0 Then  'si es descongela
        Exit Sub

    End If

    mytablex.Open "select * from apertura where caja='" & terminal & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "Caja No aperturado ", 48, "Mensaje"
        mytablex.Close
        Exit Sub

    End If

    bfecha = Format("" & mytablex.Fields("fechai"), "dd/mm/yyyy")
    mytablex.Close
    opcion1 = "5"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
   
    usuariopos = gusuario
    'tcuadrc1.tipoexterno.Visible = True
    'tcuadrc1.numcuadre.Visible = True
    'tcuadrc1.flagdiario = "1"
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = terminal
    tcuadrc1.turno = "%"
    tcuadrc1.fechai = Format(bfecha, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(bfecha, "dd/mm/yyyy")
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.Caption = "CIERRE DEL DIA"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

End Sub

Private Sub Form_Activate()
    Frame2.Top = 10: Frame2.Left = 10
    Frame3.Top = 10: Frame3.Left = 10

    'If xactivo = "" And acu = "T" Then
    'Label5_Click
    'End If
    'xactivo = "S"
    visualiza_mesa

    oempresa = globalemp
    cargas_iniciales
    cargar_grafico20

    Dim I As Integer

    For I = 0 To 4
        'turnoarrays(i).Sound = App.path & "\Sonido\producto.wav": turnoarrays(i).PlaySound = InClick
        'turnoarrays(i).Sound = "C:\Windows\producto.wav": turnoarrays(i).PlaySound = InClick
    Next

    'cmdIngresar.Sound = App.path & "\Sonido\cash.wav": cmdIngresar.PlaySound = InClick
End Sub

Sub visualiza_mesa()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        glomesa = Trim("" & mytablex.Fields("mesaseccion"))

    End If

    mytablex.Close

    If Len(Trim(glomesa)) = 0 Then
        glomesa = "Mesa"

    End If

End Sub

Sub cargar_grafico20()

    On Error GoTo cmd7779_err

    'Exit Sub
    Image1.Picture = LoadPicture(globalpath & "\ico\tpv1.jpg")
    Exit Sub
cmd7779_err:
    'MsgBox " Carga Grafico:" & error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Load()

    cajero = gusuario

End Sub

Sub cargas_iniciales()

    Dim cad      As String

    Dim I        As Integer

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    For I = 0 To 13
        cajarray(I).Caption = ""
    Next I

    For I = 0 To 3
        'turnoarray(i).Caption = ""
        turnoarrays(I).Caption = ""
    Next I

    found = abrir_caja_defecto()

    If found = 1 Then
        terminal.Enabled = False

        'Image1_Click
        'Exit Sub
    End If

    I = 0

    If terminal.Enabled = True Then
        cad = "SELECT * FROM parameca order by caja   "
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Do

            If mytablex.EOF Then Exit Do
            If turno.Visible = True Then
                If "" & mytablex.Fields("terminal") = "C" Then

                    'If "" & mytablex.Fields("caja") <> "00" Then
                    If Len(Trim("" & mytablex.Fields("caja"))) > 0 Then
                        If I < 14 Then
                            cajarray(I).Caption = "" & mytablex.Fields("caja")

                        End If

                        I = I + 1

                    End If

                    'End If
                End If

            End If

            If turno.Visible = False Then
                If "" & mytablex.Fields("terminal") = "T" Then
                    If Len(Trim("" & mytablex.Fields("caja"))) > 0 Then
                        cajarray(I).Caption = "" & mytablex.Fields("caja")
                        I = I + 1

                    End If

                End If

            End If
   
            mytablex.MoveNext
        Loop
        mytablex.Close

    End If
   
    I = 0

    If turno.Visible = True Then
        cad = "SELECT * FROM turno   order by turno "
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Do

            If mytablex.EOF Then Exit Do
            If Trim("" & mytablex.Fields("turno")) > 0 Then
                'turnoarray(i) = "" & mytablex.Fields("turno")
                turnoarrays(I).Caption = "" & mytablex.Fields("turno")
                I = I + 1

            End If
   
            mytablex.MoveNext
        Loop
        mytablex.Close

    End If

End Sub

Function validar_cajas() As Long

    Dim cad      As String

    Dim mytablex As New ADODB.Recordset

    Dim indx     As Long

    indx = 0
   
    cad = "SELECT * FROM parameca   "

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If

    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("terminal") = "C" Then
            indx = indx + 1

        End If

        mytablex.MoveNext
    Loop
    validar_cajas = indx
    mytablex.Close

End Function

Private Sub Image1_Click()
    'HACEMOS SONAR
    'iResultadoSound = mciExecute("Play " & App.path & "\Sonido\producto.wav")
    'iResultadoSound = mciExecute("Play " & "C:\Windows\producto.wav")

    Dim found As Integer

    Dim buf   As String

    Dim cam   As String

    rrlocal11 = ""
    rrtipo = ""
    rrserie = ""
    rrnumero = ""

    If Frame2.Visible = True Then
        compra.SetFocus
        Exit Sub

    End If

    'si es terminal
    If Len(terminal) = 0 Then
        If terminal.Enabled = True Then
            terminal.SetFocus

        End If

        Exit Sub

    End If

    'If terminal = "00" Then
    '   terminal = ""
    '   If terminal.Enabled = True Then
    '      terminal.SetFocus
    '      Else
    '      turno.SetFocus
    '   End If
    '   Exit Sub
    'End If
    'cam = serie_disco_duro()
    'if busca_seriehd(buf, buf1)

    'If validar_cajas() > 1 Then
    '   MsgBox "Numero de Cajas excede la Licencia", 48, "Aviso"
    '   terminal.SetFocus
    '   Exit Sub
    'End If

    'cn.Close
    'found = conectar(oempresa)
    'If found = 0 Then
    '   MsgBox "Error de Conexion Sql Server ", 48, "Aviso"
    '   Exit Sub
    'End If

    found = busca_terminal()

    If found = 2 Then
        MsgBox "Caja/Terminal No configurado debidamente,llame al administrador", 48, "Aviso"
        terminal = ""
        Exit Sub

    End If

    If found = 3 Then
        MsgBox "Caja/Terminal deshabilitado,llame al administrador", 48, "Aviso"
        terminal = ""
        Exit Sub

    End If

    If found = 0 Then
        MsgBox "Caja/Terminal No existe", 48, "Aviso"
        terminal = ""
        Exit Sub

    End If

    'validamos si esta prohibido entrara a estas partes
    'If busca_seriehd("" & terminal) = 0 Then
    '   MsgBox "Caja no permitido para esta maquina!!! ", 48, "Aviso"
    '   Exit Sub
    'End If
    found = busca_si_ingresa("" & cajero)

    If found = 0 Then
        MsgBox "Acceso Restringido,Consulte con su administrador ", 48, "Aviso"
        terminal = ""

        If terminal.Enabled = True Then
            terminal.SetFocus
      
        End If

        Exit Sub

    End If

    dia = Format(Now, "dd/mm/yyyy")

    If acu = "C" Then
        If Len(turno) = 0 Then
            turno.SetFocus
            Exit Sub

        End If

        found = busca_turno()

        If found = 0 Then
            turno = ""
            turno.SetFocus
            Exit Sub

        End If

        found = busca_apertura()

        If found = 0 Then
            buf = "LA CAJA NRO:" & terminal & Chr$(10) & Chr$(13)
            buf = buf & "TURNO          :" & turno & Chr$(10) & Chr$(13)
            buf = buf & "CAJERO         :" & cajero & Chr$(10) & Chr$(13)
            buf = buf & "NO SE ENCUENTRA APERTURADO                           " & Chr$(10) & Chr$(13)
            buf = buf & "PARA EL DIA " & Format(Now, "dd/mm/yyyy")
            mensaje_error = buf
            Frame3.Visible = True
            Command3.SetFocus
            Exit Sub

        End If

        If valida_fecha(dia) = 0 Then
            MsgBox "Error en apertura del dia ", 48, "Aviso"
            Exit Sub

        End If

    End If

    If mytable11.State = 1 Then mytable11.Close
    mytable11.Open "SELECT * FROM parameca where caja='" & terminal & "'", cn, adOpenStatic, adLockOptimistic

    If mytable11.EOF = True Or mytable11.BOF = True Then
        mytable11.Close
        Exit Sub

    End If

    If Len("" & mytable11.Fields("local")) = 0 Then
        MsgBox "NO existe Local configurado en parametros ", 48, "Aviso"
        mytable11.Close
        Exit Sub

    End If

    If Len("" & mytable11.Fields("bodega")) = 0 Then
        MsgBox "NO existe Bodega configurado en parametros ", 48, "Aviso"
        mytable11.Close
   
        Exit Sub

    End If

    'If "" & mytable11.Fields("listap") <> "01" And "" & mytable11.Fields("listap") <> "02" And "" & mytable11.Fields("listap") <> "03" And "" & mytable11.Fields("listap") <> "04" Then
    '   MsgBox "NO existe Lista Precios  en Parametros", 48, "Aviso"
    '   mytable11.Close
    'clave.Enabled = True
    'clave.SetFocus
    '   Exit Sub
    'End If
    ingreso_terminal
    mytable11.Close
    menucaja.Hide
    Unload menucaja

End Sub

Private Sub Image10_Click()

    Dim found As Integer

    compra = ""
    venta = ""
    found = busca_parame1(0)

    If found = 0 Then Exit Sub
    Frame2.Visible = True
    compra.SetFocus

End Sub

Private Sub Image3_Click()

End Sub

Private Sub Image17_Click()

End Sub

Private Sub Image18_Click()

End Sub

Private Sub Image19_Click()

End Sub

Private Sub Image20_Click()

End Sub

Private Sub Image8_Click()

End Sub

Private Sub lfohyee_Click()

    If Frame3.Visible = True Then
        Frame3.Visible = False
        Exit Sub

    End If

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    menucaja.Hide
    Unload menucaja

End Sub

Private Sub lso23023_Click()
    Frame3.Visible = False
    Command1_Click

End Sub

Private Sub terminal_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    terminal = UCase(terminal)
    Image1_Click

End Sub

Function busca_terminal()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If "" & mytablex.Fields("estadocaduca") = "S" Then
            If CVDate(Format("" & mytablex.Fields("caduca"), "dd/mm/yyyy")) <= CVDate(Now) Then
                busca_terminal = 3
                mytablex.Close
                Exit Function

            End If

        End If

    End If

    mytablex.Close

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM parameca where caja='" & terminal & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True Or mytablex.BOF = True Then
        mytablex.Close
        Exit Function

    End If

    busca_terminal = 1

    If "" & mytablex.Fields("terminal") <> acu Then
        busca_terminal = 2

    End If

    If "" & mytablex.Fields("deshab") = "S" Then
        busca_terminal = 3

    End If

    mytablex.Close
 
End Function

Function busca_turno()

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM turno where turno='" & turno & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True Or mytablex.BOF = True Then
        mytablex.Close
        Exit Function

    End If

    busca_turno = 1
    mytablex.Close
 
End Function

Function busca_apertura()

    Dim mytablex As New ADODB.Recordset

    Dim fechag   As String

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM parameca where caja='" & terminal & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If

    dia = ""

    If "" & mytablex.Fields("apertura") = "N" Then
        dia = Format(Now, "dd/mm/yyyy")
        busca_apertura = 1
        mytablex.Close
        Exit Function

    End If

    If mytablex.State = 1 Then mytablex.Close
    'mytablex.Open "SELECT * FROM apertura where cajero='" & cajero & "' and caja='" & TERMINAL & "' and turno='" & turno & "'", cn, adOpenStatic, adLockOptimistic
    mytablex.Open "SELECT * FROM apertura where caja='" & terminal & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If

    fechag = Format(Now, "dd/mm/yyyy")

    If turno123() = "2" Then  'si es tercer turno flag 2
        If CVDate(Format("" & mytablex.Fields("fechai"), "dd/mm/yyyy")) <= CVDate(fechag) Then
            dia = "" & mytablex.Fields("fechai")
            busca_apertura = 1

        End If

    Else

        If CVDate(Format("" & mytablex.Fields("fechai"), "dd/mm/yyyy")) <= CVDate(fechag) And CVDate(Format("" & mytablex.Fields("fechaf"), "dd/mm/yyyy")) >= CVDate(fechag) Then
            dia = "" & mytablex.Fields("fechai")
            busca_apertura = 1

        End If

    End If
   
    mytablex.Close
 
End Function

Sub ingreso_terminal()

    Dim found As Integer

    gocabeza = "factura"
    godetalle = "detalle"
    gofpago = "fpagov"

    If acu = "T" Then  'terminal
        gocabeza = "cproform"
        godetalle = "dproform"
        gofpago = "FPAGOV"

    End If

    If acu = "C" Then  'caja registradora
        If busca_parameg() = "D" Then
            gocabeza = "cadiario"
            godetalle = "dediario"
            gofpago = "fpdiario"

        End If

    End If

    'aqui se pone si se desea que guarde el pedido
    'found = copiar_deliveri("" & TERMINAL) 'el numero de terminal
    'If found = 0 Then
    '   MsgBox "Error al copiar archivo temporal,Esta en Uso ", 24, "Aviso"
    '   End
    '   Exit Sub
    'End If
    fpusuario = "_f" + gusuario
    found = copiar_tmpfpago("" & fpusuario)

    If found = 0 Then
        MsgBox "No se puede copiar temporal fpago ", 48, "Aviso"
        Exit Sub

    End If

    cgusuario = gocabeza
    dgusuario = "_z" & terminal
    dgusuariog = godetalle

    '-------------------------
    If "" & mytable11.Fields("parqueo") = "S" Then  'S PARQUEO  'T TOUCH
        tipoterminal = "PARKING"

    End If

    If "" & mytable11.Fields("parqueo") = "T" Then  'S PARQUEO  'T TOUCH
        tipoterminal = "TOUCH"

    End If

    If "" & mytable11.Fields("parqueo") = "M" Then  'M MINIMARKET  'T TOUCH
        tipoterminal = "MINIMARKET"

    End If

    If "" & mytable11.Fields("parqueo") = "A" Then  'M MINIMARKET  'T TOUCH
        tipoterminal = "ANDROID"

    End If

    If tipoterminal = "ANDROID" Then
        'tandroid.caja = terminal
        'tandroid.turno = turno
        'tipo_servicio = acu
        'tandroid.cajero = gusuario
        'tandroid.Show 1
        'terminal = ""
        'turno = ""
        Exit Sub

    End If

    If tipoterminal = "PARKING" Then
        parking.caja = terminal
        parking.turno = turno
        parking.cajero = gusuario
        parking.Show 1
        terminal = ""
        turno = ""
        Exit Sub

    End If

    '-------------------------
    If tipoterminal <> "TOUCH" And tipoterminal <> "TOUCH2" Then
        If acu = "C" Then
            tdeliver.Caption = "Caja Registradora"
        Else
            tdeliver.Caption = "Caja Terminal Pedidos"

        End If

    End If

    If tipoterminal = "TOUCH" Then
        If acu = "C" Then
            tptovta.Caption = "Caja Registradora"
        Else
            tptovta.Caption = "Caja Terminal Pedidos"

        End If

    End If

    If tipoterminal = "TOUCH2" Then
        If acu = "C" Then
            'tptovtaa.Caption = "Caja Registradora"
        Else

            'tptovtaa.Caption = "Caja Terminal Pedidos"
        End If

    End If

    If tipoterminal <> "TOUCH" And tipoterminal <> "TOUCH2" Then
        If "" & mytable11.Fields("hdetraccio") <> "S" Then
            tdeliver.Label7.Enabled = False
            tdeliver.Label7.Caption = "No.Cobrar.Detraccion"

        End If

        tdeliver.caja = terminal
        tdeliver.turno = turno
        tipo_servicio = acu
        tdeliver.Show 1
        terminal = ""
        turno = ""

    End If

    If tipoterminal = "TOUCH" Then
        tptovta.caja = terminal
        tptovta.turno = turno
        tipo_servicio = acu
        tptovta.cajero = gusuario
        tptovta.Show 1
        terminal = ""
        turno = ""

    End If

    If tipoterminal = "TOUCH2" Then

        'tptovtaa.caja = terminal
        'tptovtaa.turno = turno
        'tipo_servicio = acu
        'tptovtaa.cajero = gusuario
        'tptovtaa.Show 1
        'terminal = ""
        'turno = ""
    End If

End Sub

Public Function ArchivoEnUso(ByVal sFileName As String) As Boolean
    ArchivoEnUso = False

End Function

Function busca_parameg() As String

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True Or mytablex.BOF = True Then
        mytablex.Close
        Exit Function

    End If

    busca_parameg = "" & mytablex.Fields("tradiario")
    mytablex.Close
 
End Function

Function busca_parame1(sw As Integer)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM parame where codigo='01' ", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True Or mytablex.BOF = True Then
        mytablex.Close
        Exit Function

    End If

    If sw = 0 Then
        compra = "" & Val("" & mytablex.Fields("paricomp"))
        venta = "" & Val("" & mytablex.Fields("parivta"))

    End If

    If sw = 1 Then
        'mytablex.Edit
        mytablex.Fields("paricomp") = Val(compra)
        mytablex.Fields("parivta") = Val(venta)
        mytablex.Update

    End If

    busca_parame1 = 1

    mytablex.Close
 
End Function

Private Sub turno_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(terminal) = 0 Then
        If terminal.Enabled = True Then
            terminal.SetFocus

        End If

        Exit Sub

    End If

    ' borrar_archivo "_z" & terminal & ".dbf"
    ' borrar_archivo "_z" & terminal & ".cdx"

    'debo ponerlo en parametros ojo

    'found = copiar_deliveri("" & terminal) 'el numero de terminal
    'If found = 0 Then
    '   MsgBox "Otra caja Igual se encuentra aperturado ", 24, "Aviso"
    '   terminal.SetFocus
    '   Exit Sub
    'End If
    Image1_Click

End Sub

Private Sub turno_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If terminal.Enabled = True Then
            terminal.SetFocus
      
        End If

        Exit Sub

    End If

End Sub

Function busca_si_ingresa(buf As String)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM vendedor where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True Or mytablex.BOF = True Then
        mytablex.Close
        Exit Function

    End If
   
    If "" & mytablex.Fields("parame") <> "S" Then
        busca_si_ingresa = 1
        GoTo sam

    End If

    If Len("" & mytablex.Fields("parame1")) > 0 And "" & mytablex.Fields("parame1") = "" & terminal Then
        busca_si_ingresa = 1

    End If

    If Len("" & mytablex.Fields("parame2")) > 0 And "" & mytablex.Fields("parame2") = "" & terminal Then
        busca_si_ingresa = 1

    End If

    If Len("" & mytablex.Fields("parame3")) > 0 And "" & mytablex.Fields("parame3") = "" & terminal Then
        busca_si_ingresa = 1

    End If

    If Len("" & mytablex.Fields("parame4")) > 0 And "" & mytablex.Fields("parame4") = "" & terminal Then
        busca_si_ingresa = 1

    End If

sam:
    mytablex.Close

End Function

Function busca_seriehd(buf As String)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM parameca where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.EOF = True Or mytablex.BOF = True Then
        mytablex.Close
        Exit Function

    End If

    'MsgBox serial_number & " " & "" & mytablex.Fields("seriehd")
    If Len(serial_number) > 0 Then
        If serial_number = "" & mytablex.Fields("seriehd") Then
            busca_seriehd = 1

        End If

    End If

    mytablex.Close

End Function

Private Sub turnoarrays_Click(Index As Integer)

    'If Len(turnoarray(Index).Caption) = 0 Then Exit Sub
    'turno = turnoarray(Index).Caption
    If Len(turnoarrays(Index).Caption) = 0 Then Exit Sub
    turno = turnoarrays(Index).Caption

End Sub

Function abrir_caja_defecto()

    Dim DATO  As String

    Dim xcan  As Integer

    Dim found As Integer

    On Error GoTo cmd1145_err

    If Dir$(globalpath & "\caja.txt") <> "" Then
        xcan = FreeFile
        Open globalpath & "\caja.txt" For Input As #xcan
        Input #xcan, DATO
        Close #xcan
        'MsgBox dato
        found = busca_caja(DATO)

        If found = 1 Then
            cajarray(0) = Trim("" & DATO)
            terminal = Trim("" & DATO)
            abrir_caja_defecto = 1

        End If

    End If

    Exit Function
cmd1145_err:
    Exit Function

End Function

Function busca_caja(buf As String)

    Dim mytablex As New ADODB.Recordset

    turno = ""

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM parameca where caja='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_caja = 1

    End If

    mytablex.Close

End Function

Function turno123()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM turno where turno='" & turno & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        turno123 = "" & mytablex.Fields("flag")

    End If

    mytablex.Close

End Function

