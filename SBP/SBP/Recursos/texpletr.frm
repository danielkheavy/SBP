VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form texpletr 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Letras"
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   14805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Protesto"
      Height          =   3135
      Left            =   2760
      TabIndex        =   77
      Top             =   2160
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox estadop 
         Height          =   495
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   81
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox fechavp 
         Height          =   495
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   80
         Top             =   1080
         Width           =   1935
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
         Height          =   735
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "texpletr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
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
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "texpletr.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha vencimiento"
         Height          =   495
         Left            =   120
         TabIndex        =   84
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Protesto"
         Height          =   495
         Left            =   120
         TabIndex        =   83
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "0.NoProtesto 1.Protesto"
         Height          =   195
         Left            =   2040
         TabIndex        =   82
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Renovacion de Letra"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8055
      Left            =   600
      TabIndex        =   26
      Top             =   1440
      Visible         =   0   'False
      Width           =   12135
      Begin VB.CommandButton Command3 
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
         Left            =   9000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "texpletr.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
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
         Left            =   9000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "texpletr.frx":170A
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox negociado 
         Enabled         =   0   'False
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
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   46
         Top             =   6360
         Width           =   1575
      End
      Begin VB.TextBox ochodia 
         Enabled         =   0   'False
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   45
         Top             =   6000
         Width           =   1575
      End
      Begin VB.TextBox nrounico 
         Enabled         =   0   'False
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
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   44
         Top             =   5640
         Width           =   1575
      End
      Begin VB.TextBox seccionr 
         Enabled         =   0   'False
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
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   43
         Top             =   4920
         Width           =   1575
      End
      Begin VB.TextBox abono 
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
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   42
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox observa 
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
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   41
         Top             =   7080
         Width           =   5295
      End
      Begin VB.TextBox agencia 
         Enabled         =   0   'False
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
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   40
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox bancor 
         Enabled         =   0   'False
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
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   39
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox girador 
         Enabled         =   0   'False
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
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   38
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox aceptante 
         Enabled         =   0   'False
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
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   37
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox paridad 
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   36
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox importe 
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
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   35
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox fechaef 
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   34
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox fechaei 
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   33
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox letran 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   32
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox moneda 
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
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   31
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox letrar 
         Enabled         =   0   'False
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
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox saldoa 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox estador 
         Enabled         =   0   'False
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
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   28
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox estadopp 
         Enabled         =   0   'False
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
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   27
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Negociado"
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
         Left            =   240
         TabIndex        =   76
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Octavo Dia"
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
         Left            =   240
         TabIndex        =   75
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro.Unico"
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
         Left            =   240
         TabIndex        =   74
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abono"
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
         Left            =   4080
         TabIndex        =   73
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observacion"
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
         Left            =   240
         TabIndex        =   72
         Top             =   7080
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Agencia"
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
         Left            =   240
         TabIndex        =   71
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
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
         Left            =   240
         TabIndex        =   70
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Girador"
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
         Left            =   240
         TabIndex        =   69
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seccion"
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
         Left            =   240
         TabIndex        =   68
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aceptante"
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
         Left            =   240
         TabIndex        =   67
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T.Cambio"
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
         Left            =   240
         TabIndex        =   66
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
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
         Left            =   240
         TabIndex        =   65
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label29 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe"
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
         Left            =   4080
         TabIndex        =   64
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Vencimiento"
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
         Left            =   240
         TabIndex        =   63
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Emision"
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
         Left            =   240
         TabIndex        =   62
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NuevaLetraRenovada"
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
         Left            =   240
         TabIndex        =   61
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Letra a Renovar"
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
         Left            =   240
         TabIndex        =   60
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo Anterior"
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
         Left            =   4080
         TabIndex        =   59
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
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
         Left            =   240
         TabIndex        =   58
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label nombrea 
         BackColor       =   &H00FFFF00&
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
         Height          =   375
         Left            =   2400
         TabIndex        =   57
         Top             =   3480
         Width           =   5295
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
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
         Left            =   240
         TabIndex        =   56
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label nombreg 
         BackColor       =   &H00FFFF00&
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
         Height          =   375
         Left            =   2400
         TabIndex        =   55
         Top             =   4200
         Width           =   5295
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
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
         Left            =   240
         TabIndex        =   54
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label negociadon 
         BackColor       =   &H00FFFF00&
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
         Height          =   375
         Left            =   2400
         TabIndex        =   53
         Top             =   6720
         Width           =   5295
      End
      Begin VB.Label nseccion 
         BackColor       =   &H00FFFF00&
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
         Height          =   375
         Left            =   4080
         TabIndex        =   52
         Top             =   4920
         Width           =   3615
      End
      Begin VB.Label nbanco 
         BackColor       =   &H00FFFF00&
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
         Height          =   375
         Left            =   4080
         TabIndex        =   51
         Top             =   4560
         Width           =   3615
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Renova Anterior"
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
         Left            =   4080
         TabIndex        =   50
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Esta con Protesto"
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
         Left            =   4080
         TabIndex        =   49
         Top             =   1920
         Width           =   2175
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   8415
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   14843
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "X"
         Caption         =   "X"
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
         DataField       =   "Estado"
         Caption         =   "Estado"
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
         DataField       =   "Seccion"
         Caption         =   "Seccion"
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
      BeginProperty Column03 
         DataField       =   "Letra"
         Caption         =   "Letra"
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
      BeginProperty Column04 
         DataField       =   "Fechai"
         Caption         =   "Fechai"
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
      BeginProperty Column05 
         DataField       =   "Fechaf"
         Caption         =   "Fechaf"
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
      BeginProperty Column06 
         DataField       =   "Moneda"
         Caption         =   "M"
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
      BeginProperty Column07 
         DataField       =   "Saldo"
         Caption         =   "Saldo"
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
      BeginProperty Column08 
         DataField       =   "Abono"
         Caption         =   "Abono"
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
      BeginProperty Column09 
         DataField       =   "Importe"
         Caption         =   "Importe"
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
      BeginProperty Column10 
         DataField       =   "Observa"
         Caption         =   "Observa"
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
      BeginProperty Column11 
         DataField       =   "Estador"
         Caption         =   "Renovado"
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
      BeginProperty Column12 
         DataField       =   "Estadop"
         Caption         =   "Protesto"
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
            ColumnWidth     =   239.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   555.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   255.118
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2564.788
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   689.953
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFF80&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   14745
      TabIndex        =   0
      Top             =   0
      Width           =   14805
      Begin VB.ComboBox Local1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox condicion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox seccion 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   0
         Width           =   1455
      End
      Begin VB.ComboBox banco 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   5880
         MaxLength       =   11
         TabIndex        =   22
         Text            =   "%"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox nombre 
         Height          =   375
         Left            =   5880
         MaxLength       =   11
         TabIndex        =   18
         Text            =   "%"
         Top             =   0
         Width           =   1455
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Refresca"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "texpletr.frx":1EB8
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   0
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
         Picture         =   "texpletr.frx":2666
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "texpletr.frx":3878
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "texpletr.frx":4A8A
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "texpletr.frx":5C9C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Borrar registro"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label35 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   7560
         TabIndex        =   88
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label28 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Condicion"
         Height          =   375
         Left            =   7560
         TabIndex        =   86
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aceptante"
         Height          =   375
         Left            =   4800
         TabIndex        =   23
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seccion"
         Height          =   375
         Left            =   7560
         TabIndex        =   21
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
         Height          =   375
         Left            =   4800
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Girador"
         Height          =   375
         Left            =   4800
         TabIndex        =   19
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   1560
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado Letra"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Label dolares 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   9720
      Width           =   2055
   End
   Begin VB.Label Label36 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Dolares"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   9720
      Width           =   2055
   End
   Begin VB.Label soles 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   9720
      Width           =   2055
   End
   Begin VB.Label Label34 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Soles"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   9720
      Width           =   2055
   End
   Begin VB.Label acu 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Menu nier33 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu dkjbo7232 
      Caption         =   "&Borra"
   End
   Begin VB.Menu moifd343 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu dkireno5 
      Caption         =   "&Renovacion"
   End
   Begin VB.Menu dfju8343 
      Caption         =   "&Protesto"
   End
   Begin VB.Menu fj7834 
      Caption         =   "&Cobrar"
   End
   Begin VB.Menu dkii2232 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu lo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "texpletr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xnameclie As String
Dim tbasele As New adodb.Recordset

Private Sub cmdCancelar_Click()
End Sub

Private Sub cmdDelete_Click()
dkjbo7232_Click
End Sub

Private Sub cmdGrabar_Click()
End Sub

Private Sub cmdPrint_Click()
REPLETRA.titulo = "Letras "
REPLETRA.acu = acu
REPLETRA.Show 1

End Sub

Private Sub Command1_Click()
If estadop <> "0" And estadop <> "1" Then
   estadop = ""
   estadop.SetFocus
   Exit Sub
End If
If Len(fechavp) <> 10 Then
   fechavp.SetFocus
   Exit Sub
End If
If Not IsDate(fechavp) Then
   fechavp.SetFocus
   Exit Sub
End If
If MsgBox("Desea Grabar el protesto", 1, "Aviso") <> 1 Then Exit Sub
'tbasele.Edit
tbasele.Fields("estadop") = "" & estadop
tbasele.Fields("fechavp") = "" & fechavp
tbasele.Update
lo232_Click
Exit Sub
End Sub

Private Sub Command2_Click()
lo232_Click
End Sub

Private Sub Command5_Click()
sql_letras
End Sub

Private Sub Command3_Click()
Dim found As Integer
suma_letra
found = valida()
If found = 0 Then
   MsgBox "Datos erroneos", 48, "Aviso"
   Exit Sub
End If
If MsgBox("Desea Grabar la Renovacion", 1, "Aviso") <> 1 Then Exit Sub
found = grabar_letra()
lo232_Click
End Sub
Function valida()
Dim found As Integer
If Len(letran) = 0 Then
   letran.SetFocus
   Exit Function
End If
If Not IsDate(fechaei) Or Len(fechaei) <> 10 Then
   fechaei = ""
   fechaei.SetFocus
   Exit Function
End If
If Not IsDate(fechaef) Or Len(fechaef) <> 10 Then
   fechaef = ""
   fechaef.SetFocus
   Exit Function
End If
If moneda <> "S" And moneda <> "D" Then
   moneda.SetFocus
   Exit Function
End If
'ochodia = Format(CVDate(fechaef) + 8, "dd/mm/yyyy")
'If Len(aceptante) = 0 Then
'   aceptante.SetFocus
'   Exit Function
'End If
'found = busca_codigo("" & aceptante, 0)
'If found = 0 Then
'   MsgBox "No existe aceptante", 48, "Aviso"
'   aceptante.SetFocus
'   Exit Function
'End If
'If Len(girador) = 0 Then
'   girador.SetFocus
'   Exit Function
'End If
'found = busca_codigo("" & girador, 1)
'If found = 0 Then
'   MsgBox "No existe girador", 48, "Aviso"
'   girador.SetFocus
'   Exit Function
'End If
'If Len(bancor) > 0 Then
'   found = busca_banco()
'   If found = 0 Then
'      MsgBox "No existe Banco", 48, "Aviso"
'      bancor = ""
'      bancor.SetFocus
'      Exit Function
'   End If
'End If
'If Len(seccionr) = 0 Then
'   seccionr.SetFocus
'   Exit Function
'End If
'found = busca_seccion()
'If found = 0 Then
'   seccionr = ""
'   seccionr.SetFocus
'   Exit Function
'End If
found = valida_letrar("" & letran)
If found = 1 Then
   MsgBox "Ya existe el numero de letra", 48, "Aviso"
   Exit Function
End If
valida = 1
End Function

Private Sub Command4_Click()
lo232_Click
End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If ColIndex > 0 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 0
            If dbgrid2.columns(0) <> "1" And dbgrid2.columns(0) <> "0" Then
               Cancel = True
               Exit Sub
            End If
End Select
            
End Sub

Private Sub DBGrid2_DblClick()
On Error GoTo cmd43_err
If Trim("" & tbasele.Fields("x")) <> "S" Then
   tbasele.Fields("X") = "S"
   tbasele.Update
   Exit Sub
End If
If "" & tbasele.Fields("x") = "S" Then
   tbasele.Fields("X") = ""
   tbasele.Update
   Exit Sub
End If

Exit Sub
cmd43_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub dfju8343_Click()
On Error GoTo cmd3_err
If Frame4.Visible = True Then Exit Sub
estadop = "" & tbasele.Fields("estadop")
fechavp = "" & tbasele.Fields("fechavp")
Frame3.Visible = True
estadop.SetFocus
Exit Sub
cmd3_err:
Exit Sub
End Sub

Private Sub dkii2232_Click()

'If Frame3.Visible = True Then Exit Sub
'If Frame4.Visible = True Then Exit Sub
'reporgen.NAMETABLA = xcuentaco
'reporgen.Show 1

exporta_excel

End Sub
Sub exporta_excel()
Dim v As Long
Dim h As Long
Dim i As Long
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double

Dim Heading(16) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
Dim buf As String
    Heading(1) = "Seccion"
    Heading(2) = "Letra"
    Heading(3) = "Fechai"
    Heading(4) = "Fechaf"
    Heading(5) = "Moneda"
    Heading(6) = "Saldo"
    Heading(7) = "Abono"
    Heading(8) = "Importe"
    Heading(9) = "Observa"
    Heading(10) = "Renovado"
    Heading(11) = "Protesto"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
With objExcel.ActiveSheet
        
    For i = 1 To 15 Step 1
        .Cells(1, i) = Heading(i)
    Next i
       
        .columns("A").ColumnWidth = 5
        .columns("B").ColumnWidth = 10
        .columns("C").ColumnWidth = 10
        .columns("D").ColumnWidth = 10
        .columns("E").ColumnWidth = 5
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 10
        .columns("H").ColumnWidth = 10
        .columns("I").ColumnWidth = 10
        .columns("J").ColumnWidth = 5
        .columns("K").ColumnWidth = 5
        
End With
sdx = 0
sdx1 = 0
sdx2 = 0
   
v = 2
h = 1
     Do
     If tbasele.EOF Then Exit Do
            objExcel.ActiveSheet.Cells(v, h) = "'" & tbasele.Fields("Seccion")
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & tbasele.Fields("Letra")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & tbasele.Fields("Fechai")
            objExcel.ActiveSheet.Cells(v, h + 3) = "'" & tbasele.Fields("Fechaf")
            objExcel.ActiveSheet.Cells(v, h + 4) = "'" & tbasele.Fields("Moneda")
             objExcel.ActiveSheet.Cells(v, h + 5) = "" & tbasele.Fields("saldo")
            objExcel.ActiveSheet.Cells(v, h + 6) = "" & tbasele.Fields("abono")
            objExcel.ActiveSheet.Cells(v, h + 7) = "" & tbasele.Fields("importe")
            objExcel.ActiveSheet.Cells(v, h + 8) = "" & tbasele.Fields("Observa")
            objExcel.ActiveSheet.Cells(v, h + 9) = "'" & tbasele.Fields("estador")
            objExcel.ActiveSheet.Cells(v, h + 10) = "'" & tbasele.Fields("estadop")
            sdx = sdx + Val("" & tbasele.Fields("saldo"))
            sdx1 = sdx1 + Val("" & tbasele.Fields("abono"))
            sdx2 = sdx2 + Val("" & tbasele.Fields("importe"))
            v = v + 1
     tbasele.MoveNext
     Loop
     objExcel.ActiveSheet.Cells(v, 6) = "" & sdx
     objExcel.ActiveSheet.Cells(v, 7) = "" & sdx1
     objExcel.ActiveSheet.Cells(v, 8) = "" & sdx2
  Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
 
End Sub

Private Sub dkireno5_Click()
Dim found As Integer
On Error GoTo cmd6_err
If Frame3.Visible = True Then Exit Sub
letran = ""
nseccion = ""
nbanco = ""
saldoa = "" & tbasele.Fields("saldo")
letrar = "" & tbasele.Fields("letra")
negociado = "" & tbasele.Fields("negociado")
nrounico = "" & tbasele.Fields("nrounico")
ochodia = "" & tbasele.Fields("ochodia")
fechaei = "" & tbasele.Fields("fechai")
fechaef = "" & tbasele.Fields("fechaf")
importe = "" & tbasele.Fields("importe")
moneda = "" & tbasele.Fields("moneda")
paridad = "" & tbasele.Fields("paridad")
aceptante = "" & tbasele.Fields("aceptante")
nombrea = "" & tbasele.Fields("nombrea")
girador = "" & tbasele.Fields("girador")
nombreg = "" & tbasele.Fields("nombreg")
bancor = "" & tbasele.Fields("banco")
seccionr = "" & tbasele.Fields("seccion")
agencia = "" & tbasele.Fields("agencia")
'refactura = "" & tbasele.Fields("refactura")
observa = "" & tbasele.Fields("observa")
estadop = "" & tbasele.Fields("estadop")
fechavp = "" & tbasele.Fields("fechavp")
'abono = "" & tbasele.Fields("abono")
'saldo = "" & tbasele.Fields("saldo")
estadopp = "" & tbasele.Fields("estadop")
estador = "" & tbasele.Fields("estador")
ochodia = Format(CVDate(fechaef) + 8, "dd/mm/yyyy")
found = busca_codigo("" & aceptante, 0)
found = busca_codigo("" & girador, 1)
found = busca_codigo("" & negociado, 3)

found = busca_banco()
found = busca_seccion()
suma_letra
Frame4.Visible = True
letran.SetFocus
Exit Sub
cmd6_err:
MsgBox "Elija un Dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub dkjbo7232_Click()
Dim found As Integer
On Error GoTo cmd5623_err
If Frame3.Visible = True Then Exit Sub
If Frame4.Visible = True Then Exit Sub
If MsgBox("Desea Borrar la letra:" & tbasele.Fields("letra"), 1, "Desea") <> 1 Then Exit Sub
tbasele.Delete

Exit Sub
cmd5623_err:
Exit Sub
End Sub



Private Sub estadop_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fechavp.SetFocus
End Sub

Private Sub fechavp_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub Form_Activate()
Dim mytablex As New adodb.Recordset
banco.Clear
banco.AddItem "%"

mytablex.Open "select * from banco ", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   Do
   If mytablex.EOF Then Exit Do
      banco.AddItem "" & mytablex.Fields("banco") & "|" & "" & mytablex.Fields("descripcio")
   mytablex.MoveNext
   Loop
End If
mytablex.Close
banco.ListIndex = 0

local1.Clear
local1.AddItem "%"

mytablex.Open "select * from tlocal ", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   Do
   If mytablex.EOF Then Exit Do
      local1.AddItem "" & mytablex.Fields("codigo") & "|" & "" & mytablex.Fields("nombre")
   mytablex.MoveNext
   Loop
End If
mytablex.Close
local1.ListIndex = 0



seccion.Clear
seccion.AddItem "%"

mytablex.Open "select * from carsec ", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   Do
   If mytablex.EOF Then Exit Do
      seccion.AddItem "" & mytablex.Fields("carsec") & "|" & "" & mytablex.Fields("descripcio")
   mytablex.MoveNext
   Loop
End If
mytablex.Close
seccion.ListIndex = 0


If acu = "V" Then
   xnameclie = "clientes"
   'xcuentaco = "letrav"
End If
If acu = "C" Then
   xnameclie = "proveedo"
   'xcuentaco = "letrac"
End If
fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
fechaf = Format(Now, "dd/mm/yyyy")
sql_letras

End Sub


Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "%"
'Combo1.AddItem "Pendiente"
'Combo1.AddItem "Cancelado"
Combo1.AddItem "Protestado"
Combo1.AddItem "Renovado"

Combo1.ListIndex = 0
condicion.Clear
condicion.AddItem "%"
condicion.AddItem "SALDO>0"
condicion.AddItem "SALDO>=0"
condicion.AddItem "SALDO=0"
condicion.AddItem "SALDO<0"
condicion.AddItem "SALDO<=0"
condicion.ListIndex = 0

End Sub

Private Sub lo232_Click()
If Frame3.Visible = True Then
   Frame3.Visible = False
   Exit Sub
End If
If Frame4.Visible = True Then
   Frame4.Visible = False
   Exit Sub
End If


texpletr.Hide
Unload texpletr
End Sub
Sub sql_letras()
On Error GoTo cmd37_err
Dim buf As String

If Not IsDate(fechai) Then Exit Sub
If Not IsDate(fechaf) Then Exit Sub

buf = "select * from " & xcuentaco & " where "
buf = buf & "  fechai>='" & Format(fechai, "YYYYMMDD") & "'"
buf = buf & " and fechai<='" & Format(fechaf, "YYYYMMDD") & "' "

If local1 <> "%" Then
buf = buf & " and local like '" & extra_loquesea(local1) & "'"
End If

If banco <> "%" Then
buf = buf & " and banco like '" & extra_loquesea(banco) & "'"
End If
If seccion <> "%" Then
buf = buf & " and seccion like '" & extra_loquesea(seccion) & "'"
End If

If codigo <> "%" Then
buf = buf & " and aceptante like '" & codigo & "'"
End If
If nombre <> "%" Then
buf = buf & " and nombrea like '" & nombre & "'"
End If

If Combo1 = "Pendiente" Then
buf = buf & " and estado='0'"
End If
If Combo1 = "Cancelado" Then
buf = buf & " and estado='1'"
End If
If Combo1 = "Protestado" Then
buf = buf & " and estadop='1'"
End If
If Combo1 = "Renovado" Then
buf = buf & " and estador='1'"
End If
If condicion = "SALDO>0" Then
buf = buf & " and saldo>0"
End If
If condicion = "SALDO>=0" Then
buf = buf & " and saldo>=0"
End If
If condicion = "SALDO=0" Then
buf = buf & " and saldo=0"
End If
If condicion = "SALDO<0" Then
buf = buf & " and saldo<0"
End If
If condicion = "SALDO<=0" Then
buf = buf & " and saldo<=0"
End If

buf = buf & " order by fechai,letra"

If tbasele.State = 1 Then tbasele.Close
tbasele.Open buf, cn, adOpenStatic, adLockOptimistic
Set dbgrid2.DataSource = tbasele
'suma_letras
               
Exit Sub
cmd37_err:
MsgBox "Error en select " & error$, 48, "Aviso"
Exit Sub

End Sub
Sub suma_letras()
Dim sdx As Double
Dim sdx1 As Double
sdx = 0
sdx1 = 0

Do
If tbasele.EOF Then Exit Do
If "" & tbasele.Fields("moneda") = "S" Then
   sdx = sdx + Val("" & tbasele.Fields("saldo"))
End If
If "" & tbasele.Fields("moneda") = "D" Then
   sdx1 = sdx1 + Val("" & tbasele.Fields("saldo"))
End If
tbasele.MoveNext
Loop
soles = Format(sdx, "0.00")
dolares = Format(sdx1, "0.00")
End Sub

Private Sub moifd343_Click()
On Error GoTo cmd89_err
If Frame3.Visible = True Then Exit Sub
If Frame4.Visible = True Then Exit Sub
tletra.bandera = "MODIFICA"
tletra.codigo.Enabled = False
tletra.codigo = "" & tbasele.Fields("letra")
tletra.acu = acu
tletra.Show 1
Exit Sub
cmd89_err:
Exit Sub
End Sub

Private Sub nier33_Click()
If Frame3.Visible = True Then Exit Sub
If Frame4.Visible = True Then Exit Sub
tletra.codigo.Enabled = True
tletra.bandera = "NUEVO"
tletra.acu = acu
tletra.Show 1
End Sub

Sub suma_letra()
Dim sdx As Double
sdx = Val(saldoa) - Val(abono)
importe = Format(sdx, "0.00")
End Sub
Function grabar_letra()
'tbasele.Delete
'tbasele.AddNew
tbasele.Fields("letra") = letran
tbasele.Fields("letraant") = letrar
tbasele.Fields("importe") = Format(saldoa, "0.00")
tbasele.Fields("abono") = Val(abono)
tbasele.Fields("saldo") = Format(importe, "0.00")
tbasele.Fields("nrounico") = nrounico
tbasele.Fields("ochodia") = ochodia
tbasele.Fields("negociado") = negociado
tbasele.Fields("fechai") = Format(fechaei, "dd/mm/yyyy")
tbasele.Fields("fechaf") = Format(fechaef, "dd/mm/yyyy")
tbasele.Fields("moneda") = moneda
tbasele.Fields("paridad") = Val(paridad)
tbasele.Fields("aceptante") = Trim(aceptante)
tbasele.Fields("girador") = Trim(girador)
tbasele.Fields("banco") = bancor
tbasele.Fields("estador") = "1"
tbasele.Fields("estadop") = "0"
tbasele.Fields("estado") = "0"
tbasele.Fields("seccion") = seccionr
tbasele.Fields("agencia") = agencia
tbasele.Fields("observa") = observa
tbasele.Fields("nombreg") = nombreg
tbasele.Fields("nombrea") = nombrea
tbasele.Update
End Function
Function valida_letrar(buf As String)
Dim sw As Integer

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable(xcuentaco)
mytablex.Index = "letra"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   valida_letrar = 1
End If
mytablex.Close
 
End Function
Function busca_codigo(buf As String, sw As Integer)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable(xnameclie)
mytablex.Index = "codigo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_codigo = 1
   If sw = 0 Then
   nombreg = "" & mytablex.Fields("nombre")
   End If
   If sw = 1 Then
   nombrea = "" & mytablex.Fields("nombre")
   End If
   If sw = 3 Then
   negociadon = "" & mytablex.Fields("nombre")
   End If
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Function busca_banco()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("banco")
mytablex.Index = "banco"
mytablex.Seek "=", bancor
If Not mytablex.NoMatch Then
   nbanco = "" & mytablex.Fields("descripcio")
   busca_banco = 1
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Function busca_seccion()

Dim mytablex As Table
nseccion = ""

Set mytablex = mydbxglo.OpenTable("carsec")
mytablex.Index = "carsec"
mytablex.Seek "=", seccionr
If Not mytablex.NoMatch Then
   nseccion = "" & mytablex.Fields("descripcio")
   busca_seccion = 1
End If
'------------------------------------- ------------
mytablex.Close
 

End Function



