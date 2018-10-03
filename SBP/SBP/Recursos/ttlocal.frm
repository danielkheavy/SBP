VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ttlocal 
   BackColor       =   &H00808080&
   Caption         =   "Tabla de Locales"
   ClientHeight    =   9780
   ClientLeft      =   105
   ClientTop       =   -105
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   11280
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Height          =   9495
      Left            =   0
      TabIndex        =   11
      Top             =   30
      Visible         =   0   'False
      Width           =   12255
      Begin VB.Frame FraFacturacionElectronica 
         Caption         =   "Facturacion Electronica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   0
         TabIndex        =   56
         Top             =   8400
         Width           =   11055
         Begin VB.TextBox CodSede 
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
            Left            =   8880
            MaxLength       =   4
            TabIndex        =   64
            Top             =   240
            Width           =   1935
         End
         Begin VB.ComboBox cmb_ESunat 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5400
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox Cmb_timpresion 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8880
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   720
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ComboBox cmb_toperacion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "ttlocal.frx":0000
            Left            =   1680
            List            =   "ttlocal.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label LblCodigoSede 
            Caption         =   "Codigo sede Portal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   7200
            TabIndex        =   63
            Top             =   360
            Width           =   1620
         End
         Begin VB.Label lblEnvioSunat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Envio Sunat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4200
            TabIndex        =   61
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label lblTipoImpresion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Impresion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7560
            TabIndex        =   59
            Top             =   720
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lblTipoOPeracion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Operacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   1320
         End
      End
      Begin VB.TextBox cabecera 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   2280
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   54
         Top             =   6720
         Width           =   7215
      End
      Begin VB.TextBox basedatostransmision 
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
         MaxLength       =   30
         TabIndex        =   52
         Top             =   6120
         Width           =   3495
      End
      Begin VB.TextBox basedatos 
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
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   51
         Top             =   6120
         Width           =   3495
      End
      Begin VB.TextBox claverecepcion 
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
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   49
         Top             =   5760
         Width           =   3495
      End
      Begin VB.TextBox usuariorecepcion 
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
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   48
         Top             =   5400
         Width           =   3495
      End
      Begin VB.TextBox iprecepcion 
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
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   47
         Top             =   5040
         Width           =   3495
      End
      Begin VB.TextBox clavetransmision 
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
         MaxLength       =   30
         TabIndex        =   44
         Top             =   5760
         Width           =   3495
      End
      Begin VB.TextBox usuariotransmision 
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
         MaxLength       =   30
         TabIndex        =   42
         Top             =   5400
         Width           =   3495
      End
      Begin VB.TextBox iptransmision 
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
         MaxLength       =   30
         TabIndex        =   40
         Top             =   5040
         Width           =   3495
      End
      Begin VB.TextBox defecto 
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
         MaxLength       =   1
         TabIndex        =   38
         Top             =   4200
         Width           =   255
      End
      Begin VB.TextBox codigo 
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
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox codigo1 
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
         MaxLength       =   11
         TabIndex        =   25
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox nombre 
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
         Left            =   2280
         MaxLength       =   60
         TabIndex        =   24
         Top             =   960
         Width           =   5775
      End
      Begin VB.TextBox nombrec 
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
         TabIndex        =   23
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox contacto 
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
         TabIndex        =   22
         Top             =   1680
         Width           =   5775
      End
      Begin VB.TextBox direccion 
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
         TabIndex        =   21
         Top             =   2040
         Width           =   5775
      End
      Begin VB.TextBox dpto 
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
         TabIndex        =   20
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox distrito 
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
         TabIndex        =   19
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox telefono 
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
         TabIndex        =   18
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox telefono1 
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
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   17
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox telefono2 
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
         Left            =   6120
         MaxLength       =   15
         TabIndex        =   16
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox correo 
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
         TabIndex        =   15
         Top             =   3840
         Width           =   5775
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   8400
         Picture         =   "ttlocal.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Imprimir todo"
         Top             =   3120
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   8400
         Picture         =   "ttlocal.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2160
         Width           =   1470
      End
      Begin VB.TextBox zona 
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
         TabIndex        =   12
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cabecera Ticket"
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
         TabIndex        =   55
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Basedatos"
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
         TabIndex        =   53
         Top             =   6120
         Width           =   2175
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recepcion"
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
         Left            =   6000
         TabIndex        =   50
         Top             =   4680
         Width           =   3495
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transmision"
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
         Left            =   2280
         TabIndex        =   46
         Top             =   4680
         Width           =   3495
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave"
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
         TabIndex        =   45
         Top             =   5760
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario"
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
         TabIndex        =   43
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ip Remoto "
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
         TabIndex        =   41
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Por Defecto (S)"
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
         TabIndex        =   39
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
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
         TabIndex        =   37
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Alterno"
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
         TabIndex        =   36
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Razon Social"
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
         TabIndex        =   35
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre Comercial"
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
         TabIndex        =   34
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contacto"
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
         TabIndex        =   33
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
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
         TabIndex        =   32
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Departamento"
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
         TabIndex        =   31
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distrito"
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
         TabIndex        =   30
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefonos"
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
         TabIndex        =   29
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Correo Electronico"
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
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Zona"
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
         TabIndex        =   27
         Top             =   3120
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
         Picture         =   "ttlocal.frx":1198
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
         Picture         =   "ttlocal.frx":23AA
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
         Picture         =   "ttlocal.frx":35BC
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
         Picture         =   "ttlocal.frx":47CE
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
         Picture         =   "ttlocal.frx":59E0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
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
         ColumnCount     =   2
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   5940.284
            EndProperty
            BeginProperty Column01 
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
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "ttlocal"
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
    codigo.Enabled = True
    codigo = ""
    codigo.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    buf = txempre.Fields("codigo")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txempre.Fields("codigo"), 1, "Aviso") <> 1 Then
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

Private Sub cmdAddEntry_Click()
    ajdu1_Click

End Sub

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

    ' 01/08/2018 Testing Facturacion Electronica

    ' 01/08/2018 Testing Facturacion Electronica

    'aqui se detiene el servicio web
    'Shell ("net stop C:\BYH\SERVICIO\VISITEC\facturador-local.exe")
    found = grabar()
    'inicio 29/01/2018 pll
    'Call lee_conf_BYH("C:\BYH\SERVICIO\VISITEC\application.yml", Mid(cmb_ESunat.Text, 1, 1))
    'fin 29/01/2018 pll
    'aqui reiniciamos el servicio web
    'Shell ("net start C:\BYH\SERVICIO\VISITEC\facturador-local.exe")

    ' Testing Proyecto Facturacion Electronica 05/04/2018
    'Shell ("net stop C:\BYH\SERVICIO\VISITEC\facturador-local.exe")
    'found = grabar()
    'Call lee_conf_BYH("C:\BYH\SERVICIO\VISITEC\application.yml", Mid(cmb_ESunat.Text, 1, 1))
    'Shell ("net start C:\BYH\SERVICIO\VISITEC\facturador-local.exe")
    ' Testing Proyecto Facturacion Electronica 05/04/2018

End Sub

Private Sub cmdPrint_Click()
    djuer1_Click

End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(codigo) = 0 Then Exit Sub
    codigo1.SetFocus

End Sub

Private Sub codigo1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    nombre.SetFocus

End Sub

Private Sub codigo1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If codigo.Enabled = True Then
            codigo.SetFocus

        End If

        Exit Sub

    End If

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
            cad = "SELECT * from tlocal    "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT *  from tlocal   where  " & Combo1 & " like '" & buffer & "%'"

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

Private Sub contacto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    direccion.SetFocus

End Sub

Private Sub contacto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        nombrec.SetFocus
        Exit Sub

    End If

End Sub

Private Sub correo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub

End Sub

Private Sub correo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        telefono2.SetFocus
        Exit Sub

    End If

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'codigo = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'codigo.SetFocus
        'codigo_KeyPress 13
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

Private Sub direccion_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    dpto.SetFocus

End Sub

Private Sub direccion_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        contacto.SetFocus
        Exit Sub

    End If

End Sub

Private Sub distrito_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    zona.SetFocus

End Sub

Private Sub distrito_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        dpto.SetFocus
        Exit Sub

    End If

End Sub

Private Sub djuer1_Click()

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "tlocal"
    reporgen.Show 1

End Sub

Private Sub dlo132_Click()

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    ttlocal.Hide
    Unload ttlocal

End Sub

Private Sub dpto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    distrito.SetFocus

End Sub

Private Sub dpto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        direccion.SetFocus
        Exit Sub

    End If

End Sub

Private Sub estado_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        correo.SetFocus
        Exit Sub

    End If

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = txempre.Fields("codigo")

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
    codigo.Enabled = False
    nombre.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = txempre.Fields("codigo")

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
    codigo.Enabled = False
    nombre.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    Command1_Click

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "NOMBRE"
    Combo1.AddItem "CODIGO"
    Combo1.ListIndex = 0

End Sub

Sub inicializa()
    cabecera = ""
    basedatostransmision = ""
    basedatos = ""
    iptransmision = ""
    usuariotransmision = ""
    clavetransmision = ""

    iprecepcion = ""
    usuariorecepcion = ""
    claverecepcion = ""

    defecto = ""
    codigo1 = ""
    nombre = ""
    nombrec = ""
    contacto = ""
    direccion = ""
    dpto = ""
    distrito = ""
    telefono = ""
    telefono1 = ""
    telefono2 = ""
    correo = ""

    ' Varios Locales FE 18/05/2018
    CodSede = ""
    ' Varios Locales FE 18/05/2018

    ' 01/08/2018 Testing Facturacion Electronica
    cmb_toperacion.Clear
    cmb_toperacion.AddItem "Exonerado"
    cmb_toperacion.AddItem "Inafecto"
    cmb_toperacion.ListIndex = 0

    Cmb_timpresion.Clear
    Cmb_timpresion.AddItem "Generico"
    Cmb_timpresion.AddItem "Codigo QR"
    Cmb_timpresion.ListIndex = 0

    cmb_ESunat.Clear
    cmb_ESunat.AddItem "Automatico"
    cmb_ESunat.AddItem "Manual"
    cmb_ESunat.ListIndex = 0
    ' 01/08/2018 Testing Facturacion Electronica

End Sub

Sub pone_registro()
    basedatostransmision = Trim("" & txempre.Fields("basedatostransmision"))
    basedatos = Trim("" & txempre.Fields("basedatos"))
    iptransmision = Trim("" & txempre.Fields("iptransmision"))
    usuariotransmision = Trim("" & txempre.Fields("usuariotransmision"))
    clavetransmision = Trim("" & txempre.Fields("clavetransmision"))

    iprecepcion = Trim("" & txempre.Fields("iprecepcion"))
    usuariorecepcion = Trim("" & txempre.Fields("usuariorecepcion"))
    claverecepcion = Trim("" & txempre.Fields("claverecepcion"))

    'defecto = Trim("" & txempre.Fields("defecto"))
    cabecera = Trim("" & txempre.Fields("cabecera"))
    codigo = Trim("" & txempre.Fields("codigo"))
    codigo1 = Trim("" & txempre.Fields("codigo1"))

    ' Varios Locales FE 18/05/2018
    CodSede = Trim("" & txempre.Fields("codsede"))
    ' Varios Locales FE 18/05/2018

    nombre = Trim("" & txempre.Fields("nombre"))
    nombrec = Trim("" & txempre.Fields("nombrec"))
    contacto = Trim("" & txempre.Fields("contacto"))
    direccion = Trim("" & txempre.Fields("direccion"))
    dpto = Trim("" & txempre.Fields("dpto"))
    distrito = Trim("" & txempre.Fields("distrito"))
    zona = Trim("" & txempre.Fields("zona"))
    telefono = Trim("" & txempre.Fields("telefono"))
    telefono1 = Trim("" & txempre.Fields("telefono1"))
    telefono2 = Trim("" & txempre.Fields("telefono2"))
    correo = Trim("" & txempre.Fields("correo"))

    'inicio 14/02/2018 pll
    If Trim("" & txempre.Fields("esunat")) = "M" Then
        cmb_ESunat.ListIndex = 1
    ElseIf Trim("" & txempre.Fields("esunat")) = "A" Then
        cmb_ESunat.ListIndex = 0

    End If

    'fin 14/02/2018 pll

    ' 01/08/2018 Testing Facturacion Electronica
    cmb_toperacion.ListIndex = 0

    If "" & txempre.Fields("toperacion") = "I" Then
        cmb_toperacion.ListIndex = 1

    End If

    ' 01/08/2018 Testing Facturacion Electronica

End Sub

Sub grabando()
    txempre.Fields("cabecera") = Trim(cabecera)
    txempre.Fields("basedatostransmision") = Trim(basedatostransmision)
    txempre.Fields("basedatos") = Trim(basedatos)
    txempre.Fields("iptransmision") = Trim(iptransmision)
    txempre.Fields("usuariotransmision") = Trim(usuariotransmision)
    txempre.Fields("clavetransmision") = Trim(clavetransmision)

    txempre.Fields("iprecepcion") = Trim(iprecepcion)
    txempre.Fields("usuariorecepcion") = Trim(usuariorecepcion)
    txempre.Fields("claverecepcion") = Trim(claverecepcion)

    'txempre.Fields("defecto") = Trim(defecto)
    txempre.Fields("codigo1") = Trim(codigo1)

    ' Varios Locales FE 18/05/2018)
    txempre.Fields("codsede") = Trim(CodSede)
    ' Varios Locales FE 18/05/2018)

    txempre.Fields("nombre") = Trim(nombre)
    txempre.Fields("nombrec") = Trim(nombrec)
    txempre.Fields("contacto") = Trim(contacto)
    txempre.Fields("direccion") = Trim(direccion)
    txempre.Fields("dpto") = Trim(dpto)
    txempre.Fields("distrito") = Trim(distrito)
    txempre.Fields("zona") = Trim(zona)
    txempre.Fields("telefono") = Trim(telefono)
    txempre.Fields("telefono1") = Trim(telefono1)
    txempre.Fields("telefono2") = Trim(telefono2)
    txempre.Fields("correo") = Trim(correo)
    'inicio 09/02/2018 pll para la facturacion electronica
    txempre.Fields("toperacion") = Mid(cmb_toperacion.Text, 1, 1)
    txempre.Fields("timpresion") = Mid(Cmb_timpresion.Text, 1, 1)
    txempre.Fields("ESUNAT") = Mid(cmb_ESunat.Text, 1, 1)

    'fin 09/02/2018 pll
End Sub

Private Sub grba1_Click()

End Sub

Private Sub Label1_Click()

    'cmdSort_Click
End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    nombrec.SetFocus

End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        codigo1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub nombrec_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    contacto.SetFocus

End Sub

Private Sub nombrec_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        nombre.SetFocus
        Exit Sub

    End If

End Sub

Private Sub telefono_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    telefono1.SetFocus

End Sub

Private Sub telefono_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        zona.SetFocus
        Exit Sub

    End If

End Sub

Private Sub telefono1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    telefono2.SetFocus

End Sub

Private Sub telefono1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        telefono.SetFocus
        Exit Sub

    End If

End Sub

Private Sub telefono2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    correo.SetFocus

End Sub

Private Sub telefono2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        telefono1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub zona_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    telefono.SetFocus

End Sub

Private Sub zona_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        distrito.SetFocus
        Exit Sub

    End If

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
        If Len(codigo) = 0 Then
            codigo.SetFocus
            Exit Function

        End If

        rbusca.Open "select codigo from tlocal where codigo='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe codigo ", 48, "Aviso"
            Exit Function

        End If

        txempre.AddNew
        txempre.Fields("codigo") = codigo
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        txempre.Fields("codigo") = codigo
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    Dim mytablex As New ADODB.Recordset

    'mytablex.Open "select * from tlocal where codigo='" & codigo & "' and defecto='S'", cn, adOpenStatic, adLockOptimistic
    'If mytablex.RecordCount > 0 Then
    '   If "" & mytablex.Fields("codigo") <> codigo Then
    '      MsgBox "Ya existe validado en Local=" & "" & mytablex.Fields("codigo"), 48, "Aviso"
    '      mytablex.Close
    '      Exit Function
    '   End If
    'End If
    'mytablex.Close
    If Len(codigo) = 0 Then
        codigo.SetFocus
        Exit Function

    End If

    If Len(nombre) = 0 Then
        nombre.SetFocus
        Exit Function

    End If

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

