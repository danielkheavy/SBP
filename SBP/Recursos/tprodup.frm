VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tprodup 
   BackColor       =   &H00FFFF00&
   Caption         =   "Tabla de productos"
   ClientHeight    =   7635
   ClientLeft      =   165
   ClientTop       =   420
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   12285
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
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
      Height          =   7455
      Left            =   0
      TabIndex        =   199
      Top             =   0
      Visible         =   0   'False
      Width           =   12255
      Begin VB.TextBox buffer1 
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
         TabIndex        =   201
         TabStop         =   0   'False
         Top             =   6960
         Width           =   2895
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   200
         TabStop         =   0   'False
         Top             =   6960
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "tprodup.frx":0000
         Height          =   6375
         Left            =   120
         OleObjectBlob   =   "tprodup.frx":0014
         TabIndex        =   202
         TabStop         =   0   'False
         Top             =   600
         Width           =   12015
      End
      Begin VB.Label buffer 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         TabIndex        =   203
         Top             =   240
         Width           =   75
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   0
      TabIndex        =   194
      Top             =   0
      Visible         =   0   'False
      Width           =   9375
      Begin VB.CommandButton Command11 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   5520
         TabIndex        =   197
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   196
         TabStop         =   0   'False
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox xbusca1 
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   195
         TabStop         =   0   'False
         Top             =   360
         Width           =   2895
      End
      Begin MSDBGrid.DBGrid DBGrid5 
         Bindings        =   "tprodup.frx":09DF
         Height          =   3855
         Left            =   240
         OleObjectBlob   =   "tprodup.frx":09F3
         TabIndex        =   198
         Top             =   840
         Width           =   8895
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   181
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton Command6 
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
         Left            =   9840
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tprodup.frx":13C6
         Style           =   1  'Graphical
         TabIndex        =   191
         ToolTipText     =   "Salir"
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Frame Frame9 
         Caption         =   "Poner codigo Proveedor"
         Height          =   2415
         Left            =   3480
         TabIndex        =   182
         Top             =   1560
         Visible         =   0   'False
         Width           =   5295
         Begin VB.TextBox rcodigo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   187
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command9 
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
            Height          =   855
            Left            =   4200
            MaskColor       =   &H00E0E0E0&
            Picture         =   "tprodup.frx":25D8
            Style           =   1  'Graphical
            TabIndex        =   186
            ToolTipText     =   "Salir"
            Top             =   240
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
            Height          =   855
            Left            =   4200
            MaskColor       =   &H00E0E0E0&
            Picture         =   "tprodup.frx":37EA
            Style           =   1  'Graphical
            TabIndex        =   185
            ToolTipText     =   "Grabar registro"
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox costo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   184
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox fechauc 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   183
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label52 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Codigo Proveedor"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   190
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label50 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Costo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   189
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label53 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FechaUl.Compra"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   188
            Top             =   1200
            Width           =   1935
         End
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Bindings        =   "tprodup.frx":49FC
         Height          =   3255
         Left            =   240
         OleObjectBlob   =   "tprodup.frx":4A10
         TabIndex        =   192
         Top             =   240
         Width           =   11175
      End
      Begin VB.Label Label51 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nota. F7.CodigoProveedor Esc.Salir Del.Borra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   193
         Top             =   3840
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   0
      TabIndex        =   174
      Top             =   0
      Visible         =   0   'False
      Width           =   10815
      Begin VB.TextBox barras2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   178
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Borra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9000
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tprodup.frx":53E3
         Style           =   1  'Graphical
         TabIndex        =   177
         ToolTipText     =   "Borrar registro"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Graba"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9000
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tprodup.frx":65F5
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Grabar registro"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cierra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9000
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tprodup.frx":7807
         Style           =   1  'Graphical
         TabIndex        =   175
         ToolTipText     =   "Salir"
         Top             =   1920
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "tprodup.frx":8A19
         Height          =   5655
         Left            =   0
         OleObjectBlob   =   "tprodup.frx":8A2D
         TabIndex        =   179
         Top             =   720
         Width           =   7815
      End
      Begin VB.Label Label59 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CodigoAdicionar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   180
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   173
      Top             =   6600
      Width           =   3375
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Lotes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   172
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Series"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   171
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Traslado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   170
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Recal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   169
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ajustes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6480
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   168
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   735
   End
   Begin VB.ComboBox local2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   166
      Top             =   2760
      Width           =   855
   End
   Begin VB.ComboBox monedav 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   149
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox unidad1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   148
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox factor1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   147
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox pventa1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   146
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox margen1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   145
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox fechai11 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   144
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox fechaf11 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   143
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox fechaid 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   142
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox fechafd 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   141
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox margen11 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   140
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox pventa11 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   139
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox maximo11 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      MaxLength       =   5
      TabIndex        =   138
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox minimo11 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      MaxLength       =   5
      TabIndex        =   137
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox dscto 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   136
      Top             =   6840
      Width           =   735
   End
   Begin VB.TextBox margen2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   135
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox pventa2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   134
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox factor2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   133
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox unidad2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   132
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox margen3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   131
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox pventa3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   130
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox factor3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   129
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox unidad3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   128
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox margen4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   127
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox pventa4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   126
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox factor4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   125
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox unidad4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   124
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox margen5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   123
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox pventa5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   122
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox factor5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   121
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox unidad5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   120
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox margen6 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   119
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox pventa6 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   118
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox factor6 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   117
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox unidad6 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   116
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox margen7 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   115
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox pventa7 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   114
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox factor7 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   113
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox unidad7 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   112
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox margen8 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   111
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox pventa8 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   110
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox factor8 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   109
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox unidad8 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   108
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox margen9 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   107
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox pventa9 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   106
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox factor9 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   105
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox unidad9 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   104
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox margen10 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   103
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox pventa10 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      MaxLength       =   10
      TabIndex        =   102
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox factor10 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   101
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox unidad10 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      MaxLength       =   6
      TabIndex        =   100
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox minimo12 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      MaxLength       =   5
      TabIndex        =   99
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox maximo12 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      MaxLength       =   5
      TabIndex        =   98
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox pventa12 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   97
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox margen12 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   96
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox minimo13 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      MaxLength       =   5
      TabIndex        =   95
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox maximo13 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      MaxLength       =   5
      TabIndex        =   94
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox pventa13 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   93
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox margen13 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   92
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox minimo14 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      MaxLength       =   5
      TabIndex        =   91
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox maximo14 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      MaxLength       =   5
      TabIndex        =   90
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox pventa14 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   89
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox margen14 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   88
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox minimo15 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      MaxLength       =   5
      TabIndex        =   87
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox maximo15 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      MaxLength       =   5
      TabIndex        =   86
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox pventa15 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   85
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox margen15 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   84
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox ccosto 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      MaxLength       =   6
      TabIndex        =   83
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox local1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   80
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox flete 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      MaxLength       =   10
      TabIndex        =   76
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox percepcion 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   74
      Top             =   3480
      Width           =   855
   End
   Begin VB.ComboBox estado 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   72
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Precios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "Ayuda"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12600
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
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
      Width           =   1335
   End
   Begin VB.ComboBox xproveedor 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   13680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1140
   End
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
      Left            =   9360
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tprodup.frx":9400
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Ayuda"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grafico"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "Ayuda"
      Top             =   0
      Width           =   735
   End
   Begin VB.CheckBox insumo 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ingresar solo Insumos"
      Height          =   375
      Left            =   7680
      TabIndex        =   67
      Top             =   8760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox fabrica 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaxLength       =   11
      TabIndex        =   66
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   13680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   13680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   13680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox fechavence 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      MaxLength       =   10
      TabIndex        =   62
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox costop 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   61
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox costou 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   59
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox factor 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   57
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox unidad 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   55
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox monedac 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   2760
      Width           =   615
   End
   Begin VB.ComboBox vecaja 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10680
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   1920
      Width           =   495
   End
   Begin VB.ComboBox oferta 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10680
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   1560
      Width           =   495
   End
   Begin VB.ComboBox servicio 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9360
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox serie 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox vtaund 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9360
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   1920
      Width           =   615
   End
   Begin VB.ComboBox peso 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox comision 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      MaxLength       =   10
      TabIndex        =   39
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox pesokgr 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      MaxLength       =   10
      TabIndex        =   37
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox isc 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   35
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox igv 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   33
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox color 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   29
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox lineatalla 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   27
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox categoria 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   25
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox marca 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   23
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox seccion 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   21
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox subfamilia 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      MaxLength       =   6
      TabIndex        =   19
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox familia 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      MaxLength       =   6
      TabIndex        =   17
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox presenta 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   15
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox descorto 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   22
      TabIndex        =   13
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox barras 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox descripcio 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
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
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tprodup.frx":A612
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Grabar registro"
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
      Picture         =   "tprodup.frx":B824
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Nuevo registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Kardex"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ayuda"
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
      Picture         =   "tprodup.frx":CA36
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Left            =   8640
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tprodup.frx":DC48
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
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
      Left            =   720
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tprodup.frx":EE5A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tprodup.frx":1006C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "tprodup.frx":1127E
      Height          =   1695
      Left            =   120
      OleObjectBlob   =   "tprodup.frx":11292
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label Label33 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ListaNro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   167
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label36 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mon.Pvta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   165
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label37 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Und"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   164
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label38 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   163
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label39 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pvta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   162
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label40 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%Margen"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   161
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label41 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Min"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   160
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label42 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Max"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   159
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label43 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pvta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   158
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label44 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%Margen"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   157
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label45 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   156
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label46 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaFinal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   155
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label47 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   154
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label48 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaFinal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   153
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label49 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dscto Autom."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   152
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label Label54 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Oferta.precio=0 acepta"
      Height          =   195
      Left            =   9240
      TabIndex        =   151
      Top             =   7200
      Width           =   1635
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ccosto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   150
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Foto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   82
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label56 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   81
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label55 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F.Venc."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   79
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label35 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoProm."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   78
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label31 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Flete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   77
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Percep."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   75
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label32 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Activo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   73
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label fotonombre 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   64
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image foto 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label paridad 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   63
      Top             =   120
      Width           =   120
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoUlt."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   60
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   58
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unidad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   56
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mon.Costo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   54
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VeCaja"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   52
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Oferta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   50
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servicio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   48
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   46
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VtaUnidad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   44
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Peso"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   42
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comis."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PesoKgr"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Isc"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   36
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Igv"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   34
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7440
      TabIndex        =   32
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fabricante"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7440
      TabIndex        =   31
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   30
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LineaTalla"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   28
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Categoria"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   26
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marca"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   24
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seccion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   22
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SubFamilia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Presentac."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descr.Corto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cod.Barras"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu rect398912 
      Caption         =   "Re&Cetas"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tprodup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ArrBarCode(43) As String

Private Sub ajdu1_Click()
Dim found As Integer
If Frame4.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
inicializa
codigo = ""
found = busca_parame(0)
codigo.SetFocus

End Sub


Private Sub barras_Change()
hacer_barras
End Sub

Private Sub barras_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
descripcio.SetFocus
End Sub

Private Sub barras_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   Label2_Click
End If
End Sub


Private Sub barras2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
Command3_Click
End Sub

Private Sub bo712_Click()
Dim found As Integer
If Frame4.Visible = True Then Exit Sub
If Frame2.Visible = True Then
   borrar_barras
   Label2_Click
   Exit Sub
End If
If Frame1.Visible = True Then Exit Sub
If insumo.Value = 1 Then
If Mid$(codigo, 1, 1) <> "I" Then
   MsgBox "Codigo debe empezar con I", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If
End If
If insumo.Value = 0 Then
If Mid$(codigo, 1, 1) = "I" Then
   MsgBox "Codigo No debe empezar con I,por ser usado para insumo", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If
End If

found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
found = borra_registro()
If found = 0 Then Exit Sub
MsgBox "Ok,Registro Borrado", 48, "Aviso"
codigo = ""
inicializa
codigo.SetFocus
End Sub

Private Sub buffer1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 And Combo3 = "Proveedor" Then 'f1
   consulta_proveedor
End If
End Sub

Private Sub categoria_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
lineatalla.SetFocus

End Sub

Private Sub categoria_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   marca.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_categoria
End If
If KeyCode = &H76 Then  'f7
   tcategor.Show 1
End If


End Sub

Private Sub ccosto_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
End Sub

Private Sub ccosto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   vecaja.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_ccosto
End If

End Sub

Private Sub cmdAddEntry_Click()
ajdu1_Click
End Sub

Private Sub cmdDelete_Click()
bo712_Click
End Sub

Private Sub cmdExit_Click()
dlo132_Click

End Sub

Private Sub cmdHelp_Click()
opcion2 = "1"
repinv.Label15.Visible = True
repinv.Label16.Visible = True
repinv.fechai.Visible = True
repinv.fechaf.Visible = True
repinv.producto = codigo
repinv.Show 1
End Sub

Private Sub cmdPrint_Click()
djuer1_Click
End Sub


Private Sub cmdSave_Click()
grba1_Click
End Sub

Private Sub cmdSort_Click()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Producto"
Combo3.AddItem "Descripcio"
Combo3.AddItem "familia"
Combo3.AddItem "Subfamilia"
Combo3.AddItem "Categoria"
Combo3.AddItem "Seccion"
Combo3.AddItem "Color"
Combo3.AddItem "Proveedor"
Combo3.ListIndex = 0


opcion1 = "1"
Frame1.Visible = True
buffer = ""
found = ejecuta(0)
DBGrid1.SetFocus
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
If Len(codigo) = 0 Then Exit Sub
If insumo.Value = 1 Then
   If Mid$(codigo, 1, 1) <> "I" Then
      MsgBox "El Codigo debe empezar con I,Por ser insumo", 48, "Aviso"
      codigo = ""
      codigo.SetFocus
      Exit Sub
   End If
   
End If
If insumo.Value = 0 Then
   If Mid$(codigo, 1, 1) = "I" Then
      MsgBox "El Codigo NO debe empezar con I,Por ser insumo", 48, "Aviso"
      codigo = ""
      codigo.SetFocus
      Exit Sub
   End If
End If

found = busca_registro()
If found = 0 Then
   inicializa
End If
found = busca_parame(3)
consulta_bodega
barras.SetFocus
End Sub

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
cmdSort_Click
End If
End Sub


Private Sub color_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
fabrica.SetFocus

End Sub

Private Sub color_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   lineatalla.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_color
End If
If KeyCode = &H76 Then  'f7
   tcolor.Show 1
End If

End Sub

Private Sub comision_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
pesokgr.SetFocus

End Sub

Private Sub comision_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   vecaja.SetFocus
   Exit Sub
End If

End Sub

Function ejecuta(sw As Integer)
Dim buf As String
Dim indx As Integer
On Error GoTo cmd34_err
indx = -1
If opcion1 = "1" Then
   If Combo3 = "Proveedor" Then
     If Len(buffer) = 0 Then
        buf = "select Producto.descripcio,Producto.producto,producto.Unidad,producto.Factor,producto.costou,Producto.Costop,producto.Unidad1,Producto.Pventa1,Producto.Proveedor1 from codprov left join producto on codprov.producto=producto.producto where codprov.codigo='" & buffer1 & "'"
        If insumo.Value = 1 Then
         buf = buf & " and mid$(producto.producto,1,1)='I' "
        End If
        If insumo.Value = 0 Then
         buf = buf & " and mid$(producto.producto,1,1)<>'I' "
        End If
        Else
        buf = "select Producto.descripcio,Producto.producto,producto.Unidad,producto.Factor,producto.costou,Producto.Costop,producto.Unidad1,Producto.Pventa1,Producto.Proveedor1 from codprov left join producto on codprov.producto=producto.producto where codprov.codigo='" & buffer1 & "' and "
        If insumo.Value = 1 Then
         buf = buf & " and mid$(producto.producto,1,1)='I'  and "
         buf = buf & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
         End If
        If insumo.Value = 0 Then
         buf = buf & " and mid$(producto.producto,1,1)<>'I'  and "
         buf = buf & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
        End If
        indx = DBGrid1.Col
     End If
     GoTo ajk1
   End If
   If Len(buffer) = 0 Then
      buf = "select Descripcio,Producto,Unidad,Factor,Costou,Costop,Unidad1,Factor1,Pventa1,Proveedor1 from producto "
      If insumo.Value = 1 Then
         buf = buf & " where mid$(producto,1,1)='I' "
      End If
      If insumo.Value = 0 Then
         buf = buf & " where mid$(producto,1,1)<>'I' "
      End If
      Else
      buf = "select Descripcio,Producto,Unidad,Factor,Costou,Costop,Unidad1,Factor1,Pventa1,Proveedor1 from producto where "
      If insumo.Value = 1 Then
         buf = buf & " mid$(producto,1,1)='I'  and "
        buf = buf & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      End If
      If insumo.Value = 0 Then
         buf = buf & " mid$(producto,1,1)<>'I'  and "
         buf = buf & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      End If
      indx = DBGrid1.Col
   End If
End If
ajk1:
If opcion1 = "2" Then
   If Len(buffer) = 0 Then
      buf = "select Descripcio,Familia from familia "
      Else
      buf = "select Descripcio,Familia from familia where " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
   End If
End If
If opcion1 = "190" Then  'local
   If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from tlocal "
      Else
      buf = "select Nombre,Codigo from tlocal where " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
   End If
End If

If opcion1 = "3" Then
   If Len(buffer) = 0 Then
      buf = "select Descripcio,SubFamilia,Familia from Subfamil where familia='" & familia & "'"
      Else
      buf = "select Descripcio,SubFamilia,Familia from Subfamil where familia='" & familia & "' and " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
   End If
End If
If opcion1 = "4" Then
   If Len(buffer) = 0 Then
      buf = "select Descripcio,Seccion from seccion "
      Else
      buf = "select Descripcio,Seccion from seccion where " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
   End If
End If
If opcion1 = "5" Then
   If Len(buffer) = 0 Then
      buf = "select Descripcio,Marca from Marca "
      Else
      buf = "select Descripcio,Marca from Marca where " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
      End If
End If
If opcion1 = "6" Then
   If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from Fabrica "
      Else
      buf = "select Nombre,Codigo from Fabrica where " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
   End If
End If
If opcion1 = "7" Then
   If Len(buffer) = 0 Then
      buf = "select Descripcio,categoria from categori "
      Else
      buf = "select Descripcio,categoria from categori where " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
   End If
End If
If opcion1 = "8" Then
   If Len(buffer) = 0 Then
      buf = "select Descripcio,Linea from Linea "
      Else
      buf = "select Descripcio,Linea from Linea where " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
   End If
End If
If opcion1 = "9" Then
   If Len(buffer) = 0 Then
      buf = "select Descripcio,Color from Color "
      Else
      buf = "select Descripcio,Color from Color where " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
   End If
End If
If opcion1 = "10" Then
   If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from proveedor "
      Else
      buf = "select Nombre,Codigo from proveedor where " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
      End If
End If
If opcion1 = "11" Then
   If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from proveedor "
      Else
      buf = "select Nombre,Codigo from proveedor where " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
   End If
End If
   If opcion1 = "27" Or opcion1 = 28 Or opcion1 = 29 Or opcion1 = 30 Or opcion1 = 31 Then
   If Len(buffer) = 0 Then
      buf = "select Descripcio,Ccosto from ccosto "
      Else
      buf = "select Descripcio,Ccosto from ccosto where " & "" & Data1.Recordset.Fields("" & DBGrid1.Columns(DBGrid1.Col).Caption).name & " like '" & buffer & "*'"
      indx = DBGrid1.Col
   End If
   End If
   Data1.Connect = "foxpro 2.5;"
   Data1.DatabaseName = globaldir
   Data1.RecordSource = buf
   Data1.Refresh
           If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer = ""
                  Exit Function
               End If
   pone_tamano
   If indx <> -1 Then
      DBGrid1.Col = indx
   End If
   If sw = 1 Then
      DBGrid1.SetFocus
   End If
   ejecuta = 1
   Exit Function
cmd34_err:
'MsgBox "Error en Consulta " & error$, 48, "Aviso"
buffer = ""
Exit Function

End Function




Private Sub Command1_Click()
Dim found As Integer
If Len(codigo) = 0 Then Exit Sub
found = busca_registro()
If found = 0 Then
   codigo.SetFocus
   Exit Sub
End If
Frame2.Visible = True
Frame2.Caption = "NUMERO SERIES"
barras2 = ""
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = "select Descripcio,Producto from proserie where producto='" & codigo & "'"
               Data2.Refresh
               DBGrid2.Columns(0).Width = 3500
               DBGrid2.Columns(1).Width = 1500
               DBGrid2.SetFocus
               'barras2.SetFocus

End Sub

Private Sub Command10_Click()
Dim found As Integer
If Len(fechauc) > 0 Then
   If Len(fechauc) <> 10 Then
      fechauc = ""
      fechauc.SetFocus
      Exit Sub
   End If
   If Not IsDate(fechauc) Then
   fechauc = ""
   fechauc.SetFocus
   Exit Sub
   End If
End If
found = graba_rcodigo()
busca_selec_proveedor
carga_proveedor
Command9_Click
End Sub

Private Sub Command11_Click()
Dim buf As String
   buf = "select Nombre,Codigo from proveedo "
   If xbusca1 <> "*" Then
      buf = buf & " where  " & Combo2 & " like '" & xbusca1 & "*'"
   End If
   
   Data5.Connect = "foxpro 2.5;"
   Data5.DatabaseName = globaldir
   Data5.RecordSource = buf
   Data5.Refresh
   If Data5.Recordset.EOF = True And Data5.Recordset.BOF = True Then
      Data5.Recordset.Close
      Exit Sub
   End If
   Frame10.Visible = True
   DBGrid5.Columns(0).Width = 6000
   DBGrid5.Columns(1).Width = 2000
   DBGrid5.SetFocus

End Sub

Private Sub Command12_Click()
Dim found As Integer
If local2.Visible <> True Then Exit Sub  'si no es precios x locales
If Frame4.Visible = True Then Exit Sub
If Frame2.Visible = True Then
   borrar_barras
   Label2_Click
   Exit Sub
End If
If Frame1.Visible = True Then Exit Sub
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
tprecios.monedac = monedac
tprecios.unidad = unidad
tprecios.factor = factor
tprecios.costou = costou
tprecios.monedav = monedav
tprecios.producto = codigo
tprecios.descripcio = descripcio
tprecios.Show 1
found = busca_registro()
codigo.SetFocus
End Sub

Private Sub Command13_Click()
Dim found As Integer
If Frame4.Visible = True Then Exit Sub
If Frame2.Visible = True Then
   borrar_barras
   Label2_Click
   Exit Sub
End If
If Frame1.Visible = True Then Exit Sub
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
trecalcu.producto = codigo
'Cargastk.descripcio = descripcio
trecalcu.Show 1
consulta_bodega

End Sub

Private Sub Command14_Click()
Dim found As Integer
If Frame4.Visible = True Then Exit Sub
If Frame2.Visible = True Then
   borrar_barras
   Label2_Click
   Exit Sub
End If
If Frame1.Visible = True Then Exit Sub
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
doctrasl.producto = codigo
doctrasl.descripcio = descripcio
doctrasl.Show 1

'traslara.Show 1
consulta_bodega

End Sub

Private Sub Command15_Click()
Dim found As Integer
If Len(codigo) = 0 Then Exit Sub
found = busca_registro()
If found = 0 Then
   codigo.SetFocus
   Exit Sub
End If
Frame2.Visible = True
Frame2.Caption = "LOTES"
barras2 = ""
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = "select Descripcio,Producto from profecha where producto='" & codigo & "'"
               Data2.Refresh
               DBGrid2.Columns(0).Width = 3500
               DBGrid2.Columns(1).Width = 1500
               DBGrid2.SetFocus
               barras2.SetFocus

End Sub

Private Sub Command2_Click()
On Error GoTo cmd89911_err
Data2.Recordset.Delete
Exit Sub
cmd89911_err:
Exit Sub
End Sub

Private Sub Command3_Click()
Dim found As Integer
Dim buf2 As String
buf2 = ""
If Len(barras2) = 0 Then
   barras2.SetFocus
   Exit Sub
End If
If Frame2.Caption = "LOTES" Then
found = grabar_barras()
barras2.SetFocus
Exit Sub
End If
If Frame2.Caption = "NUMERO SERIES" Then
found = grabar_barras()
barras2.SetFocus
Exit Sub
End If
found = valida_barras20("" & barras2, buf2) 'si existe el codigo en la database
If found = 1 Then
   MsgBox "Ya existe Barras Ingresado " + buf2, 48, "Aviso"
   barras2 = ""
   barras2.SetFocus
   Exit Sub
End If
buf2 = ""
found = valida_barras2("" & barras2, buf2) 'si existe la barra en producto
If found = 1 Then
   MsgBox "Ya existe Barras Ingresado " + buf2, 48, "Aviso"
   barras2 = ""
   barras2.SetFocus
   Exit Sub
End If
found = grabar_barras()
barras2.SetFocus
End Sub

Private Sub Command4_Click()
dlo132_Click
End Sub


Private Sub Command5_Click()
Label56_Click
End Sub

Private Sub Command6_Click()
dlo132_Click
End Sub

Private Sub Command7_Click()
Dim found As Integer
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
FrmChart.producto = codigo
FrmChart.acu = "V"
FrmChart.docu = "1"
FrmChart.Show 1

End Sub



Private Sub Command8_Click()
Dim found As Integer
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
FrmChart.producto = codigo
FrmChart.acu = "C"
FrmChart.docu = "1"
FrmChart.Show 1

End Sub

Private Sub Command9_Click()
dlo132_Click
End Sub

Private Sub costo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fechauc.SetFocus
End Sub

Private Sub costo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   rcodigo.SetFocus
   Exit Sub
End If

End Sub

Private Sub costop_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes
monedav.SetFocus

End Sub

Private Sub costop_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   costou.SetFocus
   Exit Sub
End If

End Sub

Private Sub costou_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes
costop.SetFocus

End Sub

Private Sub costou_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   factor.SetFocus
   Exit Sub
End If

End Sub

Private Sub DBGrid1_DblClick()
buffer1 = ""
DBGrid1_KeyDown 13, 0
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
If KeyCode = 27 Then
   dlo132_Click
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = "1" Then
      codigo = DBGrid1.Columns(1)
      Frame1.Visible = False
      codigo.SetFocus
      codigo_KeyPress 13
   End If
   If opcion1 = "27" Then
      ccosto = DBGrid1.Columns(1)
      Frame1.Visible = False
      ccosto.SetFocus
      ccosto_KeyPress 13
   End If
   If opcion1 = "190" Then
      'tlocal = DBGrid1.Columns(1)
      'found = busca_registro()
      'Frame1.Visible = False
      'tlocal.SetFocus
      'tlocal_KeyPress 13
   End If
   If opcion1 = "29" Then
      'xccosto2 = DBGrid1.Columns(1)
      'Frame1.Visible = False
      'xccosto2.SetFocus
      'xccosto2_KeyPress 13
   End If
   If opcion1 = "30" Then
      'xccosto3 = DBGrid1.Columns(1)
      'Frame1.Visible = False
      'xccosto3.SetFocus
      'xccosto3_KeyPress 13
   End If
   If opcion1 = "31" Then
      'xccosto4 = DBGrid1.Columns(1)
      'Frame1.Visible = False
      'xccosto4.SetFocus
      'xccosto4_KeyPress 13
   End If

   If opcion1 = "2" Then
      familia = DBGrid1.Columns(1)
      Frame1.Visible = False
      familia.SetFocus
      familia_KeyPress 13
   End If
      If opcion1 = "3" Then
      subfamilia = DBGrid1.Columns(1)
      Frame1.Visible = False
      subfamilia.SetFocus
      subfamilia_KeyPress 13
   End If
   If opcion1 = "4" Then
      seccion = DBGrid1.Columns(1)
      Frame1.Visible = False
      seccion.SetFocus
      seccion_KeyPress 13
   End If
   If opcion1 = "5" Then
      marca = DBGrid1.Columns(1)
      Frame1.Visible = False
      marca.SetFocus
      marca_KeyPress 13
   End If
   If opcion1 = "6" Then
      fabrica = DBGrid1.Columns(1)
      Frame1.Visible = False
      fabrica.SetFocus
      fabrica_KeyPress 13
   End If
   If opcion1 = "7" Then
      categoria = DBGrid1.Columns(1)
      Frame1.Visible = False
      categoria.SetFocus
      categoria_KeyPress 13
   End If
   If opcion1 = "8" Then
      lineatalla = DBGrid1.Columns(1)
      Frame1.Visible = False
      lineatalla.SetFocus
      lineatalla_KeyPress 13
   End If
   If opcion1 = "9" Then
      color = DBGrid1.Columns(1)
      Frame1.Visible = False
      color.SetFocus
      color_KeyPress 13
   End If
   If opcion1 = "10" Then
      fabrica = DBGrid1.Columns(1)
      Frame1.Visible = False
      fabrica.SetFocus
      fabrica_KeyPress 13
   End If
   If opcion1 = "11" Then
      'proveedor2 = DBGrid1.Columns(1)
      'Frame1.Visible = False
      'proveedor2.SetFocus
      'proveedor2_KeyPress 13
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
         If Chr(KeyAscii) = "*" Then
            buf = ""
            buffer = buf
         End If
         If KeyAscii <> 13 Then
            buffer = buffer + buf
         End If
         buf = buffer
         found = ejecuta(0)
         If found = 0 Then
             ejecuta (1)
         End If
End If
End Sub

Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   dlo132_Click
   Exit Sub
End If

End Sub

Private Sub DBGrid4_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo cmd56_err
Dim found As Integer
If KeyCode = &H2E Then  'borrar linea
   found = borra_proveedor("" & Data4.Recordset.Fields("codigo"), "" & codigo)
   If found = 1 Then
      busca_selec_proveedor
      carga_proveedor
   End If
End If
If KeyCode = &H76 Then  'f7
   rcodigo = "" & Data4.Recordset.Fields("codigop")
   costo = "" & Data4.Recordset.Fields("costo")
   fechauc = "" & Data4.Recordset.Fields("fecha")
   Frame9.Visible = True
   rcodigo.SetFocus
End If
Exit Sub
cmd56_err:
Exit Sub
End Sub

Private Sub DBGrid5_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
   dlo132_Click
End If
If KeyCode = 13 Then
   buffer1 = "" & Data5.Recordset.Fields("codigo")
   dlo132_Click
End If

End Sub

Private Sub descorto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If Len(descorto) = 0 Then
   descorto = Mid$(descripcio, 1, 22)
End If
presenta.SetFocus

End Sub

Private Sub descorto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   descripcio.SetFocus
   Exit Sub
End If

End Sub

Private Sub descripcio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If Len(descripcio) = 0 Then Exit Sub
descorto.SetFocus

End Sub



Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   barras.SetFocus
   Exit Sub
End If

End Sub

Private Sub djuer1_Click()
If Frame4.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
reporgen.NAMETABLA = "producto"
reporgen.Show 1

End Sub

Private Sub dlo132_Click()
If Frame1.Visible = True Then
   If Frame10.Visible = True Then
      Frame10.Visible = False
      buffer1.SetFocus
      Exit Sub
   End If
End If
If Frame9.Visible = True Then
   Frame9.Visible = False
   DBGrid4.SetFocus
   Exit Sub
End If
If Frame4.Visible = True Then
   If Frame9.Visible = True Then
      Frame9.Visible = False
      DBGrid4.SetFocus
      Exit Sub
   End If
   Frame4.Visible = False
   fabrica.SetFocus
   Exit Sub
End If

If Frame2.Visible = True Then
   Frame2.Visible = False
   If Frame2.Caption = "LOTES" Or Frame2.Caption = "NUMERO SERIES" Then
      codigo.SetFocus
      Exit Sub
   End If
   barras.SetFocus
   Exit Sub
End If

If Frame1.Visible = True Then
If opcion1 = "1" Then
   
   Frame1.Visible = False
   codigo.SetFocus
   Exit Sub
End If
If opcion1 = "27" Then
   
   Frame1.Visible = False
   ccosto.SetFocus
   Exit Sub
End If

If opcion1 = "28" Then
   
   Frame1.Visible = False
   vecaja.SetFocus
   Exit Sub
End If
If opcion1 = "29" Then
   
   Frame1.Visible = False
   vecaja.SetFocus
   Exit Sub
End If
If opcion1 = "30" Then
   
   Frame1.Visible = False
   vecaja.SetFocus
   Exit Sub
End If
If opcion1 = "31" Then
   
   Frame1.Visible = False
   vecaja.SetFocus
   Exit Sub
End If


If opcion1 = "2" Then
   
   Frame1.Visible = False
   familia.SetFocus
   Exit Sub
End If
If opcion1 = "3" Then
   
   Frame1.Visible = False
   subfamilia.SetFocus
   Exit Sub
End If
If opcion1 = "4" Then
   
   Frame1.Visible = False
   seccion.SetFocus
   Exit Sub
End If
If opcion1 = "5" Then
   
   Frame1.Visible = False
   marca.SetFocus
   Exit Sub
End If
If opcion1 = "6" Then
   
   Frame1.Visible = False
   fabrica.SetFocus
   Exit Sub
End If
If opcion1 = "7" Then
   
   Frame1.Visible = False
   categoria.SetFocus
   Exit Sub
End If
If opcion1 = "8" Then
   
   Frame1.Visible = False
   lineatalla.SetFocus
   Exit Sub
End If
If opcion1 = "9" Then
   
   Frame1.Visible = False
   color.SetFocus
   Exit Sub
End If
If opcion1 = "10" Then
   Frame1.Visible = False
   fabrica.SetFocus
   Exit Sub
End If
If opcion1 = "11" Then
   'Frame1.Visible = False
   'proveedor2.SetFocus
   Exit Sub
End If

End If
tprodup.Hide
Unload tprodup
End Sub



Private Sub fabrica_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
xproveedor.SetFocus

End Sub

Private Sub fabrica_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   color.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_fabrica
End If

End Sub

Private Sub factor_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
costou.SetFocus

End Sub

Private Sub factor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   unidad.SetFocus
   Exit Sub
End If

End Sub

Private Sub factor1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
pventa1.SetFocus

End Sub

Private Sub factor1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   unidad1.SetFocus
   Exit Sub
End If

End Sub

Private Sub factor2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
pventa2.SetFocus

End Sub

Private Sub factor2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   unidad2.SetFocus
   Exit Sub
End If

End Sub

Private Sub factor3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
pventa3.SetFocus

End Sub

Private Sub factor3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   unidad3.SetFocus
   Exit Sub
End If

End Sub

Private Sub factor4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
pventa4.SetFocus

End Sub

Private Sub factor4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   unidad4.SetFocus
   Exit Sub
End If

End Sub

Private Sub familia_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If Len(familia) = 0 Then
      consulta_Familia
      Exit Sub
End If
subfamilia.SetFocus

End Sub

Private Sub familia_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   presenta.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_Familia
End If
If KeyCode = &H76 Then  'f7
   tfamilia.Show 1
End If


End Sub

Private Sub fechauc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub fechauc_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   costo.SetFocus
   Exit Sub
End If
End Sub

Private Sub fechavence_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
End Sub

Private Sub fechavence_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   costop.SetFocus
   Exit Sub
End If

End Sub

Private Sub flete_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
igv.SetFocus
End Sub

Private Sub flete_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   pesokgr.SetFocus
   Exit Sub
End If

End Sub

Private Sub Form_Activate()
paridad = "T/C:" & busca_cambio()
carga_local1
End Sub
Sub carga_local1()
Dim mytablex As Table
local1.Clear
'local2.Clear
local1.AddItem "*"
Set mytablex = mydbxglo.OpenTable("tlocal")
Do
If mytablex.EOF Then Exit Do
local1.AddItem "" & mytablex.Fields("codigo")
'local2.AddItem "" & mytablex.Fields("codigo")
mytablex.MoveNext
Loop
mytablex.Close
local1.ListIndex = 0
'local2.ListIndex = 0
End Sub


Private Sub Form_Load()

serie.Clear
serie.AddItem "N"
serie.AddItem "S"
serie.ListIndex = 0

peso.AddItem "N"
peso.AddItem "S"
peso.ListIndex = 0

servicio.AddItem "N"
servicio.AddItem "S"
servicio.ListIndex = 0

vtaund.AddItem "N"
vtaund.AddItem "S"
vtaund.ListIndex = 0

oferta.AddItem "N"
oferta.AddItem "S"
oferta.ListIndex = 0

vecaja.AddItem "S"
vecaja.AddItem "N"
serie.ListIndex = 0

estado.AddItem "S"
estado.AddItem "N"
estado.ListIndex = 0


monedac.AddItem "S"
monedac.AddItem "D"
monedac.ListIndex = 0

monedav.AddItem "S"
monedav.AddItem "D"
monedav.ListIndex = 0

local2.Clear
local2.AddItem "01"
local2.AddItem "02"
local2.AddItem "03"
local2.AddItem "04"
local2.ListIndex = 0
End Sub
Sub inicializa()
Dim found As Integer
'l1 = ""
'l2 = ""
'l3 = ""
'l4 = ""
flete = ""
percepcion = ""
fotonombre = ""
'ccosto = ""
xproveedor.Clear




barras = ""
barras2 = ""
descripcio = ""
descorto = ""
presenta = ""
familia = ""
subfamilia = ""
seccion = ""
marca = ""
categoria = ""
lineatalla = ""
color = ""
fabrica = ""
'proveedor1 = ""
'proveedor2 = ""
'proveedor3 = ""
'proveedor4 = ""

'codprov1 = ""
'codprov2 = ""
'codprov3 = ""
'codprov4 = ""

serie.ListIndex = 0
peso.ListIndex = 0
servicio.ListIndex = 0
vtaund.ListIndex = 0
oferta.ListIndex = 0
vecaja.ListIndex = 0
estado.ListIndex = 0
igv = "19"
isc = ""
pesokgr = ""
comision = ""
monedac.ListIndex = 0
unidad = "UND"
factor = "1"
costop = ""
costou = ""
fechavence = ""
monedav.ListIndex = 0
unidad1 = "UND"
unidad2 = ""
unidad3 = ""
unidad4 = ""
unidad5 = ""
unidad6 = ""
unidad7 = ""
unidad8 = ""
unidad9 = ""
unidad10 = ""
'saldoini = ""

factor1 = "1"
factor2 = ""
factor3 = ""
factor4 = ""
factor5 = ""
factor6 = ""
factor7 = ""
factor8 = ""
factor9 = ""
factor10 = ""

pventa1 = ""
pventa2 = ""
pventa3 = ""
pventa4 = ""
pventa5 = ""
pventa6 = ""
pventa7 = ""
pventa8 = ""
pventa9 = ""
pventa10 = ""

margen1 = ""
margen2 = ""
margen3 = ""
margen4 = ""
margen5 = ""
margen6 = ""
margen7 = ""
margen8 = ""
margen9 = ""
margen10 = ""
minimo11 = ""
minimo12 = ""
minimo13 = ""
minimo14 = ""
minimo15 = ""

maximo11 = ""
maximo12 = ""
maximo13 = ""
maximo14 = ""
maximo15 = ""

pventa11 = ""
pventa12 = ""
pventa13 = ""
pventa14 = ""
pventa15 = ""

margen11 = ""
margen12 = ""
margen13 = ""
margen14 = ""
margen15 = ""
fechai11 = ""
fechaf11 = ""
fechaid = ""
fechafd = ""
dscto = ""
'minimo = ""
'maximo = ""

found = busca_parame(2)

End Sub
Function borra_registro()

Dim mytablex As Table
Dim tmp As String
Dim sw As Integer
Dim found As Integer
sw = 0
On Error GoTo cmd3_err
tmp = ""

Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "producto"
mytablex.Seek "=", codigo
If Not mytablex.NoMatch Then
   If MsgBox("Desea Borra el registro", 1, "Aviso") = "1" Then
      tmp = "" & mytablex.Fields("producto")
      mytablex.Delete
      borra_registro = 1
      sw = 1
   End If
End If
'------------------------------------- ------------
If sw = 1 Then
   borra_almacen_producto tmp
   
End If

mytablex.Close

Exit Function
cmd3_err:
MsgBox "Mensaje:" + error$, 48, "Aviso"
mytablex.Close
 
Exit Function

End Function
Sub borra_almacen_producto(tmp As String)
On Error GoTo cmd34_err
    mydbxglo.Execute "DELETE FROM ALMACEN WHERE producto='" & tmp & "'"
    mydbxglo.Execute "DELETE FROM productob WHERE producto='" & tmp & "'"
    mydbxglo.Execute "DELETE FROM codprov WHERE producto='" & tmp & "'"
    mydbxglo.Execute "DELETE FROM precios WHERE producto='" & tmp & "'"
    Exit Sub
cmd34_err:

Exit Sub
End Sub
Function busca_registro()
Dim mytablex As Table
Dim sw As Integer
sw = 0
Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "producto"
mytablex.Seek "=", codigo
If Not mytablex.NoMatch Then
   pone_registro mytablex
   busca_registro = 1
   sw = 1
End If
'------------------------------------- ------------
mytablex.Close
If sw = 1 Then
   carga_proveedor
End If
End Function
Sub pone_registro(mytablex As Table)
Dim found As Integer
Dim i As Integer
foto = LoadPicture()
fotonombre = "" & mytablex.Fields("fotonombre")
If Len(fotonombre) > 0 Then
If Existe_archivo(fotonombre) > 0 Then
   foto = LoadPicture(fotonombre)
End If
End If
'l1 = "" & mytablex.Fields("l1")
'l2 = "" & mytablex.Fields("l2")
'l3 = "" & mytablex.Fields("l3")
'l4 = "" & mytablex.Fields("l4")
percepcion = "" & mytablex.Fields("percepcion")
codigo = "" & mytablex.Fields("producto")
barras = "" & mytablex.Fields("barras")
'barras2 = "" & mytablex.Fields("barras2")
descripcio = "" & mytablex.Fields("descripcio")
descorto = "" & mytablex.Fields("descorto")
presenta = "" & mytablex.Fields("presenta")
familia = "" & mytablex.Fields("familia")
subfamilia = "" & mytablex.Fields("subfamilia")
seccion = "" & mytablex.Fields("seccion")
marca = "" & mytablex.Fields("marca")
categoria = "" & mytablex.Fields("categoria")
lineatalla = "" & mytablex.Fields("linea")
color = "" & mytablex.Fields("color")
flete = "" & mytablex.Fields("flete")
fabrica = "" & mytablex.Fields("fabrica")
'proveedor1 = "" & mytablex.Fields("proveedor1")
'proveedor2 = "" & mytablex.Fields("proveedor2")
'proveedor3 = "" & mytablex.Fields("proveedor3")
'proveedor4 = "" & mytablex.Fields("proveedor4")
'codprov1 = "" & mytablex.Fields("codprov1")
'codprov2 = "" & mytablex.Fields("codprov2")
'codprov3 = "" & mytablex.Fields("codprov3")
'codprov4 = "" & mytablex.Fields("codprov4")
serie.ListIndex = 0
If "" & mytablex.Fields("serie") = "S" Then
serie.ListIndex = 1
End If
peso.ListIndex = 0
If "" & mytablex.Fields("peso") = "S" Then
peso.ListIndex = 1
End If
servicio.ListIndex = 0
If "" & mytablex.Fields("servicio") = "S" Then
servicio.ListIndex = 1
End If
vtaund.ListIndex = 0
If "" & mytablex.Fields("vtaund") = "S" Then
vtaund.ListIndex = 1
End If
oferta.ListIndex = 0
If "" & mytablex.Fields("oferta") = "S" Then
oferta.ListIndex = 1
End If
vecaja.ListIndex = 0
If "" & mytablex.Fields("vecaja") = "S" Then
vecaja.ListIndex = 1
End If
estado.ListIndex = 0
If "" & mytablex.Fields("estado") <> "S" Then
   estado.ListIndex = 1
End If

igv = "" & mytablex.Fields("igv")
isc = "" & mytablex.Fields("isc")
pesokgr = "" & mytablex.Fields("pesokgr")
comision = "" & mytablex.Fields("comision")
monedac.ListIndex = 0
If "" & mytablex.Fields("monedac") = "D" Then
monedac.ListIndex = 1
End If
unidad = "" & mytablex.Fields("unidad")
factor = "" & mytablex.Fields("factor")
costop = "" & mytablex.Fields("costop")
costou = "" & mytablex.Fields("costou")
ccosto = "" & mytablex.Fields("ccosto")
fechavence = "" & mytablex.Fields("fechavence")
monedav.ListIndex = 0
If "" & mytablex.Fields("monedav") = "D" Then
monedav.ListIndex = 1
End If
carga_precios "" & codigo
If Len(unidad1) = 0 Then
   unidad1 = "UND"
End If
If Len(factor1) = 0 Then
   factor1 = "1"
End If
'minimo = "" & mytablex.Fields("minimo")
'maximo = "" & mytablex.Fields("maximo")
'ccosto = "" & mytablex.Fields("ccosto")
'For i = 1 To 15
'calcula_margenes i, 0
'Next i
'found = busca_bodega("" & codigo, "" & bodega, 0)
calcula_margenes
End Sub
Sub grabando(mytablex As Table)
Dim found As Integer
calcula_margenes
'mytablex.Fields("xlocal1") = "DONNA1"
'mytablex.Fields("xlocal1") = "DONNA2"
'mytablex.Fields("xlocal1") = "DONNA3"
'mytablex.Fields("xlocal1") = "DIMAGA"
'mytablex.Fields("xccosto1") = xccosto1
'mytablex.Fields("xccosto2") = xccosto2
'mytablex.Fields("xccosto3") = xccosto3
'mytablex.Fields("xccosto4") = xccosto4

'mytablex.Fields("l1") = Val(l1)
'mytablex.Fields("l2") = Val(l2)
'mytablex.Fields("l3") = Val(l3)
'mytablex.Fields("l4") = Val(l4)
mytablex.Fields("fotonombre") = "" & fotonombre
'mytablex.Fields("c11") = "0"
'mytablex.Fields("c12") = "0"
'mytablex.Fields("c13") = "0"
'mytablex.Fields("c14") = "0"

'mytablex.Fields("c21") = "0"
'mytablex.Fields("c22") = "0"
'mytablex.Fields("c23") = "0"
'mytablex.Fields("c24") = "0"

'mytablex.Fields("c31") = "0"
'mytablex.Fields("c32") = "0"
'mytablex.Fields("c33") = "0"
'mytablex.Fields("c34") = "0"

'mytablex.Fields("c41") = "0"
'mytablex.Fields("c42") = "0"
'mytablex.Fields("c43") = "0"
'mytablex.Fields("c44") = "0"

mytablex.Fields("percepcion") = Val(percepcion)
mytablex.Fields("producto") = codigo
mytablex.Fields("flete") = Val(flete)
'mytablex.Fields("ccosto") = ccosto
mytablex.Fields("barras") = barras
mytablex.Fields("descripcio") = descripcio
mytablex.Fields("descorto") = descorto
mytablex.Fields("presenta") = presenta
mytablex.Fields("familia") = familia
mytablex.Fields("subfamilia") = subfamilia
mytablex.Fields("seccion") = seccion
mytablex.Fields("marca") = marca
mytablex.Fields("categoria") = categoria
mytablex.Fields("linea") = lineatalla
mytablex.Fields("color") = color
mytablex.Fields("fabrica") = fabrica
'mytablex.Fields("proveedor1") = proveedor1
'mytablex.Fields("proveedor2") = proveedor2
'mytablex.Fields("proveedor3") = proveedor3
'mytablex.Fields("proveedor4") = proveedor4

'mytablex.Fields("codprov1") = codprov1
'mytablex.Fields("codprov2") = codprov2
'mytablex.Fields("codprov3") = codprov3
'mytablex.Fields("codprov4") = codprov4

mytablex.Fields("serie") = serie
mytablex.Fields("peso") = peso
mytablex.Fields("servicio") = servicio
mytablex.Fields("vtaund") = vtaund
mytablex.Fields("oferta") = oferta
mytablex.Fields("vecaja") = vecaja
mytablex.Fields("estado") = estado
mytablex.Fields("igv") = Val(igv)
mytablex.Fields("isc") = Val(isc)
mytablex.Fields("pesokgr") = Val(pesokgr)
mytablex.Fields("comision") = Val(comision)
mytablex.Fields("monedac") = monedac
mytablex.Fields("unidad") = unidad
mytablex.Fields("factor") = Val(factor)
mytablex.Fields("costou") = Val(costou)
mytablex.Fields("costop") = Val(costop)
If IsDate(fechavence) Then
   mytablex.Fields("fechavence") = fechavence
End If
mytablex.Fields("monedav") = monedav
'found = busca_bodega("" & codigo, "" & bodega, 1)
End Sub

Private Sub foto_Click()
CommonDialog1.DialogTitle = "Seleccione un archivo Grafico"
CommonDialog1.InitDir = globaldir & "\grafico"
CommonDialog1.Filter = "Archivos Grafico|*.jpg"
CommonDialog1.ShowOpen
'Si seleccionamos un archivo mostramos la ruta
If CommonDialog1.filename <> "" Then
   fotonombre = CommonDialog1.filename
   foto = LoadPicture(fotonombre)
Else
   'Si no mostramos un texto de advertencia de que no se seleccionó _   ninguno, ya que FileName devuelve una cadena vacía
   'Label1 = "No se seleccionó ningún archivo"
End If
End Sub

Private Sub grba1_Click()
Dim found As Integer
If Frame4.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
found = grabar()
If found = 1 Then
   consulta_bodega
   codigo.SetFocus
End If
End Sub

Private Sub igv_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
isc.SetFocus

End Sub

Private Sub igv_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   flete.SetFocus
   Exit Sub
End If

End Sub

Private Sub isc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
percepcion.SetFocus

End Sub

Private Sub isc_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   igv.SetFocus
   Exit Sub
End If

End Sub

Private Sub Label1_Click()
cmdSort_Click
End Sub


Function grabar()
Dim found As Integer
Dim mytablex As Table

Dim sw As Integer
found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If
sw = 0

Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "producto"
mytablex.Seek "=", codigo
If mytablex.NoMatch Then
   If MsgBox("Desea Adicionar?", 1, "Aviso") = 1 Then
      mytablex.AddNew
      grabando mytablex
      mytablex.Update
      found = busca_parame(1)
      grabar = 1
      sw = 1
   End If
End If
If Not mytablex.NoMatch Then
   If MsgBox("Desea Reescribir?", 1, "Aviso") = 1 Then
   mytablex.Edit
   grabando mytablex
   mytablex.Update
   grabar = 1
   sw = 1
   End If
End If
If sw = 1 Then  'graba los precios de locales
   graba_precios "" & codigo
End If
'------------------------------------- ------------
mytablex.Close

End Function

Function valida()
Dim found As Integer
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Function
End If
If insumo.Value = 1 Then
If Mid$(codigo, 1, 1) <> "I" Then
   MsgBox "Codigo debe empezar con I", 48, "Aviso"
   codigo.SetFocus
   Exit Function
End If
End If
If insumo.Value = 0 Then
If Mid$(codigo, 1, 1) = "I" Then
   MsgBox "Codigo No debe empezar con I,por ser usado para insumo", 48, "Aviso"
   codigo.SetFocus
   Exit Function
End If
End If

If Len(barras) > 0 Then
   found = valida_barras("" & barras)
   If found = 1 Then
      Exit Function
   End If
End If
If Len(descripcio) = 0 Then
   descripcio.SetFocus
   Exit Function
End If

If Len(familia) = 0 Then Exit Function
found = busca_familia()
If found = 0 Then
   MsgBox "No existe Familia", 48, "Aviso"
   familia.SetFocus
   Exit Function
End If
If Len(subfamilia) > 0 Then
found = busca_subfamilia()
If found = 0 Then
   MsgBox "No existe SubFamilia", 48, "Aviso"
   subfamilia.SetFocus
   Exit Function
End If
End If
If Len(seccion) > 0 Then
found = busca_seccion()
If found = 0 Then
   MsgBox "No existe Seccion", 48, "Aviso"
   seccion.SetFocus
   Exit Function
End If
End If
If Len(categoria) > 0 Then
found = busca_categoria()
If found = 0 Then
   MsgBox "No existe Categoria", 48, "Aviso"
   categoria.SetFocus
   Exit Function
End If
End If
If Len(marca) > 0 Then
found = busca_marca()
If found = 0 Then
   MsgBox "No existe Marca", 48, "Aviso"
   marca.SetFocus
   Exit Function
End If
End If
If Len(color) > 0 Then
found = busca_color()
If found = 0 Then
   MsgBox "No existe Color", 48, "Aviso"
   color.SetFocus
   Exit Function
End If
End If
If Len(lineatalla) > 0 Then
found = busca_lineatalla()
If found = 0 Then
   MsgBox "No existe Talla", 48, "Aviso"
   lineatalla.SetFocus
   Exit Function
End If
End If
If Len(ccosto) > 0 Then
found = busca_ccosto()
If found = 0 Then
   MsgBox "No existe Centro Costo", 48, "Aviso"
   ccosto.SetFocus
   Exit Function
End If

End If




If Len(unidad1) = 0 Then
   unidad1 = "UND"
End If
If Val(factor1) = 0 Then
   factor1 = "1"
End If

valida = 1
End Function

Sub busca_selec_proveedor()
Dim buf As String
   buf = "select codprov.Codigo,proveedo.nombre,codprov.codigop,Codprov.Costo,Codprov.Fecha from codprov left join proveedo on codprov.codigo=proveedo.codigo where codprov.producto='" & codigo & "'"
   Frame4.Visible = True
   Data4.Connect = "foxpro 2.5;"
   Data4.DatabaseName = globaldir
   Data4.RecordSource = buf
   Data4.Refresh
   DBGrid4.Columns(0).Width = 1500
   DBGrid4.Columns(1).Width = 4700
   DBGrid4.Columns(2).Width = 1500
   DBGrid4.SetFocus

End Sub

Private Sub Label14_Click()
busca_selec_proveedor
End Sub

Private Sub Label2_Click()
Dim found As Integer
If Len(codigo) = 0 Then Exit Sub
found = busca_registro()
If found = 0 Then
   codigo.SetFocus
   Exit Sub
End If
Frame2.Visible = True
Frame2.Caption = "CODIGO BARRAS"
barras2 = ""
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = "select Barras,Producto from productb where producto='" & codigo & "'"
               Data2.Refresh
               DBGrid2.Columns(0).Width = 3500
               DBGrid2.Columns(1).Width = 1500
               DBGrid2.SetFocus
               'barras2.SetFocus
End Sub


Private Sub Label56_Click()
Dim found As Integer
If Frame4.Visible = True Then Exit Sub
If Frame2.Visible = True Then
   borrar_barras
   Label2_Click
   Exit Sub
End If
If Frame1.Visible = True Then Exit Sub
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
Cargastk.producto = codigo
Cargastk.descripcio = descripcio
Cargastk.Show 1
consulta_bodega
End Sub

Private Sub lineatalla_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
color.SetFocus

End Sub

Private Sub lineatalla_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   categoria.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_talla
End If
If KeyCode = &H76 Then  'f7
   tlinea.Show 1
End If



End Sub

Private Sub local1_Change()
consulta_bodega
End Sub

Private Sub local1_Click()
consulta_bodega
End Sub

Private Sub local1_KeyPress(KeyAscii As Integer)
consulta_bodega
End Sub

Private Sub local2_Click()
carga_precios "" & codigo
End Sub

Private Sub local2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
carga_precios "" & codigo
End Sub

Private Sub marca_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
categoria.SetFocus

End Sub

Private Sub marca_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   seccion.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_marca
End If
If KeyCode = &H76 Then  'f7
   tmarca.Show 1
End If

End Sub

Private Sub margen1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes
unidad2.SetFocus

End Sub

Private Sub margen1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   pventa1.SetFocus
   Exit Sub
End If

End Sub

Private Sub margen10_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub margen11_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes

End Sub

Private Sub margen12_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes

End Sub

Private Sub margen13_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes

End Sub

Private Sub margen14_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes

End Sub

Private Sub margen15_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes

End Sub

Private Sub margen2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes
unidad3.SetFocus

End Sub

Private Sub margen2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   pventa2.SetFocus
   Exit Sub
End If

End Sub

Private Sub margen3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes
unidad4.SetFocus

End Sub

Private Sub margen3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   pventa3.SetFocus
   Exit Sub
End If

End Sub

Private Sub margen4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes
unidad5.SetFocus

End Sub

Private Sub margen4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   pventa4.SetFocus
   Exit Sub
End If

End Sub

Private Sub margen5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes
End Sub

Private Sub margen6_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub margen7_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub margen8_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub margen9_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub


Private Sub monedac_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
unidad.SetFocus

End Sub

Private Sub monedac_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   percepcion.SetFocus
   Exit Sub
End If

End Sub

Private Sub monedav_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
unidad1.SetFocus

End Sub

Private Sub monedav_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   costop.SetFocus
   Exit Sub
End If

End Sub

Private Sub oferta_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
vecaja.SetFocus

End Sub

Private Sub oferta_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   vtaund.SetFocus
   Exit Sub
End If

End Sub

Private Sub percepcion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
monedac.SetFocus
End Sub

Private Sub percepcion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   isc.SetFocus
   Exit Sub
End If

End Sub

Private Sub peso_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
servicio.SetFocus

End Sub

Private Sub peso_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   serie.SetFocus
   Exit Sub
End If

End Sub

Private Sub pesokgr_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
flete.SetFocus

End Sub

Private Sub pesokgr_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   comision.SetFocus
   Exit Sub
End If

End Sub

Private Sub presenta_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
familia.SetFocus

End Sub

Private Sub presenta_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   descorto.SetFocus
   Exit Sub
End If

End Sub


Private Sub proveedor1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
xproveedor.SetFocus

End Sub

Private Sub proveedor1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   color.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_proveedor1
End If
If KeyCode = &H76 Then  'f7
   tproveedo.Show 1
End If


End Sub

Private Sub pventa1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes
margen1.SetFocus

End Sub

Private Sub pventa1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   factor1.SetFocus
   Exit Sub
End If

End Sub

Private Sub pventa10_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub pventa11_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes
End Sub

Private Sub pventa12_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub pventa13_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub pventa14_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub pventa15_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub pventa2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes
margen2.SetFocus

End Sub

Private Sub pventa2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   factor2.SetFocus
   Exit Sub
End If

End Sub

Private Sub pventa3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes
margen3.SetFocus

End Sub

Private Sub pventa3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   factor3.SetFocus
   Exit Sub
End If

End Sub

Private Sub pventa4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
calcula_margenes
margen4.SetFocus

End Sub

Private Sub pventa4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   factor4.SetFocus
   Exit Sub
End If

End Sub

Private Sub pventa5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes
End Sub

Private Sub pventa6_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub pventa7_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub pventa8_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub pventa9_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes

End Sub

Private Sub rcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Command9_Click
   Exit Sub
End If
costo.SetFocus
End Sub

Private Sub rect398912_Click()
Dim found As Integer
If Frame4.Visible = True Then Exit Sub
If Frame2.Visible = True Then
   borrar_barras
   Label2_Click
   Exit Sub
End If
If Frame1.Visible = True Then Exit Sub
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
treceta.producto = codigo
treceta.linea = lineatalla
treceta.descripcio = descripcio
treceta.nro = "1"
treceta.Show 1
End Sub

Private Sub seccion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
marca.SetFocus

End Sub

Private Sub seccion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   subfamilia.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_seccion
End If
If KeyCode = &H76 Then  'f7
   tseccion.Show 1
End If

End Sub

Private Sub serie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
peso.SetFocus

End Sub

Private Sub serie_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   xproveedor.SetFocus
   Exit Sub
End If

End Sub

Private Sub servicio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
vtaund.SetFocus

End Sub

Private Sub servicio_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   peso.SetFocus
   Exit Sub
End If

End Sub

Private Sub subfamilia_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
seccion.SetFocus

End Sub

Private Sub subfamilia_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   familia.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_subFamilia
End If
If KeyCode = &H76 Then  'f7
   tsubfami.Show 1
End If

End Sub

Private Sub tlocal_KeyPress(KeyAscii As Integer)
KeyAscii = 0
If KeyAscii <> 13 Then Exit Sub
ccosto.SetFocus
End Sub

Private Sub tlocal_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   'consulta_local
End If

End Sub

Private Sub unidad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
factor.SetFocus

End Sub

Private Sub unidad_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   monedac.SetFocus
   Exit Sub
End If

End Sub

Private Sub unidad1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
factor1.SetFocus

End Sub

Private Sub unidad1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   monedav.SetFocus
   Exit Sub
End If

End Sub

Private Sub unidad2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
factor2.SetFocus

End Sub

Private Sub unidad2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   margen1.SetFocus
   Exit Sub
End If

End Sub

Private Sub unidad3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
factor3.SetFocus

End Sub

Private Sub unidad3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   margen2.SetFocus
   Exit Sub
End If

End Sub

Private Sub unidad4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
factor4.SetFocus

End Sub

Private Sub unidad4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   margen3.SetFocus
   Exit Sub
End If

End Sub

Private Sub unidad5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   margen4.SetFocus
   Exit Sub
End If

End Sub

Private Sub vecaja_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
comision.SetFocus

End Sub

Private Sub vecaja_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   oferta.SetFocus
   Exit Sub
End If

End Sub

Private Sub vtaund_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
oferta.SetFocus

End Sub

Private Sub vtaund_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   servicio.SetFocus
   Exit Sub
End If

End Sub
Sub consulta_Familia()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "familia"
Combo3.ListIndex = 0
opcion1 = "2"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)
DBGrid1.SetFocus
End Sub
Sub consulta_subFamilia()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "Subfamilia"
Combo3.ListIndex = 0

opcion1 = "3"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)
DBGrid1.SetFocus
End Sub
Sub consulta_seccion()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "seccion"
Combo3.AddItem "Descripcio"
Combo3.ListIndex = 0
opcion1 = "4"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)
DBGrid1.SetFocus
End Sub
Sub consulta_local()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Codigo"
Combo3.AddItem "Nombre"
Combo3.ListIndex = 0
opcion1 = "190"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)
DBGrid1.SetFocus

End Sub

Sub consulta_marca()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "marca"
Combo3.ListIndex = 0
opcion1 = "5"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)
DBGrid1.SetFocus
End Sub
Sub consulta_fabrica()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "codigo"
Combo3.ListIndex = 0
opcion1 = "6"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)

DBGrid1.SetFocus
End Sub
Sub consulta_categoria()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "Categoria"
Combo3.ListIndex = 0
opcion1 = "7"
Frame1.Visible = True
found = ejecuta(1)
buffer = ""

DBGrid1.SetFocus
End Sub
Sub consulta_talla()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "Linea"
Combo3.ListIndex = 0


opcion1 = "8"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)

DBGrid1.SetFocus
End Sub
Sub consulta_color()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "Color"
Combo3.ListIndex = 0

opcion1 = "9"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)

DBGrid1.SetFocus
End Sub
Sub consulta_proveedor1()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Nombre"
Combo3.AddItem "codigo"
Combo3.ListIndex = 0
opcion1 = "10"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)
DBGrid1.SetFocus
End Sub
Sub consulta_proveedor2()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Nombre"
Combo3.AddItem "codigo"
Combo3.ListIndex = 0


opcion1 = "11"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)

DBGrid1.SetFocus
End Sub
Function busca_familia()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("familia")
mytablex.Index = "familia"
mytablex.Seek "=", familia
If Not mytablex.NoMatch Then
   busca_familia = 1
   If Val(margen1) = 0 Then
      margen1 = "" & mytablex.Fields("margen1")
   End If
   If Val(margen2) = 0 Then
      margen2 = "" & mytablex.Fields("margen2")
   End If
   If Val(margen3) = 0 Then
      margen3 = "" & mytablex.Fields("margen3")
   End If
   If Val(margen4) = 0 Then
      margen4 = "" & mytablex.Fields("margen4")
   End If
   If Val(margen5) = 0 Then
      margen5 = "" & mytablex.Fields("margen5")
   End If
End If
'------------------------------------- ------------
mytablex.Close


End Function
Function busca_ccosto()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("ccosto")
mytablex.Index = "ccosto"
mytablex.Seek "=", ccosto
If Not mytablex.NoMatch Then
   busca_ccosto = 1
End If
'------------------------------------- ------------
mytablex.Close

End Function
Function busca_subfamilia()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("subfamil")
mytablex.Index = "subfamilia"
mytablex.Seek "=", familia, subfamilia
If Not mytablex.NoMatch Then
   busca_subfamilia = 1
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Function busca_seccion()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("seccion")
mytablex.Index = "seccion"
mytablex.Seek "=", seccion
If Not mytablex.NoMatch Then
   busca_seccion = 1
End If
'------------------------------------- ------------
mytablex.Close


End Function
Function busca_categoria()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("categori")
mytablex.Index = "categoria"
mytablex.Seek "=", categoria
If Not mytablex.NoMatch Then
   busca_categoria = 1
End If
'------------------------------------- ------------
mytablex.Close

End Function
Function busca_marca()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("marca")
mytablex.Index = "marca"
mytablex.Seek "=", marca
If Not mytablex.NoMatch Then
   busca_marca = 1
End If
'------------------------------------- ------------
mytablex.Close

End Function
Function busca_color()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("color")
mytablex.Index = "color"
mytablex.Seek "=", color
If Not mytablex.NoMatch Then
   busca_color = 1
End If
'------------------------------------- ------------
mytablex.Close


End Function
Function busca_lineatalla()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("linea")
mytablex.Index = "linea"
mytablex.Seek "=", lineatalla
If Not mytablex.NoMatch Then
   busca_lineatalla = 1
End If
'------------------------------------- ------------
mytablex.Close


End Function
Function busca_parame(sw As Integer)
Dim sdx As Double

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("parame")
mytablex.Index = "codigo"
mytablex.Seek "=", "01"
If Not mytablex.NoMatch Then
   If "" & mytablex.Fields("plocal") = "S" Then
      Label33.Visible = True
      local2.Visible = True
   End If
   If sw = 0 Then
      If insumo.Value = 0 Then
         sdx = Val("" & mytablex.Fields("producto")) + 1
         codigo = "" & sdx
      End If
      If insumo.Value = 1 Then
         sdx = Val("" & mytablex.Fields("insumo")) + 1
         codigo = "I" & sdx
      End If
   End If
   If sw = 3 Then
      'bodega = "" & mytablex.Fields("bodega")
      'fsaldoini = "" & mytablex.Fields("saldoini")
   End If
   If sw = 2 Then
      igv = "" & mytablex.Fields("igv")
      End If
   If sw = 1 Then
      If insumo.Value = 0 Then
      If IsNumeric(codigo) Then
         mytablex.Edit
         mytablex.Fields("producto") = codigo
         mytablex.Update
      End If
      End If
      If insumo.Value = 1 Then
      If IsNumeric(Mid$(codigo, 2, Len(codigo))) Then
         mytablex.Edit
         mytablex.Fields("insumo") = Mid$(codigo, 2, Len(codigo))
         mytablex.Update
      End If
      End If
   End If
   busca_parame = 1
End If
'------------------------------------- ------------
mytablex.Close


End Function
Function valida_barras(buf As String)

Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "barras"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   If codigo <> "" & mytablex.Fields("producto") Then
      MsgBox "Ya existe Codigo Barras en codigo:" & mytablex.Fields("producto"), 48, "Aviso"
      valida_barras = 1
   End If
End If
'------------------------------------- ------------
mytablex.Close



End Function
Sub borrar_barras()
On Error GoTo cmd1_error
Data2.Recordset.Delete
Exit Sub
cmd1_error:
Exit Sub
End Sub

Function grabar_barras()
Dim found As Integer
On Error GoTo cmd3_error
If Frame2.Caption = "LOTES" Or Frame2.Caption = "NUMERO SERIES" Then
Data2.Recordset.AddNew
Data2.Recordset.Fields("descripcio") = "" & barras2
Data2.Recordset.Fields("producto") = "" & codigo
Data2.Recordset.Update
grabar_barras = 1
Exit Function
End If
Data2.Recordset.AddNew
Data2.Recordset.Fields("barras") = "" & barras2
Data2.Recordset.Fields("producto") = "" & codigo
Data2.Recordset.Update
grabar_barras = 1
Exit Function
cmd3_error:
MsgBox "Error en grabar Barras " + error$, 48, "Aviso"
Exit Function
End Function
Function valida_barras2(buf As String, buf2 As String)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("productb")
mytablex.Index = "productb"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   valida_barras2 = 1
   buf2 = "" & mytablex.Fields("producto")
End If
'------------------------------------- ------------
mytablex.Close

End Function
Function valida_barras20(buf As String, buf2 As String)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "barras"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   valida_barras20 = 1
   buf2 = "" & mytablex.Fields("producto")
End If
'------------------------------------- ------------
mytablex.Close



End Function
Sub consulta_bodega()
Dim buf As String
buf = ""
If local1 <> "*" Then
   buf = " and local='" & local1 & "'"
End If
Data3.Connect = "foxpro 2.5;"
               Data3.DatabaseName = globaldir
               Data3.RecordSource = "select * from almacen where producto='" & codigo & "' " & buf
               Data3.Refresh
               DBGrid3.Refresh
End Sub
Sub consulta_ccosto()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "Ccosto"
Combo3.ListIndex = 0
opcion1 = "27"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)

DBGrid1.SetFocus


End Sub
Sub consulta_ccosto1()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "Ccosto"
Combo3.ListIndex = 0
opcion1 = "28"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)

DBGrid1.SetFocus

End Sub
Sub consulta_ccosto2()
Dim found As Integer

Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "Ccosto"
Combo3.ListIndex = 0


opcion1 = "29"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)

DBGrid1.SetFocus

End Sub
Sub consulta_ccosto3()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "Ccosto"
Combo3.ListIndex = 0


opcion1 = "30"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)

DBGrid1.SetFocus

End Sub
Sub consulta_ccosto4()
Dim found As Integer
Combo3.Clear
Combo3.AddItem "*"
Combo3.AddItem "Descripcio"
Combo3.AddItem "Ccosto"
Combo3.ListIndex = 0


opcion1 = "31"
Frame1.Visible = True
buffer = ""
found = ejecuta(1)

DBGrid1.SetFocus

End Sub
Function existe_proveedor(buf As String)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("proveedo")
mytablex.Index = "codigo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   existe_proveedor = 1
End If
'------------------------------------- ------------
mytablex.Close

End Function
Sub calcula_margenes()
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim acostou As String
On Error GoTo cmd786_err
sdx = Val(costou) + Val(flete)
acostou = "" & sdx
If Val(factor) <= 0 Then
   factor = "1"
End If
If Val(factor1) <= 0 Then
   factor1 = "1"
End If
If Val(costou) = 0 And Val(costop) = 0 Then
   margen1 = "0"
   margen2 = "0"
   margen3 = "0"
   margen4 = "0"
   margen5 = "0"
   margen6 = "0"
   margen7 = "0"
   margen8 = "0"
   margen9 = "0"
   margen10 = "0"
   
   margen11 = "0"
   margen12 = "0"
   margen13 = "0"
   margen14 = "0"
   margen15 = "0"
   Exit Sub
End If

          If monedac = "S" Then
             If monedav = "D" Then
                sdx = Val(acostou) / busca_cambio()
                If sdx <= 0 Then
                   sdx = 1
                End If
                acostou = "" & sdx
             End If
          End If
          If monedac = "D" Then
             If monedav = "S" Then
                sdx = Val(acostou) * busca_cambio()
                If sdx <= 0 Then
                   sdx = 1
                End If
                acostou = "" & sdx
             End If
          End If
       If Val(acostou) > 0 And Val(pventa1) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa1) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen1 = Format(sdx2, "0.00")
          GoTo siguiente1
       End If
       If Val(margen1) > 0 And Val(pventa1) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen1) / 100
          pventa1 = Format(sdx, "0.00")
          GoTo siguiente1
       End If
       If Val(acostou) <= 0 And Val(pventa1) > 0 And Val(margen1) > 0 Then
          sdx = Val(pventa1) / (1 + (Val(margen1) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente1
       End If
       
siguiente1:
       If Val(acostou) > 0 And Val(pventa2) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor2)
          sdx1 = Val(pventa2) '/ Val(factor1)
          sdx2 = (Val(sdx1) - sdx) * 100 / sdx
          margen2 = Format(sdx2, "0.00")
          GoTo siguiente2
       End If
       If Val(margen2) > 0 And Val(pventa2) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen2) / 100
          pventa2 = Format(sdx, "0.00")
          GoTo siguiente2
       End If
       If Val(acostou) <= 0 And Val(pventa2) > 0 And Val(margen2) > 0 Then
          sdx = Val(pventa2) / (1 + (Val(margen2) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente2
       End If
siguiente2:
       If Val(acostou) > 0 And Val(pventa3) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor3)
          sdx1 = Val(pventa3) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen3 = Format(sdx2, "0.00")
          GoTo siguiente3
       End If
       If Val(margen3) > 0 And Val(pventa3) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen3) / 100
          pventa3 = Format(sdx, "0.00")
          GoTo siguiente3
       End If
       If Val(acostou) <= 0 And Val(pventa3) > 0 And Val(margen3) > 0 Then
          sdx = Val(pventa3) / (1 + (Val(margen3) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente2
       End If
siguiente3:
If Val(acostou) > 0 And Val(pventa4) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor4)
          sdx1 = Val(pventa4) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen4 = Format(sdx2, "0.00")
          GoTo siguiente4
       End If
       If Val(margen4) > 0 And Val(pventa4) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen4) / 100
          pventa4 = Format(sdx, "0.00")
          GoTo siguiente4
       End If
       If Val(acostou) <= 0 And Val(pventa4) > 0 And Val(margen4) > 0 Then
          sdx = Val(pventa4) / (1 + (Val(margen4) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente4
       End If
siguiente4:
If Val(acostou) > 0 And Val(pventa5) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor5)
          sdx1 = Val(pventa5) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen5 = Format(sdx2, "0.00")
          GoTo siguiente5
       End If
       If Val(margen5) > 0 And Val(pventa5) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen5) / 100
          pventa5 = Format(sdx, "0.00")
          GoTo siguiente5
       End If
       If Val(acostou) <= 0 And Val(pventa5) > 0 And Val(margen5) > 0 Then
          sdx = Val(pventa5) / (1 + (Val(margen5) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente5
       End If
siguiente5:
If Val(acostou) > 0 And Val(pventa6) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor6)
          sdx1 = Val(pventa6) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen6 = Format(sdx2, "0.00")
          GoTo siguiente6
       End If
       If Val(margen6) > 0 And Val(pventa6) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen6) / 100
          pventa6 = Format(sdx, "0.00")
          GoTo siguiente6
       End If
       If Val(acostou) <= 0 And Val(pventa6) > 0 And Val(margen6) > 0 Then
          sdx = Val(pventa6) / (1 + (Val(margen6) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente6
       End If
siguiente6:
If Val(acostou) > 0 And Val(pventa7) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor7)
          sdx1 = Val(pventa7) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen7 = Format(sdx2, "0.00")
          GoTo siguiente7
       End If
       If Val(margen7) > 0 And Val(pventa7) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen7) / 100
          pventa7 = Format(sdx, "0.00")
          GoTo siguiente7
       End If
       If Val(acostou) <= 0 And Val(pventa7) > 0 And Val(margen7) > 0 Then
          sdx = Val(pventa7) / (1 + (Val(margen7) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente7
       End If
siguiente7:
If Val(costou) > 0 And Val(pventa8) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor8)
          sdx1 = Val(pventa8) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen8 = Format(sdx2, "0.00")
          GoTo siguiente8
       End If
       If Val(margen8) > 0 And Val(pventa8) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen8) / 100
          pventa8 = Format(sdx, "0.00")
          GoTo siguiente8
       End If
       If Val(acostou) <= 0 And Val(pventa8) > 0 And Val(margen8) > 0 Then
          sdx = Val(pventa8) / (1 + (Val(margen8) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente8
       End If
siguiente8:
If Val(acostou) > 0 And Val(pventa9) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor9)
          sdx1 = Val(pventa9) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen9 = Format(sdx2, "0.00")
          GoTo siguiente9
       End If
       If Val(margen9) > 0 And Val(pventa9) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen9) / 100
          pventa9 = Format(sdx, "0.00")
          GoTo siguiente9
       End If
       If Val(acostou) <= 0 And Val(pventa9) > 0 And Val(margen9) > 0 Then
          sdx = Val(pventa9) / (1 + (Val(margen9) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente9
       End If
siguiente9:
If Val(acostou) > 0 And Val(pventa10) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor10)
          sdx1 = Val(pventa10) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen10 = Format(sdx2, "0.00")
          GoTo siguiente10
       End If
       If Val(margen10) > 0 And Val(pventa10) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen10) / 100
          pventa2 = Format(sdx, "0.00")
          GoTo siguiente10
       End If
       If Val(acostou) <= 0 And Val(pventa10) > 0 And Val(margen10) > 0 Then
          sdx = Val(pventa10) / (1 + (Val(margen10) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente10
       End If
siguiente10:
If Val(acostou) > 0 And Val(pventa11) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa11) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen11 = Format(sdx2, "0.00")
          GoTo siguiente11
       End If
       If Val(margen11) > 0 And Val(pventa11) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen11) / 100
          pventa11 = Format(sdx, "0.00")
          GoTo siguiente11
       End If
       If Val(acostou) <= 0 And Val(pventa11) > 0 And Val(margen11) > 0 Then
          sdx = Val(pventa11) / (1 + (Val(margen11) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente11
       End If
siguiente11:
If Val(acostou) > 0 And Val(pventa12) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa12) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen12 = Format(sdx2, "0.00")
          GoTo siguiente12
       End If
       If Val(margen12) > 0 And Val(pventa12) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen12) / 100
          pventa12 = Format(sdx, "0.00")
          GoTo siguiente12
       End If
       If Val(acostou) <= 0 And Val(pventa12) > 0 And Val(margen12) > 0 Then
          sdx = Val(pventa12) / (1 + (Val(margen12) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente12
       End If
siguiente12:
If Val(acostou) > 0 And Val(pventa13) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa13) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen13 = Format(sdx2, "0.00")
          GoTo siguiente13
       End If
       If Val(margen13) > 0 And Val(pventa13) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen13) / 100
          pventa13 = Format(sdx, "0.00")
          GoTo siguiente13
       End If
       If Val(acostou) <= 0 And Val(pventa13) > 0 And Val(margen13) > 0 Then
          sdx = Val(pventa13) / (1 + (Val(margen13) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente13
       End If
siguiente13:
If Val(acostou) > 0 And Val(pventa14) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa14) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen14 = Format(sdx2, "0.00")
          GoTo siguiente14
       End If
       If Val(margen14) > 0 And Val(pventa14) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen14) / 100
          pventa14 = Format(sdx, "0.00")
          GoTo siguiente14
       End If
       If Val(acostou) <= 0 And Val(pventa14) > 0 And Val(margen14) > 0 Then
          sdx = Val(pventa14) / (1 + (Val(margen14) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente14
       End If
siguiente14:
If Val(acostou) > 0 And Val(pventa15) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa15) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen15 = Format(sdx2, "0.00")
          GoTo siguiente15
       End If
       If Val(margen15) > 0 And Val(pventa15) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen15) / 100
          pventa15 = Format(sdx, "0.00")
          GoTo siguiente10
       End If
       If Val(acostou) <= 0 And Val(pventa15) > 0 And Val(margen15) > 0 Then
          sdx = Val(pventa15) / (1 + (Val(margen15) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente15
       End If
siguiente15:
       Exit Sub
cmd786_err:
MsgBox "Error en calcula margenes", 48, "Aviso"
Exit Sub
End Sub
Function busca_cambio() As Double
Dim mytablex As Table
Dim sdx As Double
sdx = 1
Set mytablex = mydbxglo.OpenTable("parame")
mytablex.Index = "codigo"
mytablex.Seek "=", "01"
If Not mytablex.NoMatch Then
   sdx = Val("" & mytablex.Fields("paricomp"))
   If Val("" & mytablex.Fields("paricomp")) <= 0 Then
      sdx = 1
   End If
End If
busca_cambio = sdx
mytablex.Close
End Function

Sub pone_tamano()
               If opcion1 = "1" Then
               DBGrid1.Columns(0).Width = 6000
               DBGrid1.Columns(1).Width = 2000
               DBGrid1.Columns(2).Width = 1000
               DBGrid1.Columns(3).Width = 1000
               DBGrid1.Columns(4).Width = 1000
               DBGrid1.Columns(5).Width = 1000
               DBGrid1.Columns(6).Width = 1000
               DBGrid1.Columns(7).Width = 1000
               DBGrid1.Columns(8).Width = 1000
               DBGrid1.SetFocus
               End If
               If opcion1 = "2" Or opcion1 = "27" Or opcion1 = "28" Or opcion1 = "29" Or opcion1 = "30" Or opcion1 = "31" Or opcion1 = "190" Then
               DBGrid1.Columns(0).Width = 6000
               DBGrid1.Columns(1).Width = 2000
               DBGrid1.SetFocus
               End If
               If opcion1 = "3" Then
               DBGrid1.Columns(0).Width = 6000
               DBGrid1.Columns(1).Width = 2000
               DBGrid1.Columns(2).Width = 2000
               DBGrid1.SetFocus
               End If
               If opcion1 = "4" Or opcion1 = "5" Or opcion1 = "6" Or opcion1 = "7" Or opcion1 = "8" Or opcion1 = "9" Or opcion1 = "10" Or opcion1 = "11" Then
               DBGrid1.Columns(0).Width = 6000
               DBGrid1.Columns(1).Width = 2000
               DBGrid1.SetFocus
               End If

End Sub
Function borra_proveedor(buf1 As String, buf2 As String)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("codprov")
mytablex.Index = "codprov"
mytablex.Seek "=", buf1, buf2
If Not mytablex.NoMatch Then
   If MsgBox("Desea Borrar", 1, "Aviso") = 1 Then
      mytablex.Delete
      borra_proveedor = 1
   End If
End If
mytablex.Close

End Function
Sub carga_proveedor()
Dim mytablex As Table

Dim indx As Integer
xproveedor.Clear
indx = 0

Set mytablex = mydbxglo.OpenTable("codprov")
mytablex.Index = "producto"
mytablex.Seek "=", codigo
If Not mytablex.NoMatch Then
   Do
   If mytablex.EOF Then Exit Do
   If "" & mytablex.Fields("producto") = "" & codigo Then
      xproveedor.AddItem "" & mytablex.Fields("codigo") & " " & mytablex.Fields("codigop")
      indx = indx + 1
      Else: Exit Do
   End If
   mytablex.MoveNext
   Loop
End If
mytablex.Close

If indx > 0 Then
  xproveedor.ListIndex = 0
End If
End Sub

Private Sub xproveedor_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
serie.SetFocus
End Sub
Function borrar_codprov()
Dim mytablex As Table
Dim indx As Integer

Set mytablex = mydbxglo.OpenTable("codprov")
mytablex.Index = "producto"
amki1:
mytablex.Seek "=", codigo
If Not mytablex.NoMatch Then
   mytablex.Delete
   GoTo amki1
End If
mytablex.Close
xproveedor.Clear
End Function
Function graba_rcodigo()
Dim mytablex As Table


Set mytablex = mydbxglo.OpenTable("codprov")
mytablex.Index = "codprov"
mytablex.Seek "=", "" & Data4.Recordset.Fields("codigo"), codigo
If Not mytablex.NoMatch Then
   mytablex.Edit
   mytablex.Fields("codigop") = rcodigo
   mytablex.Fields("costo") = Val(costo)
   If Len(fechauc) = 10 Then
      If IsDate(fechauc) Then
         mytablex.Fields("fecha") = fechauc
      End If
   End If
   mytablex.Update
End If
mytablex.Close

End Function
Sub consulta_proveedor()
Dim buf As String
Combo2.Clear
Combo2.AddItem "Nombre"
Combo2.AddItem "Codigo"
Combo2.ListIndex = 0
xbusca1 = "*"
Command11_Click
End Sub
Sub carga_precios(buf As String)
Dim mytablex As Table
inicializa_precios
Set mytablex = mydbxglo.OpenTable("precios")
mytablex.Index = "tprecios"
mytablex.Seek "=", buf, local2
If Not mytablex.NoMatch Then
   pone_xprecio mytablex
   calcula_margenes
End If
mytablex.Close
'pventa1.SetFocus
End Sub
Sub graba_precios(buf As String)
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("precios")
mytablex.Index = "tprecios"
mytablex.Seek "=", buf, local2
If Not mytablex.NoMatch Then
   mytablex.Edit
   graba_xprecio mytablex
   mytablex.Update
End If
If mytablex.NoMatch Then
   mytablex.AddNew
   mytablex.Fields("producto") = codigo
   mytablex.Fields("local") = local2
   graba_xprecio mytablex
   mytablex.Update
End If
mytablex.Close

End Sub
Sub graba_xprecio(mytablex As Table)
mytablex.Fields("ccosto") = ccosto
mytablex.Fields("unidad1") = unidad1
mytablex.Fields("unidad2") = unidad2
mytablex.Fields("unidad3") = unidad3
mytablex.Fields("unidad4") = unidad4
mytablex.Fields("unidad5") = unidad5
mytablex.Fields("unidad6") = unidad6
mytablex.Fields("unidad7") = unidad7
mytablex.Fields("unidad8") = unidad8
mytablex.Fields("unidad9") = unidad9
mytablex.Fields("unidad10") = unidad10
mytablex.Fields("factor1") = Val(factor1)
mytablex.Fields("factor2") = Val(factor2)
mytablex.Fields("factor3") = Val(factor3)
mytablex.Fields("factor4") = Val(factor4)
mytablex.Fields("factor5") = Val(factor5)
mytablex.Fields("factor6") = Val(factor6)
mytablex.Fields("factor7") = Val(factor7)
mytablex.Fields("factor8") = Val(factor8)
mytablex.Fields("factor9") = Val(factor9)
mytablex.Fields("factor10") = Val(factor10)
mytablex.Fields("pventa1") = Val(pventa1)
mytablex.Fields("pventa2") = Val(pventa2)
mytablex.Fields("pventa3") = Val(pventa3)
mytablex.Fields("pventa4") = Val(pventa4)
mytablex.Fields("pventa5") = Val(pventa5)
mytablex.Fields("pventa6") = Val(pventa6)
mytablex.Fields("pventa7") = Val(pventa7)
mytablex.Fields("pventa8") = Val(pventa8)
mytablex.Fields("pventa9") = Val(pventa9)
mytablex.Fields("pventa10") = Val(pventa10)
mytablex.Fields("margen1") = Val(margen1)
mytablex.Fields("margen2") = Val(margen2)
mytablex.Fields("margen3") = Val(margen3)
mytablex.Fields("margen4") = Val(margen4)
mytablex.Fields("margen5") = Val(margen5)
mytablex.Fields("margen6") = Val(margen6)
mytablex.Fields("margen7") = Val(margen7)
mytablex.Fields("margen8") = Val(margen8)
mytablex.Fields("margen9") = Val(margen9)
mytablex.Fields("margen10") = Val(margen10)
mytablex.Fields("minimo11") = Val(minimo11)
mytablex.Fields("minimo12") = Val(minimo12)
mytablex.Fields("minimo13") = Val(minimo13)
mytablex.Fields("minimo14") = Val(minimo14)
mytablex.Fields("minimo15") = Val(minimo15)
mytablex.Fields("maximo11") = Val(maximo11)
mytablex.Fields("maximo12") = Val(maximo12)
mytablex.Fields("maximo13") = Val(maximo13)
mytablex.Fields("maximo14") = Val(maximo14)
mytablex.Fields("maximo15") = Val(maximo15)
mytablex.Fields("pventa11") = Val(pventa11)
mytablex.Fields("pventa12") = Val(pventa12)
mytablex.Fields("pventa13") = Val(pventa13)
mytablex.Fields("pventa14") = Val(pventa14)
mytablex.Fields("pventa15") = Val(pventa15)
mytablex.Fields("margen11") = Val(margen11)
mytablex.Fields("margen12") = Val(margen12)
mytablex.Fields("margen13") = Val(margen13)
mytablex.Fields("margen14") = Val(margen14)
mytablex.Fields("margen15") = Val(margen15)
If IsDate(fechai11) Then
mytablex.Fields("fechai11") = fechai11
End If
If IsDate(fechaf11) Then
mytablex.Fields("fechaf11") = fechaf11
End If
If IsDate(fechaid) Then
mytablex.Fields("fechaid") = fechaid
End If
If IsDate(fechafd) Then
mytablex.Fields("fechafd") = fechafd
End If


End Sub
Sub pone_xprecio(mytablex As Table)
unidad1 = "" & mytablex.Fields("unidad1")
unidad2 = "" & mytablex.Fields("unidad2")
unidad3 = "" & mytablex.Fields("unidad3")
unidad4 = "" & mytablex.Fields("unidad4")
unidad5 = "" & mytablex.Fields("unidad5")
unidad6 = "" & mytablex.Fields("unidad6")
unidad7 = "" & mytablex.Fields("unidad7")
unidad8 = "" & mytablex.Fields("unidad8")
unidad9 = "" & mytablex.Fields("unidad9")
unidad10 = "" & mytablex.Fields("unidad10")
factor1 = "" & mytablex.Fields("factor1")
factor2 = "" & mytablex.Fields("factor2")
factor3 = "" & mytablex.Fields("factor3")
factor4 = "" & mytablex.Fields("factor4")
factor5 = "" & mytablex.Fields("factor5")
factor6 = "" & mytablex.Fields("factor6")
factor7 = "" & mytablex.Fields("factor7")
factor8 = "" & mytablex.Fields("factor8")
factor9 = "" & mytablex.Fields("factor9")
factor10 = "" & mytablex.Fields("factor10")
pventa1 = "" & mytablex.Fields("pventa1")
pventa2 = "" & mytablex.Fields("pventa2")
pventa3 = "" & mytablex.Fields("pventa3")
pventa4 = "" & mytablex.Fields("pventa4")
pventa5 = "" & mytablex.Fields("pventa5")
pventa6 = "" & mytablex.Fields("pventa6")
pventa7 = "" & mytablex.Fields("pventa7")
pventa8 = "" & mytablex.Fields("pventa8")
pventa9 = "" & mytablex.Fields("pventa9")
pventa10 = "" & mytablex.Fields("pventa10")
margen1 = "" & mytablex.Fields("margen1")
margen2 = "" & mytablex.Fields("margen2")
margen3 = "" & mytablex.Fields("margen3")
margen4 = "" & mytablex.Fields("margen4")
margen5 = "" & mytablex.Fields("margen5")
margen6 = "" & mytablex.Fields("margen6")
margen7 = "" & mytablex.Fields("margen7")
margen8 = "" & mytablex.Fields("margen8")
margen9 = "" & mytablex.Fields("margen9")
margen10 = "" & mytablex.Fields("margen10")
minimo11 = "" & mytablex.Fields("minimo11")
minimo12 = "" & mytablex.Fields("minimo12")
minimo13 = "" & mytablex.Fields("minimo13")
minimo14 = "" & mytablex.Fields("minimo14")
minimo15 = "" & mytablex.Fields("minimo15")
maximo11 = "" & mytablex.Fields("maximo11")
maximo12 = "" & mytablex.Fields("maximo12")
maximo13 = "" & mytablex.Fields("maximo13")
maximo14 = "" & mytablex.Fields("maximo14")
maximo15 = "" & mytablex.Fields("maximo15")
pventa11 = "" & mytablex.Fields("pventa11")
pventa12 = "" & mytablex.Fields("pventa12")
pventa13 = "" & mytablex.Fields("pventa13")
pventa14 = "" & mytablex.Fields("pventa14")
pventa15 = "" & mytablex.Fields("pventa15")
margen11 = "" & mytablex.Fields("margen11")
margen12 = "" & mytablex.Fields("margen12")
margen13 = "" & mytablex.Fields("margen13")
margen14 = "" & mytablex.Fields("margen14")
margen15 = "" & mytablex.Fields("margen15")
fechai11 = "" & mytablex.Fields("fechai11")
fechaf11 = "" & mytablex.Fields("fechaf11")
fechaid = "" & mytablex.Fields("fechaid")
fechafd = "" & mytablex.Fields("fechafd")
dscto = "" & mytablex.Fields("dscto")
ccosto = "" & mytablex.Fields("ccosto")
If Len(unidad1) = 0 Then
   unidad1 = "UND"
End If
If Len(factor1) = 0 Then
   factor1 = "1"
End If
End Sub
Sub inicializa_precios()
unidad1 = "UND"
unidad2 = ""
unidad3 = ""
unidad4 = ""
unidad5 = ""
unidad6 = ""
unidad7 = ""
unidad8 = ""
unidad9 = ""
unidad10 = ""
'saldoini = ""

factor1 = "1"
factor2 = ""
factor3 = ""
factor4 = ""
factor5 = ""
factor6 = ""
factor7 = ""
factor8 = ""
factor9 = ""
factor10 = ""

pventa1 = ""
pventa2 = ""
pventa3 = ""
pventa4 = ""
pventa5 = ""
pventa6 = ""
pventa7 = ""
pventa8 = ""
pventa9 = ""
pventa10 = ""

margen1 = ""
margen2 = ""
margen3 = ""
margen4 = ""
margen5 = ""
margen6 = ""
margen7 = ""
margen8 = ""
margen9 = ""
margen10 = ""
minimo11 = ""
minimo12 = ""
minimo13 = ""
minimo14 = ""
minimo15 = ""

maximo11 = ""
maximo12 = ""
maximo13 = ""
maximo14 = ""
maximo15 = ""

pventa11 = ""
pventa12 = ""
pventa13 = ""
pventa14 = ""
pventa15 = ""
margen11 = ""
margen12 = ""
margen13 = ""
margen14 = ""
margen15 = ""
fechai11 = ""
fechaf11 = ""
fechaid = ""
fechafd = ""
dscto = ""
ccosto = ""
End Sub
Sub hacer_barras()
Dim X As Integer, Y As Integer, z As Integer, pos As Integer
Dim temp As String
Dim Codevalue As Integer
Dim BarCode As String
Call equivalentvalue
    
Picture1.Cls
pos = 10
BarCode = UCase(barras.Text)

For X = 1 To Len(BarCode)
    temp = Mid$(BarCode, X, 1)
    Select Case temp
    Case "0" To "9"
        Codevalue = Val(temp)
    Case "A" To "Z"
        Codevalue = Asc(temp) - 55
    Case "-"
        Codevalue = 36
    Case "."
        Codevalue = 37
    Case " "
        Codevalue = 38
    Case "$"
        Codevalue = 39
    Case "/"
        Codevalue = 40
    Case "+"
        Codevalue = 41
    Case "%"
        Codevalue = 42
    Case "*"
        Codevalue = 43
    Case Else
        Picture1.Cls
        Picture1.Print temp & " is not valid"
        Exit Sub
    End Select
    
    For Y = 1 To 9
        If Y / 2 = Int(Y / 2) Then
            pos = pos + 1 + (3 * Val(Mid$(ArrBarCode(Codevalue), Y, 1)))
        Else
            For z = 1 To 1 + (3 * Val(Mid$(ArrBarCode(Codevalue), Y, 1)))
                Picture1.Line (pos, 1)-(pos, 50)
                pos = pos + 1
            Next z
        End If
    Next Y
    pos = pos + 1
Next X

    Picture1.CurrentX = Len(BarCode) * 7
    Picture1.Print BarCode
End Sub
Private Sub equivalentvalue()
    ArrBarCode(0) = "000110100"
    ArrBarCode(1) = "100100001"
    ArrBarCode(2) = "001100001"
    ArrBarCode(3) = "101100000"
    ArrBarCode(4) = "000110001"
    ArrBarCode(5) = "100110000"
    ArrBarCode(6) = "001110000"
    ArrBarCode(7) = "000100101"
    ArrBarCode(8) = "100100100"
    ArrBarCode(9) = "001100100"
    ArrBarCode(10) = "100001001"
    ArrBarCode(11) = "001001001"
    ArrBarCode(12) = "101001000"
    ArrBarCode(13) = "000011001"
    ArrBarCode(14) = "100011000"
    ArrBarCode(15) = "001011000"
    ArrBarCode(16) = "000001101"
    ArrBarCode(17) = "100001100"
    ArrBarCode(18) = "001001100"
    ArrBarCode(19) = "000011100"
    ArrBarCode(20) = "100000011"
    ArrBarCode(21) = "001000011"
    ArrBarCode(22) = "101000010"
    ArrBarCode(23) = "000010011"
    ArrBarCode(24) = "100010010"
    ArrBarCode(25) = "001010010"
    ArrBarCode(26) = "000000111"
    ArrBarCode(27) = "100000110"
    ArrBarCode(28) = "001000110"
    ArrBarCode(29) = "000010110"
    ArrBarCode(30) = "110000001"
    ArrBarCode(31) = "011000001"
    ArrBarCode(32) = "111000000"
    ArrBarCode(33) = "010010001"
    ArrBarCode(34) = "110010000"
    ArrBarCode(35) = "011010000"
    ArrBarCode(36) = "010000101"
    ArrBarCode(37) = "110000100"
    ArrBarCode(38) = "011000100"
    ArrBarCode(39) = "010101000"
    ArrBarCode(40) = "010100010"
    ArrBarCode(41) = "010001010"
    ArrBarCode(42) = "000101010"
    ArrBarCode(43) = "010010100"
End Sub


