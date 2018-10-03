VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tinicont 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Saldos Iniciales"
   ClientHeight    =   10710
   ClientLeft      =   150
   ClientTop       =   0
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Generar"
      Height          =   5175
      Left            =   3240
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   7815
      Begin VB.TextBox geperiodof 
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
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   42
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox geperiodoi 
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
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   36
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "generar"
         Height          =   615
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   615
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final (Ultima Transaccion)"
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
         TabIndex        =   43
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label nregistrop 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1440
         TabIndex        =   41
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Procesado"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label nregistro 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1440
         TabIndex        =   39
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro registros"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label geperiodog 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Left            =   4200
         TabIndex        =   37
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label gebodega 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Left            =   4200
         TabIndex        =   33
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label gelocal 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Left            =   4200
         TabIndex        =   32
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
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
         TabIndex        =   31
         Top             =   2400
         Width           =   3735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
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
         TabIndex        =   30
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo Inicial a ser generado"
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
         TabIndex        =   29
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicial (Inventario Inicial)"
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
         TabIndex        =   28
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame2"
      Height          =   6615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   12495
      Begin VB.TextBox periodof 
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
         MaxLength       =   10
         TabIndex        =   46
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox periodoi 
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
         MaxLength       =   10
         TabIndex        =   44
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox nbodega 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox nlocal 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox periodo 
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
         MaxLength       =   10
         TabIndex        =   22
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox bodega 
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
         TabIndex        =   20
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox local1 
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
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox id 
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
         Left            =   2280
         MaxLength       =   6
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
         MaxLength       =   30
         TabIndex        =   14
         Top             =   3720
         Width           =   6015
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   8400
         Picture         =   "tinicont.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir todo"
         Top             =   3120
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   8400
         Picture         =   "tinicont.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2160
         Width           =   1470
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         Caption         =   "dd/mm/yyyy"
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
         Left            =   4320
         TabIndex        =   50
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF00&
         Caption         =   "dd/mm/yyyy"
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
         Left            =   4320
         TabIndex        =   49
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rango Transaccion"
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
         TabIndex        =   48
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Periodo Final"
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
         TabIndex        =   47
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Periodo Inicial"
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
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         Caption         =   "dd/mm/yyyy"
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
         Left            =   4320
         TabIndex        =   26
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inv.Inicial"
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
         TabIndex        =   23
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bodega"
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
         TabIndex        =   21
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
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
         TabIndex        =   19
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Id"
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
         BackColor       =   &H00FFFF00&
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
         Top             =   3720
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
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
         Picture         =   "tinicont.frx":1194
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
         Picture         =   "tinicont.frx":23A6
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
         Picture         =   "tinicont.frx":35B8
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
         Picture         =   "tinicont.frx":47CA
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
         Picture         =   "tinicont.frx":59DC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Descripcio"
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
            DataField       =   "Id"
            Caption         =   "Id"
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
            DataField       =   "Local"
            Caption         =   "Local"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Bodega"
            Caption         =   "Bodega"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Periodo"
            Caption         =   "Periodo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
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
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
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
   Begin VB.Menu fdk8943 
      Caption         =   "&Generar"
   End
   Begin VB.Menu dfl8n8 
      Caption         =   "&IngresarSaldoInicial"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tinicont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TXEMPREINI As New ADODB.Recordset
Private Sub ajdu1_Click()
If Frame3.Visible = True Then Exit Sub
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If
inicializa
Frame2.Visible = True
Frame2.Caption = "Nuevo"
cmdGuardar.Enabled = True
habilita 1
id.Enabled = False
id = ""
local1.Enabled = True
bodega.Enabled = True
periodo.Enabled = True
nlocal.Enabled = True
nbodega.Enabled = True
local1.SetFocus
End Sub

Private Sub bo712_Click()
Dim buf As String
On Error GoTo cmd656_err
If Frame3.Visible = True Then Exit Sub
buf = Trim("" & TXEMPREINI.Fields("id"))
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If
If MsgBox("Desea Borra " + "" & TXEMPREINI.Fields("id"), 1, "Aviso,Se va a borrar tambien los saldos iniciales") <> 1 Then
   Exit Sub
End If
cn.Execute ("delete from saldoini where local='" & "" & TXEMPREINI.Fields("local") & "' and bodega='" & "" & TXEMPREINI.Fields("bodega") & "' and fecha='" & "" & TXEMPREINI.Fields("periodo") & "'")
TXEMPREINI.Delete
Command1_Click



Exit Sub
cmd656_err:
MsgBox "Seleccione un dato " + error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub bodega_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
periodo.SetFocus

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
found = grabar()
End Sub

Private Sub cmdPrint_Click()
'djuer1_Click
End Sub

Private Sub cmdSave_Click()
f8443_Click
End Sub


Private Sub Command3_Click()
If Len(geperiodoi) <> 10 Then
   geperiodoi.SetFocus
   Exit Sub
End If
If Not IsDate(geperiodoi) Then
   geperiodoi.SetFocus
   Exit Sub
End If
If Len(geperiodof) <> 10 Then
   geperiodof.SetFocus
   Exit Sub
End If
If Not IsDate(geperiodof) Then
   geperiodof.SetFocus
   Exit Sub
End If
proceso_sunat
End Sub

Private Sub dfl8n8_Click()
On Error GoTo cmd89121_err
If Frame3.Visible = True Then Exit Sub
tsaldoin.local1 = "" & TXEMPREINI.Fields("local")
tsaldoin.bodega = "" & TXEMPREINI.Fields("bodega")
tsaldoin.fecha = "" & TXEMPREINI.Fields("periodo")
tsaldoin.Show 1
Exit Sub
cmd89121_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub dk9893_Click()
If Frame3.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
reporgen.NAMETABLA = "saldoinicontrol"
reporgen.Show 1

End Sub
Sub prueba_reporte()
'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\idesproducto.rpt", "")
End Sub

Private Sub fdk8943_Click()
On Error GoTo cmdf1_err
geperiodog = "" & TXEMPREINI.Fields("periodo")
gelocal = "" & TXEMPREINI.Fields("local")
gebodega = "" & TXEMPREINI.Fields("bodega")
geperiodoi = "" & TXEMPREINI.Fields("periodoi")
geperiodof = "" & TXEMPREINI.Fields("periodof")
Frame3.Visible = True
Command3.Visible = True
'geperiodoi.SetFocus
Exit Sub
cmdf1_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub id_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
If Len(id) = 0 Then Exit Sub
descripcio.SetFocus
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
      cad = "SELECT * from saldoinicontrol  order by periodo "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT *  from saldoinicontrol   where  " & Combo1 & " like '" & buffer & "%' order by periodo"
   End If
   
   If TXEMPREINI.State = 1 Then TXEMPREINI.Close
   TXEMPREINI.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbGrid1.DataSource = TXEMPREINI
   dbGrid1.columns(0).Width = 4000
   dbGrid1.columns(1).Width = 2000
   If TXEMPREINI.RecordCount > 0 Then
     dbGrid1.SetFocus
  End If
End If
End Sub

Private Sub Command2_Click()
Frame3.Visible = False
End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   'id = dbGrid1.Columns(1)
   'Frame1.Visible = False
   'Frame1.Enabled = False
   'id.SetFocus
   'id_KeyPress 13
End If
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
Dim buf As String
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
If Frame3.Visible = True Then
   Frame3.Visible = False
   Exit Sub
End If
If Frame2.Visible = True Then
   habilita 0
   Frame2.Visible = False
   dbGrid1.Enabled = True
   Exit Sub
End If
tinicont.Hide
Unload tinicont
End Sub


Private Sub f8443_Click()
Dim buf As String
On Error GoTo cmd456_err
If Frame3.Visible = True Then Exit Sub
buf = "" & TXEMPREINI.Fields("id")
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
id.Enabled = False
local1.Enabled = False
bodega.Enabled = False
periodo.Enabled = False
nlocal.Enabled = False
nbodega.Enabled = False
descripcio.SetFocus
Exit Sub
cmd456_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub fjh433_Click()
Dim buf As String
On Error GoTo cmd556_err
If Frame3.Visible = True Then Exit Sub
buf = "" & TXEMPREINI.Fields("id")
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
id.Enabled = False
descripcio.SetFocus
Exit Sub
cmd556_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub Form_Activate()
'agregar_menus
nlocal.ListIndex = 0

nbodega.ListIndex = 0

Command1_Click
End Sub

Private Sub Form_Load()
Dim mytablex As New ADODB.Recordset
Combo1.Clear
Combo1.AddItem "local"
Combo1.ListIndex = 0

nlocal.Clear
nlocal.AddItem "%"
mytablex.Open "select * from tlocal", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
nlocal.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
mytablex.MoveNext
Loop
mytablex.Close

nbodega.Clear
nbodega.AddItem "%"
mytablex.Open "select * from bodega", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
nbodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
mytablex.MoveNext
Loop
mytablex.Close

End Sub
Sub inicializa()
descripcio = ""
local1 = ""
bodega = ""
periodo = ""
periodoi = ""
periodof = ""
End Sub
Sub pone_registro()
id = Trim("" & TXEMPREINI.Fields("id"))
local1 = Trim("" & TXEMPREINI.Fields("local"))
bodega = Trim("" & TXEMPREINI.Fields("bodega"))
periodo = Trim("" & TXEMPREINI.Fields("periodo"))
descripcio = Trim("" & TXEMPREINI.Fields("descripcio"))
periodoi = Trim("" & TXEMPREINI.Fields("periodoi"))
periodof = Trim("" & TXEMPREINI.Fields("periodof"))
End Sub
Sub grabando()
TXEMPREINI.Fields("local") = Trim(local1)
TXEMPREINI.Fields("bodega") = Trim(bodega)
TXEMPREINI.Fields("periodo") = Trim(periodo)
TXEMPREINI.Fields("periodoi") = Trim(periodoi)
TXEMPREINI.Fields("periodof") = Trim(periodof)
TXEMPREINI.Fields("descripcio") = Trim(descripcio)
End Sub

Private Sub grba1_Click()

End Sub

Function grabar()
Dim found As Integer
Dim rbusca As New ADODB.Recordset
found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If
If Frame2.Caption = "Nuevo" Then
   'If Len(id) = 0 Then
   '   id.SetFocus
   '   Exit Function
   'End If
   rbusca.Open "select id from saldoinicontrol where local='" & Trim(local1) & "' and bodega='" & Trim(bodega) & "' AND periodo='" & periodo & "'", cn, adOpenStatic, adLockOptimistic
   If rbusca.RecordCount > 0 Then
      rbusca.Close
      MsgBox "Ya existe id ", 48, "Aviso"
      Exit Function
   End If
   TXEMPREINI.AddNew
   'TXEMPREINI.Fields("id") = ID
   grabando
   TXEMPREINI.Update
   dlo132_Click
   Exit Function
End If
If Frame2.Caption = "Modifica" Then
   'TXEMPREINI.Fields("id") = ID
   grabando
   TXEMPREINI.Update
   dlo132_Click
   Exit Function
End If

End Function

Function valida()
If Len(local1) = 0 Then
   local1.SetFocus
   Exit Function
End If
If Len(bodega) = 0 Then
   bodega.SetFocus
   Exit Function
End If
If Len(periodo) <> 10 Then
   periodo = ""
   periodo.SetFocus
   Exit Function
End If
If Not IsDate(periodo) Then
   periodo = ""
   periodo.SetFocus
   Exit Function
End If

If Len(periodoi) <> 10 Then
   periodoi = ""
   periodoi.SetFocus
   Exit Function
End If
If Not IsDate(periodoi) Then
   periodoi = ""
   periodoi.SetFocus
   Exit Function
End If

If Len(periodof) <> 10 Then
   periodof = ""
   periodof.SetFocus
   Exit Function
End If
If Not IsDate(periodof) Then
   periodof = ""
   periodof.SetFocus
   Exit Function
End If


If Len(descripcio) = 0 Then
   descripcio.SetFocus
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
            dfl8n8.Enabled = True
            fdk8943.Enabled = True
            'dfj400.Enabled = True
            
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
dfl8n8.Enabled = False
fdk8943.Enabled = False
'dfj400.Enabled = False
           
End If

      
End Sub
Sub agregar_menus()
Dim i As Integer
For i = 1 To mnuArchivoArray.count - 1
    Unload mnuArchivoArray(i)
Next
     
Dim mytablex As New ADODB.Recordset
   mytablex.Open "select * from archivo where menu='id' and   estado='S'", cn, adOpenStatic, adLockOptimistic
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

Private Sub local1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
bodega.SetFocus

End Sub

Sub mnuarchivoarray_click(Index As Integer)
Dim mytablex As New ADODB.Recordset
Dim buf As String
buf = mnuArchivoArray(Index).Caption
   mytablex.Open "select * from archivo where menu='id' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
   End If
   'busca el reporte
   buf = mytablex.Fields("archivo")
   mytablex.Close
   'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub

Private Sub nbodega_Click()
If nbodega <> "%" Then
   bodega = extra_loquesea(nbodega)
End If
End Sub

Private Sub nlocal_Click()
If nlocal <> "%" Then
   local1 = extra_loquesea(nlocal)
End If
End Sub

Private Sub periodo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
periodoi.SetFocus

End Sub
Sub proceso_sunat()
Dim sdx As Double
Dim vr
Dim saldoini As Double
Dim costo_ultimo As Double
Dim buf As String
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim mytablez As New ADODB.Recordset
mytablex.Open "Select * from producto", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount = 0 Then Exit Sub
mytablex.MoveFirst
nregistro = "" & mytablex.RecordCount
cn.Execute ("delete from saldoini where local='" & Trim(gelocal) & "' and bodega='" & Trim(gebodega) & "' and fecha='" & geperiodog & "'")
mytabley.Open "Select * from saldoini where local='" & Trim(gelocal) & "' and bodega='" & Trim(gebodega) & "' and fecha='" & geperiodog & "'", cn, adOpenStatic, adLockOptimistic
sdx = 0
Do
If mytablex.EOF Then Exit Do
   vr = DoEvents()
   sdx = sdx + 1
   nregistrop = "" & sdx
   mytabley.AddNew
   mytabley.Fields("fecha") = Format(geperiodog, "dd/mm/yyyy")
   mytabley.Fields("bodega") = Trim(gebodega)
   mytabley.Fields("local") = Trim(gelocal)
   mytabley.Fields("producto") = Trim("" & mytablex.Fields("producto"))
   mytabley.Fields("descripcio") = Trim("" & mytablex.Fields("descripcio"))
   mytabley.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
   mytabley.Fields("factor") = Val("" & mytablex.Fields("factor"))
   mytabley.Fields("cantidad") = 0
   mytabley.Fields("saldoant") = 0
   mytabley.Fields("cantidad1") = 0
   mytabley.Fields("linea") = Trim("" & mytablex.Fields("linea"))
   mytabley.Fields("familia") = Trim("" & mytablex.Fields("familia"))
   mytabley.Fields("costo") = Val("" & mytablex.Fields("costou"))

   saldoini = 0
   costo_ultimo = 0
   'ahora vamos a buscar sus transacciones
   mytablez.Open "Select * from saldoini where local='" & Trim(gelocal) & "' and bodega='" & Trim(gebodega) & "' and fecha='" & geperiodoi & "' and producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic
   If mytablez.RecordCount > 0 Then
      saldoini = Val("" & mytablez.Fields("cantidad1")) * Val("" & mytablez.Fields("factor"))
      costo_ultimo = Val("" & mytablez.Fields("costo"))
      'MsgBox saldoini
   'MsgBox saldoini
   End If
   mytablez.Close
   
   'el saldoini
   
   buf = "select * from detalle where "
   buf = buf & "  fecha>='" & Format(geperiodoi, "YYYYMMDD") & "'"
   buf = buf & " and fecha<='" & Format(geperiodof, "YYYYMMDD") & "' "
   buf = buf & " and producto='" & "" & mytablex.Fields("producto") & "'"
   buf = buf & " and local='" & Trim(gelocal) & "'"
   buf = buf & " and estado='2'"
   buf = buf & " and bodega='" & Trim(gebodega) & "'  order by fecha"
   mytablez.Open buf, cn, adOpenStatic, adLockOptimistic
   sdx = 0
   Do
     If mytablez.EOF Then
        'mytablez.Close
        GoTo queso
     End If
     '------------------------------
      Select Case "" & mytablez.Fields("acu")
          Case "A", "B", "C", "D", "G", "T"
          saldoini = saldoini - Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
          Case "J", "K", "L", "M", "P", "S"
          saldoini = saldoini + Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
          costo_ultimo = Val("" & mytablez.Fields("precio"))
     End Select
     mytablez.MoveNext
   Loop
queso:
   mytablez.Close
   mytabley.Fields("cantidad1") = saldoini
   mytabley.Fields("costo") = costo_ultimo
   
   mytabley.Update
   '-------- bucle de producto-----------------
   mytablex.MoveNext
   Loop

   mytablex.Close
   mytabley.Close
      
End Sub

Private Sub periodof_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
descripcio.SetFocus
End Sub

Private Sub periodoi_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
periodof.SetFocus
End Sub
