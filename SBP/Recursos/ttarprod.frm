VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ttarprod 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Tarjetas de Control"
   ClientHeight    =   10065
   ClientLeft      =   165
   ClientTop       =   15
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   15135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consultas"
      Height          =   9615
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   14775
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   8655
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   15266
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   25
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
      Begin VB.TextBox bproducto 
         Height          =   375
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   35
         Top             =   240
         Width           =   2895
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12480
         TabIndex        =   39
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acepta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12480
         TabIndex        =   37
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ejecutar Busqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   36
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame2"
      Height          =   9735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   15015
      Begin VB.TextBox fechaf 
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
         TabIndex        =   47
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox fechai 
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
         TabIndex        =   45
         Top             =   4200
         Width           =   1935
      End
      Begin VB.TextBox estado 
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
         TabIndex        =   40
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox planonumero 
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
         TabIndex        =   31
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Seccion 
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
         TabIndex        =   29
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox costo 
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
         TabIndex        =   26
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox merma 
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
         TabIndex        =   24
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox cantidad 
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
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox factor 
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
         MaxLength       =   4
         TabIndex        =   20
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox unidad 
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
         Top             =   1680
         Width           =   975
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
         MaxLength       =   80
         TabIndex        =   16
         Top             =   1320
         Width           =   6975
      End
      Begin VB.TextBox tarjeta 
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
         TabIndex        =   13
         Top             =   240
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
         TabIndex        =   12
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10680
         Picture         =   "ttarprod.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir todo"
         Top             =   1440
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10680
         Picture         =   "ttarprod.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaF"
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
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fechai"
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
         TabIndex        =   46
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
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
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero Orden Prod."
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
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label seccion11 
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
         Left            =   120
         TabIndex        =   30
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Costo"
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
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Merma"
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
         TabIndex        =   25
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   23
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor"
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
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad"
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
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Descripcio1 
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
         TabIndex        =   17
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tarjeta"
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
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Producto1 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   14
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   15075
      TabIndex        =   1
      Top             =   0
      Width           =   15135
      Begin VB.TextBox xtarjeta 
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
         Left            =   6360
         MaxLength       =   11
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "%"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox xfechaf 
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
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox xfechai 
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
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
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
         Picture         =   "ttarprod.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   1695
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
         Left            =   8160
         TabIndex        =   6
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
         Picture         =   "ttarprod.frx":23A6
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "ttarprod.frx":35B8
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "ttarprod.frx":47CA
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "ttarprod.frx":59DC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tarjeta"
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
         Left            =   3600
         TabIndex        =   52
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fechaf"
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
         Left            =   3600
         TabIndex        =   50
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fechai"
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
         Left            =   3600
         TabIndex        =   44
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orden Produccion Nro"
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
         Left            =   3600
         TabIndex        =   42
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   15015
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   0
         TabIndex        =   28
         Top             =   120
         Width           =   14895
         _ExtentX        =   26273
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "Numero"
            Caption         =   "Numero"
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
            DataField       =   "Tarjeta"
            Caption         =   "Tarjeta"
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
            DataField       =   "Producto"
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "Unidad"
            Caption         =   "Und"
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
            DataField       =   "Factor"
            Caption         =   "Factor"
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
            DataField       =   "Cantidad"
            Caption         =   "Cant"
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
            DataField       =   "Merma"
            Caption         =   "Merma"
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
            DataField       =   "Costo"
            Caption         =   "Costo"
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
            DataField       =   "CostoTotal"
            Caption         =   "CostoTotal"
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
         BeginProperty Column10 
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
         BeginProperty Column11 
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
         BeginProperty Column12 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   4305.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
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
Attribute VB_Name = "ttarprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txplanop As New ADODB.Recordset
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
planonumero.Enabled = True
tarjeta.Enabled = True
tarjeta = ""
tarjeta.SetFocus
End Sub

Private Sub bo712_Click()
Dim buf As String
On Error GoTo cmd656_err
buf = "" & txplanop.Fields("tarjeta")
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If
If MsgBox("Desea Borra " + ("" & txplanop.Fields("tarjeta")), 1, "Aviso") <> 1 Then
   Exit Sub
End If
txplanop.Delete
Command1_Click



Exit Sub
cmd656_err:
MsgBox "Seleccione un dato " + error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub bproducto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Label11_Click
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


Private Sub DBGrid2_DblClick()
Label13_Click
End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Label13_Click
End If
End Sub

Private Sub dk9893_Click()
If Frame2.Visible = True Then Exit Sub
reporgen.NAMETABLA = "tarjetaproduccion"
reporgen.Show 1

End Sub
Sub prueba_reporte()
'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\tarjetaproduccionesproducto.rpt", "")
End Sub

'Private Sub tarjetaproduccion_KeyPress(KeyAscii As Integer)
'Dim found As Integer
'If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
'If KeyAscii = 27 Then
'   dlo132_Click
'   Exit Sub
'End If
'If Len(tarjetaproduccion) = 0 Then Exit Sub
'descripcio.SetFocus
'End Sub


Private Sub Command1_Click()
Frame1.Visible = True
Frame1.Enabled = True

If Not IsDate(xfechai) Then Exit Sub
If Not IsDate(xfechaf) Then Exit Sub
ejecuta 1
End Sub
Sub ejecuta(sw As Integer)
Dim cad As String
   'buffer = "" & Val("" & buffer)
  cad = "SELECT *  from tarjetaproduccion where  "
  cad = cad & "  fechai>='" & Format(xfechai, "YYYYMMDD") & "'"
  cad = cad & " and fechaf<='" & Format(xfechaf, "YYYYMMDD") & "' "
  If IsNumeric(buffer) Then
     cad = cad & " and numero=" & Val(buffer)
  End If
  If xtarjeta <> "%" Then
     cad = cad & " and tarjeta like " & xtarjeta & "'"
  End If
   cad = cad & " order by numero"
   'MsgBox cad
   If txplanop.State = 1 Then txplanop.Close
   txplanop.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbGrid1.DataSource = txplanop
   If txplanop.RecordCount > 0 Then
     dbGrid1.SetFocus
  End If

End Sub

Private Sub Command2_Click()

End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = &H71 Then
   txplanop.Fields("estado") = "1"
   totaliza_secciones
   txplanop.Update
   txplanop.Requery
End If

If KeyCode = 13 Then
   'tarjetaproduccion = dbGrid1.Columns(1)
   'Frame1.Visible = False
   'Frame1.Enabled = False
   'tarjetaproduccion.SetFocus
   'tarjetaproduccion_KeyPress 13
End If
End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
Dim buf As String
Dim buf2 As String
If KeyAscii <> 13 And KeyAscii <> 27 Then
         If KeyAscii = 8 Then
            If Len(bproducto) > 0 Then
               buf = Mid$(bproducto, 1, Len(bproducto) - 1)
               bproducto = buf
               KeyAscii = 0
               Else
               KeyAscii = 0
               Exit Sub
            End If
         End If
         buf = Chr(KeyAscii)
         If Chr(KeyAscii) = "*" Then
            buf = ""
            bproducto = buf
         End If
         If KeyAscii <> 13 Then
            bproducto = bproducto + buf
         End If
         buf = bproducto
         ejecuta 0
         
End If
End Sub


Private Sub dlo132_Click()
If Frame2.Visible = True Then
   habilita 0
   Frame2.Visible = False
   dbGrid1.Enabled = True
   Exit Sub
End If
ttarprod.Hide
Unload ttarprod
End Sub


Private Sub f8443_Click()
Dim buf As String
On Error GoTo cmd456_err
buf = "" & txplanop.Fields("tarjeta")
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
planonumero.Enabled = False
tarjeta.Enabled = False
producto.SetFocus
Exit Sub
cmd456_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub fjh433_Click()
Dim buf As String
On Error GoTo cmd556_err
buf = "" & txplanop.Fields("tarjeta")
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
tarjeta.Enabled = False
producto.SetFocus
Exit Sub
cmd556_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub Form_Activate()
'agregar_menus
'rfechai = Format(Now, "dd/mm/yyyy")
'rfechaf = Format(Now, "dd/mm/yyyy")
Command1_Click
End Sub

Private Sub Form_Load()
xfechai = Format(Now, "dd/mm/yyyy")
xfechaf = Format(Now, "dd/mm/yyyy")
End Sub
Sub inicializa()
fechai = Format(Now, "dd/mm/yyyy")
fechaf = Format(Now, "dd/mm/yyyy")
planonumero = ""
producto = ""
descripcio = ""
cantidad = ""
factor = ""
unidad = ""
costo = ""
seccion = ""
merma = ""
tarjeta = ""
estado = "0"

End Sub
Sub pone_registro()
planonumero = Trim("" & txplanop.Fields("numero"))
tarjeta = Trim("" & txplanop.Fields("tarjeta"))
producto = Trim("" & txplanop.Fields("producto"))
descripcio = Trim("" & txplanop.Fields("descripcio"))
unidad = Trim("" & txplanop.Fields("unidad"))
factor = Trim("" & txplanop.Fields("factor"))
cantidad = Trim("" & txplanop.Fields("merma"))
merma = Trim("" & txplanop.Fields("merma"))
costo = Trim("" & txplanop.Fields("costo"))
seccion = Trim("" & txplanop.Fields("seccion"))
cantidad = Trim("" & txplanop.Fields("cantidad"))
estado = Trim("" & txplanop.Fields("estado"))
fechai = Trim("" & txplanop.Fields("fechai"))
fechaf = Trim("" & txplanop.Fields("fechaf"))
End Sub
Sub grabando()
txplanop.Fields("estado") = Val(estado)
txplanop.Fields("numero") = Val(planonumero)
txplanop.Fields("tarjeta") = Trim(tarjeta)
txplanop.Fields("producto") = Trim(producto)
txplanop.Fields("descripcio") = Trim(descripcio)
txplanop.Fields("unidad") = Trim(unidad)
txplanop.Fields("factor") = Val(factor)
txplanop.Fields("cantidad") = Val(cantidad)
txplanop.Fields("costo") = Val(costo)
txplanop.Fields("merma") = Val(merma)
txplanop.Fields("seccion") = Trim(seccion)
txplanop.Fields("fechai") = Trim(fechai)
txplanop.Fields("fechaf") = Trim(fechaf)
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
   'If Len(tarjetaproduccion) = 0 Then
   '   tarjetaproduccion.SetFocus
   '   Exit Function
   'End If
   rbusca.Open "select * from tarjetaproduccion where tarjeta='" & tarjeta & "'", cn, adOpenStatic, adLockOptimistic
   If rbusca.RecordCount > 0 Then
      rbusca.Close
      MsgBox "Ya existe tarjetaproduccion ", 48, "Aviso"
      Exit Function
   End If
   txplanop.AddNew
   'txplanop.Fields("numero") = tarjetaproduccion
   grabando
   txplanop.Update
   dlo132_Click
   Exit Function
End If
If Frame2.Caption = "Modifica" Then
   'txplanop.Fields("tarjetaproduccion") = tarjetaproduccion
   grabando
   txplanop.Update
   dlo132_Click
   Exit Function
End If

End Function

Function valida()
If Not IsNumeric(planonumero) Then
   planonumero.SetFocus
  Exit Function
End If
If Len(tarjeta) = 0 Then
   tarjeta.SetFocus
   Exit Function
End If
If Len(planonumero) = 0 Then
   planonumero.SetFocus
   Exit Function
End If
If existe_numero() <> 1 Then
   planonumero.SetFocus
   Exit Function
End If

If Len(producto) = 0 Then
   producto.SetFocus
   Exit Function
End If
If Len(descripcio) = 0 Then
   descripcio.SetFocus
   Exit Function
End If
If Len(unidad) = 0 Then
   unidad.SetFocus
   Exit Function
End If
If Val(factor) <= 0 Then
   factor.SetFocus
   Exit Function
End If
If Val(cantidad) <= 0 Then
   cantidad.SetFocus
   Exit Function
End If
If Val(merma) < 0 Then
   merma.SetFocus
   Exit Function
End If
If Val(costo) < 0 Then
   costo.SetFocus
   Exit Function
End If
If Not IsDate(fechai) Then
   fechai.SetFocus
   Exit Function
End If
If Not IsDate(fechaf) Then
   fechaf.SetFocus
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
Sub agregar_menus()
Dim i As Integer
For i = 1 To mnuArchivoArray.count - 1
    Unload mnuArchivoArray(i)
Next
     
Dim mytablex As New ADODB.Recordset
   mytablex.Open "select * from archivo where menu='tarjetaproduccion' and   estado='S'", cn, adOpenStatic, adLockOptimistic
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

Private Sub Label17_Click()
Frame3.Visible = True
End Sub

Private Sub Label11_Click()
Dim buf As String
Dim mytablex As New ADODB.Recordset

Set mytablex = Nothing
dbgrid2.refresh
If opcion1 = "1" Then
buf = "select Descripcio,Producto,Unidad,Factor,Costou as Costo from producto where "
If Combo2 = "Descripcio" Then
   buf = buf & " descripcio like '" & bproducto & "%'" '
End If
If Combo2 = "Producto" Then
   buf = " Producto like '" & bproducto & "%'" '
End If
End If

If opcion1 = "2" Then
buf = "select Descripcio,seccion from pseccion where "
buf = buf & " descripcio like '" & bproducto & "%'" '
End If
If opcion1 = "3" Then
buf = "select Numero,Fechai,Fecha,Estado from cproducc where "
buf = buf & " vendedor like '" & bproducto & "%'" '
End If

mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
Set dbgrid2.DataSource = mytablex
   dbgrid2.columns(0).Width = 4000
   dbgrid2.columns(1).Width = 2000
   
   dbgrid2.refresh
End Sub

Private Sub Label13_Click()
On Error GoTo cmd89123_err
If opcion1 = "1" Then
producto = "" & dbgrid2.columns("producto")
descripcio = "" & dbgrid2.columns("descripcio")
unidad = "" & dbgrid2.columns("unidad")
factor = "" & dbgrid2.columns("factor")
costo = "" & dbgrid2.columns("costo")
cantidad.SetFocus
End If
If opcion1 = "2" Then
seccion = "" & dbgrid2.columns("seccion")
seccion.SetFocus
End If
If opcion1 = "3" Then
planonumero = "" & dbgrid2.columns("numero")
seccion.SetFocus
End If

Frame3.Visible = False
Exit Sub
cmd89123_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub Label2_Click()
Combo2.Clear
Combo2.AddItem "Vendedor"
Combo2.ListIndex = 0
opcion1 = "3"
bproducto = ""
Frame3.Visible = True
bproducto.SetFocus
Label11_Click
End Sub

Private Sub Label9_Click()
Frame3.Visible = False
End Sub

Sub mnuarchivoarray_click(Index As Integer)
Dim mytablex As New ADODB.Recordset
Dim buf As String
buf = mnuArchivoArray(Index).Caption
   mytablex.Open "select * from archivo where menu='tarjetaproduccion' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
   End If
   'busca el reporte
   buf = mytablex.Fields("archivo")
   mytablex.Close
   'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub

Function existe_numero()
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from cproducc where numero='" & planonumero & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
  existe_numero = 1
End If
mytablex.Close
End Function

Private Sub Producto1_Click()
Combo2.Clear
Combo2.AddItem "Descripcio"
Combo2.AddItem "Producto"
Combo2.ListIndex = 0

opcion1 = "1"
bproducto = ""
Frame3.Visible = True
bproducto.SetFocus
Label11_Click
End Sub

Private Sub seccion11_Click()
Combo2.Clear
Combo2.AddItem "Descripcio"
Combo2.ListIndex = 0

opcion1 = "2"
bproducto = ""
Frame3.Visible = True
bproducto.SetFocus
Label11_Click
End Sub
Sub totaliza_secciones()
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim mytablex As New ADODB.Recordset
sdx = 0
sdx1 = 0
sdx2 = 0
mytablex.Open "select * from seccionproduccion where tarjeta='" & "" & txplanop.Fields("tarjeta") & "'", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
'MsgBox "abc"
sdx = sdx + Val("" & mytablex.Fields("costomaterial"))
sdx1 = sdx1 + Val("" & mytablex.Fields("costomanoobra"))
sdx2 = sdx2 + Val("" & mytablex.Fields("merma"))
mytablex.MoveNext
Loop
mytablex.Close
txplanop.Fields("merma") = sdx2
txplanop.Fields("costomaterial") = sdx
txplanop.Fields("costomanoobra") = sdx1
txplanop.Fields("costototal") = sdx + sdx1
txplanop.Fields("costo") = (sdx + sdx1) / Val("" & txplanop.Fields("cantidad"))
txplanop.Update
txplanop.Requery
End Sub
