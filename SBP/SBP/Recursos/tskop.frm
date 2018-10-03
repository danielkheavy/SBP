VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form tskop 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Ordenes de Produccion"
   ClientHeight    =   10650
   ClientLeft      =   165
   ClientTop       =   -45
   ClientWidth     =   18030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   18030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   9975
      Left            =   0
      TabIndex        =   80
      Top             =   0
      Visible         =   0   'False
      Width           =   15015
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
         Height          =   495
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   240
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   8895
         Left            =   120
         TabIndex        =   82
         Top             =   840
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   15690
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   29
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
               LCID            =   10250
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
               LCID            =   10250
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
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Materia Prima"
      Height          =   8655
      Left            =   9240
      TabIndex        =   74
      Top             =   6480
      Visible         =   0   'False
      Width           =   15015
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
         Height          =   735
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Borra"
         Height          =   735
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Modify"
         Height          =   735
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add"
         Height          =   735
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   360
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dbgrid14 
         Height          =   8175
         Left            =   240
         TabIndex        =   75
         Top             =   360
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   14420
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   29
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
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "producto"
            Caption         =   "Producto"
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
         BeginProperty Column01 
            DataField       =   "Descripcio"
            Caption         =   "Descripcio"
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
         BeginProperty Column02 
            DataField       =   "Cantidad"
            Caption         =   "Cant"
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
            DataField       =   "Unidad"
            Caption         =   "Und"
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
            DataField       =   "Factor"
            Caption         =   "Factor"
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
         BeginProperty Column05 
            DataField       =   "Precio"
            Caption         =   "ValorU"
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
         BeginProperty Column06 
            DataField       =   "Total"
            Caption         =   "Total"
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
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5040
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1019.906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame2"
      Height          =   9975
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   15015
      Begin VB.TextBox oobserva 
         Height          =   3135
         Left            =   1680
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   64
         Top             =   6240
         Width           =   5055
      End
      Begin VB.TextBox occosto 
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
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   62
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox onumero1 
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
         TabIndex        =   61
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox oserie1 
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
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   59
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox oestado 
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
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   52
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox ocantidad 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   50
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox oproducto 
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
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   48
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox ocodigo 
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
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   46
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox ofechaf 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   44
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox ofechae 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   42
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox ofechai 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   40
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox ofecha 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   38
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox olote 
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
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   36
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox oid 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   945
         Left            =   13080
         Picture         =   "tskop.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Imprimir todo"
         Top             =   1320
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   13080
         Picture         =   "tskop.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1470
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   8520
         TabIndex        =   69
         Top             =   1680
         Visible         =   0   'False
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16777152
         Appearance      =   1
         StartOfWeek     =   114294785
         CurrentDate     =   41762
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   71
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8520
         TabIndex        =   70
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcio"
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
         TabIndex        =   68
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Indicaciones para Produccion"
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
         TabIndex        =   67
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label onombre 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   66
         Top             =   2760
         Width           =   6375
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   65
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Centro Costo"
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
         TabIndex        =   63
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie/Numero"
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
         TabIndex        =   60
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label ofactor 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   58
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   57
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label ounidad 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   56
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   55
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label odescripcio 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1680
         TabIndex        =   54
         Top             =   3840
         Width           =   6375
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   53
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   51
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   49
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   47
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaVenc."
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
         Width           =   1575
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaEntrega"
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
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
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
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Doc"
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
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote"
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
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   14955
      TabIndex        =   1
      Top             =   0
      Width           =   15015
      Begin VB.TextBox numero1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         MaxLength       =   11
         TabIndex        =   34
         Text            =   "%"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox serie1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         MaxLength       =   4
         TabIndex        =   32
         Text            =   "%"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox numero 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         MaxLength       =   11
         TabIndex        =   31
         Text            =   "%"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox serie 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "%"
         Top             =   480
         Width           =   735
      End
      Begin Proyecto1.EC_Button command1 
         Height          =   495
         Left            =   12120
         TabIndex        =   28
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         Caption         =   "Filtrar"
         FontColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShadowText      =   0   'False
         Angle           =   0
         GradientColor1  =   0
         GradientColor2  =   0
         GradientButton  =   0   'False
         MaskColor       =   0
         BackColor       =   0
         Style           =   0
      End
      Begin VB.ComboBox estado 
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   27
         Top             =   120
         Width           =   2175
      End
      Begin VB.ComboBox lote 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox producto 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         MaxLength       =   15
         TabIndex        =   23
         Text            =   "%"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox codigo 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         MaxLength       =   11
         TabIndex        =   21
         Text            =   "%"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox fechaef 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox fechaei 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   16
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox fechaf 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox fechai 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   12
         Top             =   120
         Width           =   1575
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
         Picture         =   "tskop.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tskop.frx":23A6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Borrar registro"
         Top             =   600
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tskop.frx":35B8
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tskop.frx":47CA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprimir"
         Top             =   600
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
         Picture         =   "tskop.frx":59DC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Doc/Relac."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   33
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie/Num"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   30
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   25
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CodigoCliente"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   20
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaEntregaFn"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaEntregaIn"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   15015
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Procesos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   7200
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "MateriaPrima"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   7200
         Width           =   2535
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   6735
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   11880
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   17
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
            Name            =   "Arial"
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
               LCID            =   10250
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
               LCID            =   10250
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
Attribute VB_Name = "tskop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sktabla As New ADODB.Recordset
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
oid.Enabled = False
oid = ""
olote.SetFocus
End Sub

Private Sub bo712_Click()
Dim buf As String
On Error GoTo cmd656_err
buf = "" & sktabla.Fields("id")
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If
If MsgBox("Desea Borra " + "" & sktabla.Fields("id"), 1, "Aviso") <> 1 Then
   Exit Sub
End If
sktabla.Delete
Command1_Click



Exit Sub
cmd656_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub buffer_DblClick()

Command2_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
Command2_Click
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


Private Sub Command2_Click()
filtro
End Sub

Private Sub dbgrid13_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
If opcion1 = "1" Then
   oproducto = Trim("" & dbgrid13.columns(1))
   odescripcio = Trim("" & dbgrid13.columns(0))
   ounidad = Trim("" & dbgrid13.columns(2))
   ofactor = Trim("" & dbgrid13.columns(3))
   Frame3.Visible = False
   ocantidad.SetFocus
End If
If opcion1 = "2" Then
   ocodigo = Trim("" & dbgrid13.columns(1))
   onombre = Trim("" & dbgrid13.columns(0))
   Frame3.Visible = False
   oserie1.SetFocus
End If
If opcion1 = "3" Then
   olote = Trim("" & dbgrid13.columns(1))
   Frame3.Visible = False
   olote.SetFocus
End If
If opcion1 = "4" Then
   occosto = Trim("" & dbgrid13.columns(1))
   Frame3.Visible = False
   occosto.SetFocus
End If
If opcion1 = "5" Then
   oestado = Trim("" & dbgrid13.columns(1))
   Frame3.Visible = False
   oestado.SetFocus
End If

End If

End Sub

Private Sub dk9893_Click()
If Frame2.Visible = True Then Exit Sub
reporgen.NAMETABLA = "skydopro"
reporgen.Show 1

End Sub
Sub prueba_reporte()
'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\seccionesproducto.rpt", "")
End Sub



Private Sub Command1_Click()
Frame1.Visible = True
Frame1.Enabled = True

ejecuta 1

End Sub
Sub ejecuta(sw As Integer)
Dim cad As String
   cad = "SELECT skydopro.Id,skydopro.estado,skydopro.producto,producto.descripcio,skydopro.unidad,skydopro.factor,skydopro.Cantidad,skydopro.codigo,clientes.nombre,skydopro.fechaI as Inicio,skydopro.fechaE as Entrega ,skydopro.fechaf as Vence,skydopro.fecha as FechaDoc from skydopro    "
   cad = cad & " inner join producto on skydopro.producto=producto.producto inner join clientes on skydopro.codigo=clientes.codigo "
   cad = cad & " and skydopro.fechai>='" & Format(fechai, "YYYYMMDD") & "'"
   cad = cad & " and skydopro.fechai<='" & Format(fechaf, "YYYYMMDD") & "' "
   cad = cad & " and skydopro.fechae>='" & Format(fechai, "YYYYMMDD") & "'"
   cad = cad & " and skydopro.fechae<='" & Format(fechaf, "YYYYMMDD") & "' "
   If lote <> "%" Then
   cad = cad & " and skydopro.lote='" & extra_loquesea(lote) & "'"
   End If
  
   If estado <> "%" Then
   cad = cad & " and skydopro.estado='" & extra_loquesea(estado) & "'"
   End If
   If codigo <> "%" Then
   cad = cad & " and skydopro.codigo like '" & codigo & "'"
   End If
   If serie1 <> "%" Then
   cad = cad & " and skydopro.serie1 like '" & serie1 & "'"
   End If
   If numero1 <> "%" Then
   cad = cad & " and skydopro.numero1 like '" & numero1 & "'"
   End If
   If producto <> "%" Then
   cad = cad & " and skydopro.producto like '" & producto & "'"
   End If
   'MsgBox cad
   
   
   If sktabla.State = 1 Then sktabla.Close
   sktabla.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbGrid1.DataSource = sktabla
   dbGrid1.columns(0).Width = 800
   dbGrid1.columns(1).Width = 1000
   dbGrid1.columns(2).Width = 1200
   dbGrid1.columns(3).Width = 2500
   dbGrid1.columns(4).Width = 800
   dbGrid1.columns(5).Width = 800
   dbGrid1.columns(6).Width = 800
   dbGrid1.columns(7).Width = 1200
   dbGrid1.columns(8).Width = 2000
   dbGrid1.columns(9).Width = 1000
   dbGrid1.columns(10).Width = 1000
   dbGrid1.columns(11).Width = 1000
   If sktabla.RecordCount > 0 Then
     dbGrid1.SetFocus
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
tskop.Hide
Unload tskop
End Sub


Private Sub EC_Button1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub f8443_Click()
Dim buf As String
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd456_err
buf = sktabla.Fields("id")
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If

inicializa
Frame2.Visible = True
Frame2.Caption = "Modifica"
cmdGuardar.Enabled = True
mytablex.Open "select * from skydopro where id=" & buf & "", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   pone_registro mytablex
End If
mytablex.Close
habilita 1
oid.Enabled = False
olote.SetFocus
Exit Sub
cmd456_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub fjh433_Click()
Dim buf As String
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd556_err
buf = sktabla.Fields("id")
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If

inicializa
Frame2.Visible = True
Frame2.Caption = "Zoom"
cmdGuardar.Enabled = False
mytablex.Open "select * from skydopro where id=" & buf & "", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   pone_registro mytablex
End If
mytablex.Close
habilita 1
oid.Enabled = False
olote.SetFocus
Exit Sub
cmd556_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub Form_Activate()
'agregar_menus
cargas

Command1_Click
End Sub

Private Sub Form_Load()
fechai = Format(Now, "dd/mm/yyyy")
fechaf = Format(Now, "dd/mm/yyyy")
fechaei = Format(Now, "dd/mm/yyyy")
fechaef = Format(Now, "dd/mm/yyyy")
End Sub
Sub inicializa()
onombre = ""
odescripcio = ""
oid = ""
'oserie = ""
'onumero = ""
oserie1 = ""
onumero1 = ""
olote = ""
ofecha = Format(Now, "dd/mm/yyyy")
ofechai = Format(Now, "dd/mm/yyyy")
ofechaf = Format(Now, "dd/mm/yyyy")
ofechae = Format(Now, "dd/mm/yyyy")
ocodigo = ""
'ovendedor = ""
oproducto = ""
odescripcio = ""
ounidad = ""
ofactor = ""
ocantidad = "1"
oestado = ""
occosto = ""
'oproyecto = ""
'opartida = ""
oobserva = ""
'ototal = "0.00"
End Sub
Sub pone_registro(mytablex As ADODB.Recordset)
oid = Trim("" & mytablex.Fields("id"))
oserie1 = Trim("" & mytablex.Fields("serie1"))
onumero1 = Trim("" & mytablex.Fields("numero1"))
ofecha = Trim("" & mytablex.Fields("fecha"))
ofechai = Trim("" & mytablex.Fields("fechai"))
ofechaf = Trim("" & mytablex.Fields("fechaf"))
ofechae = Trim("" & mytablex.Fields("fechae"))
ocodigo = Trim("" & mytablex.Fields("codigo"))
oproducto = Trim("" & mytablex.Fields("producto"))
ounidad = Trim("" & mytablex.Fields("unidad"))
ofactor = Trim("" & mytablex.Fields("factor"))
ocantidad = Trim("" & mytablex.Fields("cantidad"))
oestado = Trim("" & mytablex.Fields("estado"))
occosto = Trim("" & mytablex.Fields("ccosto"))
oobserva = Trim("" & mytablex.Fields("observa"))
'ototal = Trim("" & mytablex.Fields("total"))
End Sub

Sub grabando(mytablex As ADODB.Recordset)
mytablex.Fields("serie1") = Trim(oserie1)
mytablex.Fields("numero1") = Trim(onumero1)
mytablex.Fields("fecha") = Trim(ofecha)
mytablex.Fields("fechai") = Trim(ofechai)
mytablex.Fields("fechaf") = Trim(ofechaf)
mytablex.Fields("fechae") = Trim(ofechae)
mytablex.Fields("codigo") = Trim(ocodigo)
mytablex.Fields("producto") = Trim(oproducto)
mytablex.Fields("cantidad") = Trim(ocantidad)
mytablex.Fields("unidad") = Trim(ounidad)
mytablex.Fields("factor") = Val(ofactor)
mytablex.Fields("estado") = Trim(oestado)
mytablex.Fields("ccosto") = Trim(occosto)
mytablex.Fields("observa") = Trim(oobserva)
End Sub

Private Sub grba1_Click()

End Sub

Function grabar()
Dim found As Integer
Dim mytablex As New ADODB.Recordset
found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If
If Frame2.Caption = "Nuevo" Then
   'If Len(seccion) = 0 Then
   '   seccion.SetFocus
   '   Exit Function
   'End If
   mytablex.Open "select * from skydopro where 2=1", cn, adOpenStatic, adLockOptimistic
   mytablex.AddNew
   grabando mytablex
   mytablex.Update
   mytablex.Close
   dlo132_Click
   Exit Function
End If
If Frame2.Caption = "Modifica" Then
    mytablex.Open "select * from skydopro where id=" & oid & "", cn, adOpenStatic, adLockOptimistic
    If mytablex.RecordCount > 0 Then
       grabando mytablex
       mytablex.Update
    End If
    mytablex.Close
   dlo132_Click
   Exit Function
End If

End Function

Function valida()
'If Len(seccion) = 0 Then
'   seccion.SetFocus
'   Exit Function
'End If
'If Len(oserie) = 0 Then
'   oserie.SetFocus
'   Exit Function
'End If
'If Len(onumero) = 0 Then
'   onumero.SetFocus
'   Exit Function
'End If
If Not IsDate(ofecha) Then
   ofecha.SetFocus
   Exit Function
End If
If Not IsDate(ofechai) Then
   ofechai.SetFocus
   Exit Function
End If
If Not IsDate(ofechaf) Then
   ofechaf.SetFocus
   Exit Function
End If
If Not IsDate(ofechae) Then
   ofechae.SetFocus
   Exit Function
End If
If Len(oproducto) = 0 Then
   oproducto.SetFocus
   Exit Function
End If
If Val(ocantidad) = 0 Then
   ocantidad.SetFocus
   Exit Function
End If
If Len(oestado) = 0 Then
   oestado.SetFocus
   Exit Function
End If
If Len(ocodigo) = 0 Then
   ocodigo.SetFocus
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
   mytablex.Open "select * from archivo where menu='SECCION' and   estado='S'", cn, adOpenStatic, adLockOptimistic
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

Private Sub Label16_Click()
   MonthView1.Visible = True
   Label2 = "FechaDoc"
End Sub

Private Sub Label17_Click()
   MonthView1.Visible = True
   Label2 = "FechaInicio"

End Sub

Private Sub Label18_Click()
   MonthView1.Visible = True
   Label2 = "FechaEntrega"

End Sub

Private Sub Label19_Click()
   MonthView1.Visible = True
   Label2 = "FechaVencimiento"

End Sub

Private Sub Label24_Click()
MonthView1.Visible = False
End Sub

Sub mnuarchivoarray_click(Index As Integer)
Dim mytablex As New ADODB.Recordset
Dim buf As String
buf = mnuArchivoArray(Index).Caption
   mytablex.Open "select * from archivo where menu='SECCION' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
   End If
   'busca el reporte
   buf = mytablex.Fields("archivo")
   mytablex.Close
   'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub
Sub cargas()
Dim mytablex As New ADODB.Recordset
lote.Clear
estado.Clear
estado.AddItem "%"
lote.AddItem "%"

 mytablex.Open "select * from lote order by descripcio", cn, adOpenStatic, adLockOptimistic
 Do
 If mytablex.EOF Then Exit Do
 lote.AddItem "" & mytablex.Fields("nombre") & "|" & mytablex.Fields("lote")
 mytablex.MoveNext
 Loop
 mytablex.Close
 lote.ListIndex = 0

mytablex.Open "select * from estadopro order by descripcio", cn, adOpenStatic, adLockOptimistic
 Do
 If mytablex.EOF Then Exit Do
 estado.AddItem "" & mytablex.Fields("descripcio") & "|" & mytablex.Fields("estadopro")
 mytablex.MoveNext
 Loop
 mytablex.Close
 estado.ListIndex = 0

End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
If Label2 = "FechaDoc" Then
   ofecha = Format(MonthView1.Value, "dd/mm/yyyy")
   MonthView1.Visible = False
   ofecha.SetFocus
   Exit Sub
End If
If Label2 = "FechaInicio" Then
   ofechai = Format(MonthView1.Value, "dd/mm/yyyy")
   MonthView1.Visible = False
   ofechai.SetFocus
   Exit Sub
End If
If Label2 = "FechaEntrega" Then
   ofechae = Format(MonthView1.Value, "dd/mm/yyyy")
   MonthView1.Visible = False
   ofechae.SetFocus
   Exit Sub
End If
If Label2 = "FechaVencimiento" Then
   ofechaf = Format(MonthView1.Value, "dd/mm/yyyy")
   MonthView1.Visible = False
   ofechaf.SetFocus
   Exit Sub
End If

End Sub

Private Sub noccosto_Click()
If noccosto <> "%" Then
   occosto = extra_loquesea1(noccosto)
End If
End Sub

Private Sub noestado_Click()
If noestado <> "%" Then
   oestado = extra_loquesea1(noestado)
End If
End Sub

Private Sub nolote_Click()
If nolote <> "%" Then
   olote = extra_loquesea1(nolote)
End If
End Sub

Private Sub ocantidad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
occosto.SetFocus

End Sub

Private Sub occosto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
oestado.SetFocus

End Sub

Private Sub occosto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_ccosto
End If

End Sub

Private Sub ocodigo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
oserie1.SetFocus

End Sub

Private Sub ocodigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_codigo
End If

End Sub

Private Sub oestado_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
oobserva.SetFocus

End Sub

Private Sub oestado_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_estado
End If

End Sub

Private Sub ofecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
ofechai.SetFocus

End Sub

Private Sub ofechae_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
ofechaf.SetFocus

End Sub

Private Sub ofechaf_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
ocodigo.SetFocus

End Sub

Private Sub ofechai_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
ofechae.SetFocus

End Sub

Private Sub olote_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
ofecha.SetFocus

End Sub

Private Sub olote_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_lote
End If

End Sub

Private Sub onumero1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
onumero1.SetFocus

End Sub

Private Sub oproducto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(Trim(oproducto)) = 0 Then
   oproducto.SetFocus
   Exit Sub
End If
found = busca_producto()
If found = 0 Then
   oproducto = ""
   oproducto.SetFocus
   Exit Sub
End If
ocantidad.SetFocus

End Sub

Private Sub oproducto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_producto
End If

End Sub

Private Sub oserie1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
oproducto.SetFocus

End Sub
Sub consulta_producto()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Producto"
Combo1.ListIndex = 0
Frame3.Enabled = True
Frame3.Visible = True
buffer = ""
opcion1 = "1"
buffer.SetFocus
Command2_Click


End Sub
Sub consulta_lote()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Lote"
Combo1.ListIndex = 0
Frame3.Enabled = True
Frame3.Visible = True
buffer = ""
opcion1 = "3"
buffer.SetFocus
Command2_Click


End Sub
Sub consulta_ccosto()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "CCosto"
Combo1.ListIndex = 0
Frame3.Enabled = True
Frame3.Visible = True
buffer = ""
opcion1 = "4"
buffer.SetFocus
Command2_Click


End Sub
Sub consulta_estado()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "EstadoPro"
Combo1.ListIndex = 0
Frame3.Enabled = True
Frame3.Visible = True
buffer = ""
opcion1 = "5"
buffer.SetFocus
Command2_Click


End Sub
Sub consulta_codigo()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
Frame3.Enabled = True
Frame3.Visible = True
buffer = ""
opcion1 = "2"
buffer.SetFocus
Command2_Click


End Sub

Sub filtro()
Dim mytablex As New ADODB.Recordset
Dim cad As String
If opcion1 = "1" Then  'producto
   If Len(buffer) = 0 Then
      cad = "select Descripcio,Producto,Unidad,factor from producto "
   End If
   If Len(buffer) > 0 Then
      cad = "select Descripcio,Producto,Unidad,Factor from producto where " & Combo1 & " like '" & buffer & "%'"
   End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbgrid13.DataSource = mytablex
               dbgrid13.columns(0).Width = 5000
               dbgrid13.columns(1).Width = 2000
               dbgrid13.columns(2).Width = 1000
               dbgrid13.columns(3).Width = 1000
           
   End If
   
   If opcion1 = "2" Then  'codigo
   If Len(buffer) = 0 Then
      cad = "select Nombre,Codigo from Clientes "
   End If
   If Len(buffer) > 0 Then
      cad = "select Nombre,Codigo from Clientes where " & Combo1 & " like '" & buffer & "%'"
   End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbgrid13.DataSource = mytablex
               dbgrid13.columns(0).Width = 5000
               dbgrid13.columns(1).Width = 2000
               
           
   End If
   
   If opcion1 = "3" Then  'codigo
   If Len(buffer) = 0 Then
      cad = "select Descripcio,Lote from Lote "
   End If
   If Len(buffer) > 0 Then
      cad = "select Descripcio,Lote from lote where " & Combo1 & " like '" & buffer & "%'"
   End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbgrid13.DataSource = mytablex
               dbgrid13.columns(0).Width = 5000
               dbgrid13.columns(1).Width = 2000
               
           
   End If
   
   If opcion1 = "4" Then  'codigo
   If Len(buffer) = 0 Then
      cad = "select Descripcio,Ccosto from Ccosto "
   End If
   If Len(buffer) > 0 Then
      cad = "select Descripcio,Ccosto from Ccosto where " & Combo1 & " like '" & buffer & "%'"
   End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbgrid13.DataSource = mytablex
               dbgrid13.columns(0).Width = 5000
               dbgrid13.columns(1).Width = 2000
               
           
   End If
   
    If opcion1 = "5" Then  'codigo
   If Len(buffer) = 0 Then
      cad = "select Descripcio,Estadopro from estadopro "
   End If
   If Len(buffer) > 0 Then
      cad = "select Descripcio,estadoprofrom estadopro where " & Combo1 & " like '" & buffer & "%'"
   End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbgrid13.DataSource = mytablex
               dbgrid13.columns(0).Width = 5000
               dbgrid13.columns(1).Width = 2000
               
           
   End If
   
   If mytablex.RecordCount > 0 Then
      dbgrid13.SetFocus
   End If
   Exit Sub

End Sub

