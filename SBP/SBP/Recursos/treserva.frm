VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form treserva 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Reservas"
   ClientHeight    =   10635
   ClientLeft      =   165
   ClientTop       =   -60
   ClientWidth     =   19170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   19170
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   9975
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   14895
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
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
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   8895
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   14175
         _ExtentX        =   25003
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame2"
      Height          =   8895
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   14895
      Begin VB.CommandButton Command2 
         Caption         =   "BorrarServicio"
         DisabledPicture =   "treserva.frx":0000
         Height          =   495
         Left            =   6600
         Picture         =   "treserva.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   7680
         Width           =   1455
      End
      Begin VB.CommandButton CmdCan 
         Caption         =   "BorrarHabitacion"
         DisabledPicture =   "treserva.frx":0684
         Height          =   495
         Left            =   120
         Picture         =   "treserva.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   7680
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2295
         Left            =   120
         TabIndex        =   57
         Top             =   5280
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4048
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Habitacion"
            Caption         =   "habitacion"
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
            DataField       =   "Precio"
            Caption         =   "Precio"
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
            DataField       =   "adulto"
            Caption         =   "Adulto"
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
            DataField       =   "nino"
            Caption         =   "Niño"
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
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   854.929
            EndProperty
         EndProperty
      End
      Begin VB.TextBox quienpaga 
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
         TabIndex        =   54
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox agente 
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
         TabIndex        =   52
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox arribohoraf 
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
         MaxLength       =   5
         TabIndex        =   51
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox salon 
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
         MaxLength       =   3
         TabIndex        =   46
         Top             =   4920
         Width           =   4095
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
         TabIndex        =   38
         Top             =   3480
         Width           =   6375
      End
      Begin VB.TextBox arribohora 
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
         MaxLength       =   5
         TabIndex        =   37
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox nino 
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
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   33
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox adulto 
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
         MaxLength       =   2
         TabIndex        =   31
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox operador 
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
         TabIndex        =   29
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox procedencia 
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
         TabIndex        =   27
         Top             =   3840
         Width           =   6375
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
         TabIndex        =   25
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox nombre 
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
         MaxLength       =   100
         TabIndex        =   23
         Top             =   2760
         Width           =   6375
      End
      Begin VB.TextBox mesa 
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
         Left            =   8760
         MaxLength       =   3
         TabIndex        =   21
         Top             =   4920
         Width           =   2775
      End
      Begin VB.TextBox arribofechaf 
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
         TabIndex        =   18
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox arribofecha 
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
         TabIndex        =   16
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox reserva 
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
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   13320
         Picture         =   "treserva.frx":0D08
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir todo"
         Top             =   1440
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   13320
         Picture         =   "treserva.frx":15D2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1470
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2295
         Left            =   6600
         TabIndex        =   58
         Top             =   5280
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4048
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
            DataField       =   "Producto"
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
            Caption         =   "Cantidad"
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
            DataField       =   "Precio"
            Caption         =   "Precio"
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
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2115.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin VB.Label totreserva 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9120
         TabIndex        =   64
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotalReserva"
         Height          =   615
         Left            =   9120
         TabIndex        =   63
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label totservicio 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   9480
         TabIndex        =   62
         Top             =   7680
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   375
         Left            =   8400
         TabIndex        =   61
         Top             =   7680
         Width           =   1095
      End
      Begin VB.Label tothabitacion 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4440
         TabIndex        =   60
         Top             =   7680
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   375
         Left            =   3360
         TabIndex        =   59
         Top             =   7680
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(H)uesped (E)mpresa"
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
         Left            =   2880
         TabIndex        =   56
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quien Paga"
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
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Agente"
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
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label noches 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6120
         TabIndex        =   50
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Noches"
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
         Left            =   6120
         TabIndex        =   49
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habitacion"
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
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Correo"
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
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hora (HH:MM)"
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
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad Personas"
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
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Niño"
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
         Left            =   3240
         TabIndex        =   34
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Adulto"
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
         TabIndex        =   32
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Operador"
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
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Procedencia"
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
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefono"
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
         TabIndex        =   26
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quien Reserva"
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
         TabIndex        =   24
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Servicios Adicionales"
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
         Left            =   6600
         TabIndex        =   22
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Check-In"
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
         TabIndex        =   20
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Check-Out"
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
         TabIndex        =   19
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
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
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reserva Id"
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
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   19110
      TabIndex        =   2
      Top             =   0
      Width           =   19170
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
         Picture         =   "treserva.frx":1E9C
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
         Picture         =   "treserva.frx":30AE
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
         Picture         =   "treserva.frx":42C0
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
         Picture         =   "treserva.frx":54D2
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
         Picture         =   "treserva.frx":66E4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label xsw 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   15480
         TabIndex        =   67
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   14895
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   12091
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
            DataField       =   "Reserva"
            Caption         =   "Reserva"
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
            DataField       =   "ArriboFecha"
            Caption         =   "ArriboFecha"
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
            DataField       =   "ArriboHora"
            Caption         =   "ArriboHora"
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
            DataField       =   "ARRIBOFechaf"
            Caption         =   "Salida"
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
            DataField       =   "arriboHoraf"
            Caption         =   "Hora"
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
            DataField       =   "Vendedor"
            Caption         =   "Vendedor"
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
            DataField       =   "Agente"
            Caption         =   "Agente"
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
         BeginProperty Column07 
            DataField       =   "Nombre"
            Caption         =   "Nombre"
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
         BeginProperty Column08 
            DataField       =   "QuienPaga"
            Caption         =   "QuienPaga"
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
         BeginProperty Column09 
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
         BeginProperty Column10 
            DataField       =   "Procedencia"
            Caption         =   "Procedencia"
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
         BeginProperty Column11 
            DataField       =   "telefono"
            Caption         =   "Telefono"
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
         BeginProperty Column12 
            DataField       =   "Correo"
            Caption         =   "Correo"
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
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2954.835
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   915.024
            EndProperty
         EndProperty
      End
      Begin VB.Label totalreserva 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   11760
         TabIndex        =   48
         Top             =   7200
         Width           =   2895
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   495
         Left            =   10080
         TabIndex        =   47
         Top             =   7200
         Width           =   1695
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
   Begin VB.Menu dk88343 
      Caption         =   "An&ticipo"
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
Attribute VB_Name = "treserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txreservax As New ADODB.Recordset
Dim mytablexx As New ADODB.Recordset
Dim mytableyy As New ADODB.Recordset


Private Sub agente_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_agente
End If

End Sub

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
reserva.Enabled = False
reserva = ""
operador.SetFocus
End Sub

Private Sub bo712_Click()
Dim buf As String

On Error GoTo cmd656_err
buf = "" & txreservax.Fields("reserva")
If Frame2.Visible = True Then
   dbGrid1.SetFocus
   Exit Sub
End If
If MsgBox("Desea Borra " + "" & txreservax.Fields("reserva"), 1, "Aviso") <> 1 Then
   Exit Sub
End If
cn.Execute ("delete from reservah where reserva=" & Val(reserva))
cn.Execute ("delete from reservas where reserva=" & Val(reserva))
cn.Execute ("delete from hotelanticipo where idreserva=" & Val(reserva))
txreservax.Delete
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

Private Sub CmdCan_Click()
On Error GoTo cmdxx90_err
mytablexx.Delete
carga_habitacion
Exit Sub
cmdxx90_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
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



Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_codigo
End If

End Sub

Private Sub Command2_Click()
On Error GoTo cmdyy90_err
mytableyy.Delete
carga_habitacion
Exit Sub
cmdyy90_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub Command4_Click()
filtro
End Sub
Sub filtro()
Dim mytablex As New ADODB.Recordset
Dim cad As String
If opcion1 = "1" Then  'producto
   If Len(Text1) = 0 Then
      cad = "select Habitacion,Descripcio,Estado,TipoHabitacion,Capacidad,precio from Habitacion "
   End If
   If Len(Text1) > 0 Then
      cad = "select Habitacion,Descripcio,Estado,TipoHabitacion,Capacidad,precio from Habitacion where  " & Combo2 & " like '" & Text1.Text & "%'"
   End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbgrid13.DataSource = mytablex
               dbgrid13.columns(0).Width = 800
               dbgrid13.columns(1).Width = 800
               dbgrid13.columns(2).Width = 1900
               dbgrid13.columns(3).Width = 900
               'dbgrid13.columns(4).Width = 900
               'dbgrid13.columns(2).Width = 1000
               'dbgrid13.columns(3).Width = 1000

   If mytablex.RecordCount > 0 Then
      dbgrid13.SetFocus
   End If
End If
If opcion1 = "2" Then  'producto
   If Len(Text1) = 0 Then
      cad = "select Nombre,Codigo,Telefono,Correo from clientes "
   End If
   If Len(Text1) > 0 Then
      cad = "select Nombre,Codigo,telefono,Correo from clientes where  " & Combo2 & " like '" & Text1.Text & "%'"
   End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbgrid13.DataSource = mytablex
               dbgrid13.columns(0).Width = 3000
               dbgrid13.columns(1).Width = 1000
              
   If mytablex.RecordCount > 0 Then
      dbgrid13.SetFocus
   End If
End If
   If opcion1 = "3" Then  'producto
   If Len(Text1) = 0 Then
      cad = "select Nombre,Codigo from Vendedor "
   End If
   If Len(Text1) > 0 Then
      cad = "select Nombre,Codigo from Vendedor where  " & Combo2 & " like '" & Text1.Text & "%'"
   End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbgrid13.DataSource = mytablex
               dbgrid13.columns(0).Width = 3000
               dbgrid13.columns(1).Width = 1000
              
   If mytablex.RecordCount > 0 Then
      dbgrid13.SetFocus
   End If
End If
If opcion1 = "4" Then  'producto
   If Len(Text1) = 0 Then
      cad = "select Nombre,Codigo from Vendedor "
   End If
   If Len(Text1) > 0 Then
      cad = "select Nombre,Codigo from Vendedor where  " & Combo2 & " like '" & Text1.Text & "%'"
   End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbgrid13.DataSource = mytablex
               dbgrid13.columns(0).Width = 3000
               dbgrid13.columns(1).Width = 1000
              
   If mytablex.RecordCount > 0 Then
      dbgrid13.SetFocus
   End If
End If
If opcion1 = "5" Then  'producto
   If Len(Text1) = 0 Then
      cad = "select producto.Descripcio,producto.producto,precios.Unidad1,precios.Factor1,precios.pventa1 from producto inner join precios on producto.producto=precios.producto "
   End If
   If Len(Text1) > 0 Then
      cad = "select producto.Descripcio,producto.producto,precios.Unidad1,precios.Factor1,precios.pventa1 from producto inner join precios on producto.producto=precios.producto and   " & Combo2 & " like '" & Text1.Text & "%'"
   End If
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbgrid13.DataSource = mytablex
               dbgrid13.columns(0).Width = 3000
               dbgrid13.columns(1).Width = 1000
              
   If mytablex.RecordCount > 0 Then
      dbgrid13.SetFocus
   End If
End If
      
   Exit Sub

End Sub

Private Sub DataGrid2_AfterColEdit(ByVal ColIndex As Integer)
Select Case ColIndex
       Case 2
           
            
End Select

End Sub

Private Sub dbgrid13_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mytablex As New ADODB.Recordset
Dim found As Integer
If KeyCode = 27 Then
   Text1.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
If opcion1 = "1" Then
   mytablex.Open "select * from reservah where reserva=" & Val(reserva) & " and  habitacion='" & Trim("" & dbgrid13.columns("habitacion")) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
   mytablex.AddNew
   mytablex.Fields("reserva") = Val(reserva)
   'mytablex.Fields("salon") = Trim("" & dbgrid13.columns("salon"))
   mytablex.Fields("habitacion") = Trim("" & dbgrid13.columns("habitacion"))
   mytablex.Fields("precio") = Val("" & dbgrid13.columns("precio"))
   mytablex.Fields("adulto") = Val(adulto)
   mytablex.Fields("nino") = Val(nino)
   mytablex.Update
   Else
   MsgBox "Ya existe ", 48, "Aviso"
   Exit Sub
   'mytablex.Fields("reserva") = Val(reserva)
   'mytablex.Fields("salon") = Trim("" & dbgrid13.columns("salon"))
   'mytablex.Fields("mesa") = Trim("" & dbgrid13.columns("mesa"))
   'mytablex.Fields("precio") = Val("" & dbgrid13.columns("precio"))
   'mytablex.Fields("adulto") = Val(adulto)
   'mytablex.Fields("nino") = Val(nino)
   'mytablex.Update
   End If
   mytablex.Close
   carga_habitacion
   salon.SetFocus
   Frame3.Visible = False
End If
If opcion1 = "2" Then
   nombre = Trim("" & dbgrid13.columns("nombre"))
   telefono = Trim("" & dbgrid13.columns("telefono"))
   correo = Trim("" & dbgrid13.columns("correo"))
   nombre.SetFocus
   Frame3.Visible = False
   
End If
If opcion1 = "3" Then
   operador = Trim("" & dbgrid13.columns("codigo"))
   Frame3.Visible = False
End If
If opcion1 = "4" Then
   agente = Trim("" & dbgrid13.columns("codigo"))
   Frame3.Visible = False
End If
If opcion1 = "5" Then
   mytablex.Open "select * from reservas where reserva=" & Val(reserva) & " and producto='" & Trim("" & dbgrid13.columns("producto")) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
   mytablex.AddNew
   mytablex.Fields("reserva") = Val(reserva)
   mytablex.Fields("producto") = Trim("" & dbgrid13.columns("producto"))
   mytablex.Fields("descripcio") = Trim("" & dbgrid13.columns("descripcio"))
   mytablex.Fields("unidad") = Trim("" & dbgrid13.columns("unidad1"))
   mytablex.Fields("factor") = Val("" & dbgrid13.columns("factor1"))
   mytablex.Fields("cantidad") = 1
   mytablex.Fields("precio") = Val("" & dbgrid13.columns("pventa1"))
   mytablex.Fields("Total") = Val("" & dbgrid13.columns("pventa1"))
   mytablex.Update
   Else
   MsgBox "Ya Existe ", 48, "Aviso"
   Exit Sub
   'mytablex.Fields("reserva") = Val(reserva)
   'mytablex.Fields("producto") = Trim("" & dbgrid13.columns("producto"))
   'mytablex.Fields("descripcio") = Trim("" & dbgrid13.columns("descripcio"))
   'mytablex.Fields("unidad") = Trim("" & dbgrid13.columns("unidad1"))
   'mytablex.Fields("factor") = Val("" & dbgrid13.columns("factor1"))
   'mytablex.Fields("cantidad") = 1
   'mytablex.Fields("precio") = Val("" & dbgrid13.columns("pventa1"))
   'mytablex.Fields("Total") = Val("" & dbgrid13.columns("pventa1"))
   'mytablex.Update
   End If
   mytablex.Close
   carga_habitacion
   salon.SetFocus
   Frame3.Visible = False


End If




End If

End Sub

Private Sub dk88343_Click()
Dim buf As String
On Error GoTo cmd86712_err
buf = "" & txreservax.Fields("reserva")
thotelan.idreserva = Trim(buf)
thotelan.tipopago = "R"
'thotelan.idhabitacion = Trim("" & txreservax.Fields("habitacion"))
thotelan.Show 1
Exit Sub
cmd86712_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub dk9893_Click()
If Frame2.Visible = True Then Exit Sub
reporgen.NAMETABLA = "reserva"
reporgen.Show 1

End Sub
Sub prueba_reporte()
'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\reservaesproducto.rpt", "")
End Sub


Private Sub Label21_Click()
consulta_mesas
End Sub

Private Sub Label6_Click()
consulta_producto
End Sub

Private Sub Label7_Click()
consulta_codigo
End Sub

Private Sub mesa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_producto
End If

End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_codigo
End If

End Sub

Private Sub operador_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_vendedor
End If

End Sub

Private Sub reserva_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
If Len(reserva) = 0 Then Exit Sub
'descripcio.SetFocus
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
Dim sdx As Double
   If Len(buffer) = 0 Then
      cad = "SELECT * from reserva    "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT *  from reserva   where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If txreservax.State = 1 Then txreservax.Close
   txreservax.Open cad, cn, adOpenStatic, adLockOptimistic
   Set dbGrid1.DataSource = txreservax
   'dbGrid1.columns(0).Width = 4000
   'dbGrid1.columns(1).Width = 2000
   If txreservax.RecordCount > 0 Then
     dbGrid1.SetFocus
  End If
  
  sdx = 0
Do
If txreservax.EOF Then Exit Do
sdx = sdx + Val("" & txreservax.Fields("Total"))
txreservax.MoveNext
Loop
totalreserva = Format(sdx, "0.00")

  

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   'reserva = dbGrid1.Columns(1)
   'Frame1.Visible = False
   'Frame1.Enabled = False
   'reserva.SetFocus
   'reserva_KeyPress 13
End If
End Sub



Private Sub dlo132_Click()
If Frame3.Visible = True Then
   Frame3.Visible = False
   ejecuta 1

   Exit Sub
End If

If Frame2.Visible = True Then
   habilita 0
   Frame2.Visible = False
   dbGrid1.Enabled = True
   ejecuta 1
   
   Exit Sub
End If
treserva.Hide
Unload treserva
End Sub


Private Sub f8443_Click()
Dim buf As String
On Error GoTo cmd456_err
buf = "" & txreservax.Fields("reserva")
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
reserva.Enabled = False
operador.SetFocus
Exit Sub
cmd456_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub fjh433_Click()
 Dim buf As String
On Error GoTo cmd556_err
buf = "" & txreservax.Fields("reserva")
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
reserva.Enabled = False
operador.SetFocus
Exit Sub
cmd556_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub

End Sub

Private Sub Form_Activate()
'agregar_menus
dk88343.Visible = False
If xsw = "ANTICIPO" Then
dk88343.Visible = True
End If
Command1_Click
End Sub
Sub consulta_mesas()
Combo2.Clear
Combo2.AddItem "Descripcio"
Combo2.AddItem "Habitacion"
Combo2.ListIndex = 0
Frame3.Enabled = True
Frame3.Visible = True
Text1 = ""
opcion1 = "1"
Text1.SetFocus
Command4_Click


End Sub
Sub consulta_codigo()
Combo2.Clear
Combo2.AddItem "Nombre"
Combo2.AddItem "Codigo"
Combo2.ListIndex = 0
Frame3.Enabled = True
Frame3.Visible = True
Text1 = ""
opcion1 = "2"
Text1.SetFocus
Command4_Click

End Sub
Sub consulta_producto()
Combo2.Clear
Combo2.AddItem "Descripcio"
Combo2.AddItem "Producto"
Combo2.ListIndex = 0
Frame3.Enabled = True
Frame3.Visible = True
Text1 = ""
opcion1 = "5"
Text1.SetFocus
Command4_Click

End Sub

Sub consulta_vendedor()
Combo2.Clear
Combo2.AddItem "Nombre"
Combo2.AddItem "Codigo"
Combo2.ListIndex = 0
Frame3.Enabled = True
Frame3.Visible = True
Text1 = ""
opcion1 = "3"
Text1.SetFocus
Command4_Click

End Sub
Sub consulta_agente()
Combo2.Clear
Combo2.AddItem "Nombre"
Combo2.AddItem "Codigo"
Combo2.ListIndex = 0
Frame3.Enabled = True
Frame3.Visible = True
Text1 = ""
opcion1 = "4"
Text1.SetFocus
Command4_Click

End Sub

Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0

End Sub
Sub inicializa()
arribofecha = Format(Now, "dd/mm/yyyy")
arribofechaf = Format(Now, "dd/mm/yyyy")
arribohora = Format(Now, "hh:mm")
arribohoraf = Format(Now, "hh:mm")
salon = ""
Mesa = ""
nombre = ""
telefono = ""
correo = ""
procedencia = ""
agente = ""
operador = ""
adulto = ""
nino = ""
noches = ""
quienpaga = ""
carga_habitacion
End Sub
Sub pone_registro()
reserva = Trim("" & txreservax.Fields("reserva"))
arribofecha = Trim("" & txreservax.Fields("arribofecha"))
arribofechaf = Trim("" & txreservax.Fields("arribofechaf"))
arribohora = Trim("" & txreservax.Fields("arribohora"))
arribohoraf = Trim("" & txreservax.Fields("arribohoraf"))
nombre = Trim("" & txreservax.Fields("nombre"))
telefono = Trim("" & txreservax.Fields("telefono"))
correo = Trim("" & txreservax.Fields("correo"))
procedencia = Trim("" & txreservax.Fields("procedencia"))
quienpaga = Trim("" & txreservax.Fields("quienpaga"))
operador = Trim("" & txreservax.Fields("operador"))
agente = Trim("" & txreservax.Fields("agente"))
totreserva = Trim("" & txreservax.Fields("total"))
carga_habitacion
End Sub
Sub grabando()
txreservax.Fields("arribofecha") = Trim(arribofecha)
txreservax.Fields("arribohora") = Trim(arribohora)
txreservax.Fields("arribofechaf") = Trim(arribofechaf)
txreservax.Fields("arribohoraf") = Trim(arribohoraf)
txreservax.Fields("nombre") = Trim(nombre)
txreservax.Fields("telefono") = Trim(telefono)
txreservax.Fields("correo") = Trim(correo)
txreservax.Fields("procedencia") = Trim(procedencia)
txreservax.Fields("agente") = Trim(agente)
txreservax.Fields("operador") = Trim(operador)
txreservax.Fields("quienpaga") = Trim(quienpaga)
txreservax.Fields("total") = Val(totreserva)
txreservax.Fields("tothabitacion") = Val(tothabitacion)
txreservax.Fields("totservicio") = Val(totservicio)

carga_habitacion
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
   'If Len(reserva) = 0 Then
   '   reserva.SetFocus
   '   Exit Function
   'End If
   'rbusca.Open "select reserva from reserva where reserva='" & reserva & "'", cn, adOpenStatic, adLockOptimistic
   'If rbusca.RecordCount > 0 Then
   '   rbusca.Close
   '   MsgBox "Ya existe reserva ", 48, "Aviso"
   '   Exit Function
   'End If
   txreservax.AddNew
   'txreservax.Fields("reserva") = reserva
   grabando
   txreservax.Update
   dlo132_Click
   Exit Function
End If
If Frame2.Caption = "Modifica" Then
   'txreservax.Fields("reserva") = reserva
   grabando
   txreservax.Update
   dlo132_Click
   Exit Function
End If

End Function

Function valida()
'If Len(reserva) = 0 Then
'   reserva.SetFocus
'   Exit Function
'End If
If Len(Trim(arribofecha)) < 10 Or Not IsDate(Trim(arribofecha)) Then
   arribofecha.SetFocus
   Exit Function
End If
If Len(Trim(arribofecha)) < 10 Or Not IsDate(Trim(arribofecha)) Then
   arribofecha.SetFocus
   Exit Function
End If
If Len(Trim(arribohora)) <> 5 Then
   arribohora.SetFocus
   Exit Function
End If
If Len(Trim(arribohoraf)) <> 5 Then
   arribohora.SetFocus
   Exit Function
End If

If Len(Trim(nombre)) = 0 Then
   nombre.SetFocus
   Exit Function
End If
If Len(Trim(operador)) = 0 Then
   operador.SetFocus
   Exit Function
End If
If Trim(quienpaga) <> "E" And Trim(quienpaga) <> "H" Then
   quienpaga.SetFocus
   Exit Function
End If
carga_habitacion

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
   mytablex.Open "select * from archivo where menu='reserva' and   estado='S'", cn, adOpenStatic, adLockOptimistic
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
Dim buf As String
buf = mnuArchivoArray(Index).Caption
   mytablex.Open "select * from archivo where menu='reserva' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
   End If
   'busca el reporte
   buf = mytablex.Fields("archivo")
   mytablex.Close
   'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")
End Sub



Private Sub salon_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_mesas
End If

End Sub

Sub carga_habitacion()
Dim sdx As Double
Dim sdx1 As Double
sdx = 0
sdx1 = 0
If mytablexx.State = 1 Then mytablexx.Close
If mytableyy.State = 1 Then mytableyy.Close
mytablexx.Open "select * from reservah where reserva=" & Val(reserva), cn, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = mytablexx
mytableyy.Open "select * from reservas where reserva=" & Val(reserva), cn, adOpenStatic, adLockOptimistic
Set DataGrid2.DataSource = mytableyy

Do
If mytablexx.EOF Then Exit Do
sdx = sdx + Val("" & mytablexx.Fields("precio"))
mytablexx.MoveNext
Loop

sdx1 = 0
Do
If mytableyy.EOF Then Exit Do
sdx1 = sdx1 + Val("" & mytableyy.Fields("total"))
mytableyy.MoveNext
Loop
tothabitacion = Format(sdx, "0.00")
totservicio = Format(sdx1, "0.00")
sdx1 = 0
totreserva = Format((sdx1 + sdx), "0.00")
End Sub

