VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tplaconc 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tabla de Conceptos de Planilla"
   ClientHeight    =   10080
   ClientLeft      =   165
   ClientTop       =   -615
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   6615
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   14040
      Begin VB.TextBox tipo 
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
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   600
         Width           =   6015
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   8400
         Picture         =   "tplaconc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Imprimir todo"
         Top             =   3120
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   8400
         Picture         =   "tplaconc.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2160
         Width           =   1470
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
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
         TabIndex        =   40
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   13935
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H8000000D&
         Caption         =   "&Ejecutar"
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
         Left            =   10800
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   6735
         Left            =   240
         TabIndex        =   34
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Creacion de Formulas"
      Height          =   6015
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   14055
      Begin VB.Frame Frame4 
         BackColor       =   &H0080FFFF&
         Height          =   3855
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   9135
         Begin VB.CommandButton Command7 
            Caption         =   "Cerrar"
            Height          =   615
            Left            =   5760
            TabIndex        =   25
            Top             =   960
            Width           =   975
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Grabar"
            Height          =   615
            Left            =   5760
            TabIndex        =   24
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox fporcentaje 
            Height          =   375
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   23
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox fcodigo 
            Height          =   375
            Left            =   1320
            MaxLength       =   6
            TabIndex        =   22
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label fdescripcio 
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1320
            TabIndex        =   29
            Top             =   840
            Width           =   4215
         End
         Begin VB.Label Label6 
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Descripcio"
            Height          =   375
            Left            =   360
            TabIndex        =   28
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label5 
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Porcentaje"
            Height          =   375
            Left            =   360
            TabIndex        =   27
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H0080FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Codigo"
            Height          =   375
            Left            =   360
            TabIndex        =   26
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Nuevo"
         Height          =   615
         Left            =   7440
         TabIndex        =   16
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Modifica"
         Height          =   615
         Left            =   7440
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Borra"
         Height          =   615
         Left            =   7440
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cerrar"
         Height          =   615
         Left            =   7440
         TabIndex        =   13
         Top             =   2640
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   3855
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   6800
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
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label atipo 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label adescripcio 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Formulas"
      Height          =   735
      Left            =   12600
      TabIndex        =   11
      Top             =   840
      Width           =   1335
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
         Picture         =   "tplaconc.frx":1194
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
         Picture         =   "tplaconc.frx":23A6
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
         Picture         =   "tplaconc.frx":35B8
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
         Picture         =   "tplaconc.frx":47CA
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
         Picture         =   "tplaconc.frx":59DC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
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
            DataField       =   "Tipo"
            Caption         =   "Tipo"
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
Attribute VB_Name = "tplaconc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txempre   As New ADODB.Recordset

Dim mytablex1 As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame5.Visible = True Then Exit Sub
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
    tipo.Enabled = True
    tipo = ""
    tipo.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    buf = txempre.Fields("tipo")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txempre.Fields("tipo"), 1, "Aviso") <> 1 Then
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

Private Sub bufferx_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame5.Visible = False
        Frame5.Enabled = False

        If opcion1 = "22" Then
            fcodigo.SetFocus
            Exit Sub

        End If
     
        Exit Sub

    End If

    Command9_Click

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
    djuer1_Click

End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub Command2_Click()
    fcodigo = ""
    fdescripcio = ""
    fporcentaje = ""
    Frame4.Visible = True
    Frame4.Caption = "Nuevo"
    fcodigo.Enabled = True
    fcodigo.SetFocus

End Sub

Private Sub Command3_Click()
    fcodigo = "" & dbgrid3.columns("codigo")
    fdescripcio = "" & dbgrid3.columns("descripcio")
    fporcentaje = "" & dbgrid3.columns("porcentaje")
    fcodigo.Enabled = False
    Frame4.Visible = True
    Frame4.Caption = "Modifica"
    fporcentaje.SetFocus

End Sub

Private Sub Command4_Click()

    On Error GoTo cmd245_err

    cn.Execute ("delete from tplanico1 where  tipo='" & dbgrid3.columns("tipo") & "' and codigo='" & dbgrid3.columns("codigo") & "'")
    Command8_Click
    Exit Sub
cmd245_err:
    Exit Sub

End Sub

Private Sub Command5_Click()
    Frame3.Visible = False

End Sub

Private Sub Command6_Click()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    If Len(fcodigo) = 0 Then
        fcodigo.SetFocus
        Exit Sub

    End If

    found = busca_codigo()

    If found = 0 Then
        MsgBox "No existe Codigo", 48, "Aviso"
        Exit Sub

    End If

    If Not IsNumeric(fporcentaje) Then
        fporcentaje.SetFocus
        Exit Sub

    End If

    If Frame4.Caption = "Nuevo" Then
        mytablex.Open "select * from tplanico1 where  tipo='" & atipo & "' and codigo='" & fcodigo & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            MsgBox "Ya existe codigo,No se puede adicionar ", 48, "Aviso"
            mytablex.Close
            Exit Sub

        End If

        mytablex.AddNew
        mytablex.Fields("tipo") = "" & atipo
        mytablex.Fields("codigo") = "" & fcodigo
        mytablex.Fields("porcentaje") = Val(fporcentaje)
        mytablex.Update
        mytablex.Close

    End If

    If Frame4.Caption = "Modifica" Then
        mytablex.Open "select * from tplanico1 where  tipo='" & atipo & "' and codigo='" & fcodigo & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            MsgBox "No existe codigo,No se puede Modificar ", 48, "Aviso"
            mytablex.Close
            Exit Sub

        End If
      
        'mytablex.Fields("tipo") = "" & atipo
        'mytablex.Fields("codigo") = "" & fcodigo
        mytablex.Fields("porcentaje") = Val(fporcentaje)
        mytablex.Update
        mytablex.Close

    End If

    Command7_Click

End Sub

Function busca_codigo()

    Dim mytablex As New ADODB.Recordset

    fdescripcio = ""
    mytablex.Open "select * from tplanico where  tipo='" & fcodigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        fdescripcio = "" & mytablex.Fields("descripcio")
        busca_codigo = 1

    End If

    mytablex.Close

End Function

Private Sub Command7_Click()
    Frame4.Visible = False
    Command8_Click

End Sub

Private Sub Command8_Click()

    On Error GoTo cmd34_err

    atipo = Trim("" & dbGrid1.columns(1))
    adescripcio = Trim("" & dbGrid1.columns(0))

    Frame3.Visible = True

    If mytablex1.State = 1 Then mytablex1.Close
    mytablex1.Open "select tplanico1.porcentaje,tplanico.descripcio,tplanico1.codigo,tplanico1.tipo from tplanico1,tplanico where tplanico.tipo=tplanico1.codigo and tplanico1.tipo='" & atipo & "'", cn, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = mytablex1
    dbgrid3.refresh
    dbgrid3.columns(0).Width = 1000
    dbgrid3.columns(1).Width = 3000
    dbgrid3.columns(2).Width = 2000
    dbgrid3.columns(3).Width = 2000

    Exit Sub
cmd34_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command9_Click()
    ejecutax 1

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        bufferx.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "22" Then
            fcodigo = Trim(DBGrid2.columns(1))
            fdescripcio = DBGrid2.columns(0)
            Frame5.Visible = False
            Frame5.Enabled = False
            fporcentaje.SetFocus

        End If

    End If

End Sub

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

Private Sub fcodigo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    fporcentaje.SetFocus

End Sub

Private Sub fcodigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_tipo

    End If

End Sub

Private Sub fporcentaje_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(tipo) = 0 Then Exit Sub
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

    If Len(buffer) = 0 Then
        cad = "SELECT * from tplanico    "

    End If

    If Len(buffer) > 0 Then
        cad = "SELECT *  from tplanico   where  " & Combo1 & " like '" & buffer & "%'"

    End If

    If txempre.State = 1 Then txempre.Close
    txempre.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txempre
    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000

    If txempre.RecordCount > 0 Then
        dbGrid1.SetFocus

    End If

End Sub

Sub ejecutax(sw As Integer)

    Dim cad      As String

    Dim mytablex As New ADODB.Recordset

    If Len(bufferx) = 0 Then
        cad = "SELECT Descripcio,Tipo from tplanico    "

    End If

    If Len(bufferx) > 0 Then
        cad = "SELECT Descripcio,Tipo from tplanico   where  " & Combo1 & " like '" & bufferx & "%'"

    End If

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = mytablex
    DBGrid2.columns(0).Width = 4000
    DBGrid2.columns(1).Width = 2000

    If mytablex.RecordCount > 0 Then
        DBGrid2.SetFocus

    End If

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'tipo = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'tipo.SetFocus
        'tipo_KeyPress 13
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

Private Sub djuer1_Click()

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "tplanico"
    reporgen.Show 1

End Sub

Private Sub dlo132_Click()

    If Frame5.Visible = True Then
        bufferx_KeyPress 27
        Exit Sub

    End If

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

    tplaconc.Hide
    Unload tplaconc

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    buf = txempre.Fields("tipo")

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
    tipo.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    If Frame5.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    buf = txempre.Fields("tipo")

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
    tipo.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    Command1_Click

End Sub

Sub consulta_tipo()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0

    Frame5.Visible = True
    Frame5.Enabled = True
    bufferx = ""
    opcion1 = "22"
    ejecutax 1

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "tipo"
    Combo1.ListIndex = 0

End Sub

Sub inicializa()

    descripcio = ""

End Sub

Sub pone_registro()
    tipo = Trim("" & txempre.Fields("tipo"))
    descripcio = Trim("" & txempre.Fields("descripcio"))

End Sub

Sub grabando()
    txempre.Fields("tipo") = Trim(tipo)
    txempre.Fields("descripcio") = Trim(descripcio)

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
        If Len(tipo) = 0 Then
            tipo.SetFocus
            Exit Function

        End If

        rbusca.Open "select tipo from tplanico where tipo='" & tipo & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe tipo ", 48, "Aviso"
            Exit Function

        End If

        txempre.AddNew
        txempre.Fields("tipo") = tipo
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        txempre.Fields("tipo") = tipo
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    'If Len(tipo) = 0 Then
    '   tipo.SetFocus
    '   Exit Function
    'End If
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

