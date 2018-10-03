VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form expfpa1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimiento de Clientes"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   14715
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
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
      Height          =   9015
      Left            =   360
      TabIndex        =   42
      Top             =   6735
      Visible         =   0   'False
      Width           =   14655
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   9720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   7680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "EXPFPA1.frx":0000
         Height          =   1215
         Left            =   120
         OleObjectBlob   =   "EXPFPA1.frx":0014
         TabIndex        =   61
         Top             =   7680
         Visible         =   0   'False
         Width           =   8655
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
         Height          =   495
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   45
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
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
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
         Left            =   8160
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "EXPFPA1.frx":09E7
         Height          =   6735
         Left            =   120
         OleObjectBlob   =   "EXPFPA1.frx":09FB
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   840
         Width           =   14415
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   15600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Mensajes"
      ForeColor       =   &H00FFFFFF&
      Height          =   3930
      Left            =   2940
      TabIndex        =   38
      Top             =   2415
      Visible         =   0   'False
      Width           =   9015
      Begin VB.Label Label4 
         Caption         =   "ESPERE UN MOMENTO..PROCESANDO..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   480
         TabIndex        =   39
         Top             =   720
         Width           =   7935
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "EXPFPA1.frx":13C6
      Height          =   6255
      Left            =   120
      OleObjectBlob   =   "EXPFPA1.frx":13DA
      TabIndex        =   30
      Top             =   1320
      Width           =   14535
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   14160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   14655
      TabIndex        =   0
      Top             =   0
      Width           =   14715
      Begin VB.ComboBox vendedor 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Nombre 
         Height          =   375
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   40
         Text            =   "%"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox local1 
         Height          =   375
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   28
         Text            =   "%"
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   2520
         MaxLength       =   13
         TabIndex        =   26
         Text            =   "%"
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox ordenado 
         BackColor       =   &H00FFFFFF&
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
         Left            =   10920
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   720
         Width           =   1815
      End
      Begin VB.ComboBox tipoclie 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   1815
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
         Left            =   13080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "EXPFPA1.frx":34AD
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.ComboBox tipo 
         BackColor       =   &H00FFFFFF&
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
         Left            =   10920
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   12
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cajero 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox caja 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   0
         Width           =   1815
      End
      Begin VB.ComboBox turno 
         BackColor       =   &H00FFFFFF&
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
         Left            =   10920
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   0
         Width           =   1815
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
         Picture         =   "EXPFPA1.frx":3C5B
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "EXPFPA1.frx":4E6D
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         Height          =   375
         Left            =   7080
         TabIndex        =   48
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   375
         Left            =   4320
         TabIndex        =   41
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   1560
         TabIndex        =   29
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   375
         Left            =   1560
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenado Por"
         Height          =   375
         Left            =   9840
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoCod"
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocto"
         Height          =   375
         Left            =   9840
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         Height          =   375
         Left            =   9840
         TabIndex        =   10
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   375
         Left            =   7080
         TabIndex        =   9
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         Height          =   375
         Left            =   7080
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label comprad 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   960
      TabIndex        =   60
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label ventad 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2280
      TabIndex        =   59
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label ingresod 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      TabIndex        =   58
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label egresod 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3600
      TabIndex        =   57
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label egresos 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3600
      TabIndex        =   56
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Egresos"
      Height          =   375
      Left            =   3600
      TabIndex        =   55
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label ingresos 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      TabIndex        =   54
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ingresos"
      Height          =   375
      Left            =   4920
      TabIndex        =   53
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label ventas 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2280
      TabIndex        =   52
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ventas"
      Height          =   375
      Left            =   2280
      TabIndex        =   51
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label compras 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   960
      TabIndex        =   50
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Compras"
      Height          =   375
      Left            =   960
      TabIndex        =   49
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "US$"
      Height          =   375
      Left            =   8280
      TabIndex        =   37
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S/."
      Height          =   375
      Left            =   8280
      TabIndex        =   36
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label cargod 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8880
      TabIndex        =   35
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label abonod 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   34
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Label saldod 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11640
      TabIndex        =   33
      Top             =   8520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label saldos 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11640
      TabIndex        =   32
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label abonos 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   31
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abono"
      Height          =   375
      Left            =   10200
      TabIndex        =   25
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label cargos 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8880
      TabIndex        =   24
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   375
      Left            =   11640
      TabIndex        =   23
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cargo"
      Height          =   375
      Left            =   8880
      TabIndex        =   22
      Top             =   7800
      Width           =   1335
   End
   Begin VB.Label afecta 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   14520
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label acu 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   13680
      TabIndex        =   3
      Top             =   7440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu dfki2323 
      Caption         =   "&Ejecuta"
   End
   Begin VB.Menu dki2232 
      Caption         =   "&Imprime"
   End
   Begin VB.Menu lfo3434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "expfpa1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()

End Sub

Private Sub cmdDelete_Click()
    dbo912_Click

End Sub

Private Sub cmdGrabar_Click()

End Sub

Private Sub cmdSort_Click()

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        lfo3434_Click
        Exit Sub

    End If

    Command2_Click

End Sub

Private Sub cmdExit_Click()
    lfo3434_Click

End Sub

Private Sub cmdPrint_Click()

    If Frame1.Visible = True Then Exit Sub

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    borrar_archivo FileName
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento1
    cuerpo_programa_documento1
    '------------------------------------
    Close #1
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_codigo

    End If

End Sub

Private Sub Command1_Click()

    If Frame1.Visible = True Then Exit Sub
    xborrar
    sql_recibos

End Sub

Private Sub dbo912_Click()

End Sub

Private Sub dki9923_Click()

End Sub

Private Sub dnu823_Click()

End Sub

Private Sub Command2_Click()

    ejecuta 1

End Sub

Private Sub dbgrid1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_pagos

    End If

End Sub

Private Sub DBGrid2_DblClick()

    On Error GoTo cmd78121_err

    If opcion1 = "2" Then
        busca_fpagox "" & Data1.Recordset.Fields("local"), "" & Data1.Recordset.Fields("tipo"), "" & Data1.Recordset.Fields("serie"), "" & Data1.Recordset.Fields("numero")

    End If

    Exit Sub
cmd78121_err:
    Exit Sub

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            codigo = DBGrid2.columns(1)
            Frame2.Visible = False
            codigo.SetFocus

            'codigo_KeyPress 13
        End If

        If opcion1 = "2" Then
            Frame2.Visible = False
            codigo.SetFocus

            'codigo_KeyPress 13
        End If

    End If

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim buf  As String

    Dim buf2 As String

    Dim sw   As Integer

    If KeyCode <> 13 And KeyCode <> 27 Then

        'MsgBox KeyCode
        If KeyCode >= 48 And KeyCode <= 57 Then
            GoTo sigue9

        End If

        If KeyCode >= 65 And KeyCode <= 90 Then
            GoTo sigue9

        End If

        If KeyCode >= 97 And KeyCode <= 122 Then
            GoTo sigue9

        End If

        If KeyCode = 8 And Chr(KeyCode) = "*" Then
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

        buf = Chr(KeyCode)

        If Chr(KeyCode) = "*" Then
            buf = ""
            buffer = buf

        End If

        If KeyCode <> 13 Then
            buffer = buffer + buf

        End If

        buf = buffer
        ejecuta 0

    End If

End Sub

Private Sub dfki2323_Click()

    If Frame2.Visible = True Then Exit Sub
    Command1_Click

End Sub

Private Sub dki2232_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    cmdPrint_Click

End Sub

Private Sub Form_Activate()
    Frame1.Top = 1800: Frame1.Left = 2880
    Frame2.Top = 20: Frame2.Left = 20

    fechai = Format(Now, "dd/mm/yyyy") '"01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    carga_inicial
    'sql_recibos

End Sub

Sub consulta_codigo()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0
    Frame2.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command2_Click

End Sub

Sub consulta_pagos()
    Combo1.Clear
    Combo1.AddItem "Numero"
    Combo1.ListIndex = 0
    Frame2.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "2"
    Command2_Click

End Sub

Sub carga_inicial()

    Dim mytablex As Table

    vendedor.Clear
    vendedor.AddItem "%"

    cajero.Clear
    cajero.AddItem "%"
    Set mytablex = mydbxglo.OpenTable("vendedor")
    Do

        If mytablex.EOF Then Exit Do
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0
    vendedor.ListIndex = 0

    caja.Clear
    caja.AddItem "%"
    Set mytablex = mydbxglo.OpenTable("parameca")
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("terminal") = "C" Then
            caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("DESCRIPCIO")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    caja.ListIndex = 0

    turno.Clear
    turno.AddItem "%"
    Set mytablex = mydbxglo.OpenTable("turno")
    Do

        If mytablex.EOF Then Exit Do
        turno.AddItem "" & mytablex.Fields("turno") & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    turno.ListIndex = 0

    tipo.Clear
    tipo.AddItem "%"
    Set mytablex = mydbxglo.OpenTable("tipo")
    Do

        If mytablex.EOF Then Exit Do
        tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    tipo.ListIndex = 0

End Sub

Private Sub Form_Load()
    tipoclie.Clear
    tipoclie.AddItem "%"
    tipoclie.AddItem "C"
    tipoclie.AddItem "P"
    tipoclie.AddItem "V"
    tipoclie.ListIndex = 0

    ordenado.Clear
    ordenado.AddItem "fecha"
    ordenado.AddItem "tipo"
    ordenado.AddItem "val(numero)"
    ordenado.AddItem "Codigo"
    ordenado.AddItem "Usuario"
    ordenado.AddItem "caja"
    ordenado.AddItem "turno"
    ordenado.AddItem "fpago"
    ordenado.AddItem "orden"
    ordenado.AddItem "observa"
    ordenado.AddItem "descripcio"
    ordenado.AddItem "nombre"
    ordenado.ListIndex = 0

End Sub

Private Sub lfo3434_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    expfpa1.Hide
    Unload expfpa1

End Sub

Sub sql_recibos()

    On Error GoTo cmd38_err

    Dim buf As String

    Dim vr

    Dim mytablez As Table

    Dim xcargo   As Double

    Dim xabono   As Double

    Dim mytabley As Table

    Dim mytablex As Snapshot

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    Set mytablez = mydbxglo.OpenTable("fpagov")
    mytablez.Index = "fpagov"
    Frame1.Visible = True
    Set mytabley = mydbxglo.OpenTable("_b" + gusuario)
    buf = "select * from factura where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario='" & extra_loquesea(cajero) & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor='" & extra_loquesea(vendedor) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo='" & extra_loquesea(tipo) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja='" & extra_loquesea(caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno='" & extra_loquesea(turno) & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If tipoclie <> "%" Then
        buf = buf & " and tipoclie='" & tipoclie & "'"

    End If

    buf = buf & " and estado='2' "
    'buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='W') "
    'buf = buf & " order by " & ordenado & ", val(numero)"
    'MsgBox buf
    Set mytablex = mydbxglo.CreateSnapshot(buf)
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Frame1.Visible = False Then
            Exit Do

        End If

        mytabley.AddNew
        mytabley.Fields("local") = "" & mytablex.Fields("local")
        mytabley.Fields("tipo") = "" & mytablex.Fields("tipo")
        mytabley.Fields("tipoclie") = "" & mytablex.Fields("tipoclie")
        mytabley.Fields("nombret") = busca_tipo("" & mytablex.Fields("tipo"))
        mytabley.Fields("serie") = "" & mytablex.Fields("serie")
        mytabley.Fields("numero") = "" & mytablex.Fields("numero")
        mytabley.Fields("codigo") = "" & mytablex.Fields("codigo")
        mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
        mytabley.Fields("fecha") = Format(CVDate("" & mytablex.Fields("fecha")), "dd/mm/yyyy")
        mytabley.Fields("moneda") = "" & mytablex.Fields("moneda")
        mytabley.Fields("usuario") = "" & mytablex.Fields("usuario")
        mytabley.Fields("caja") = "" & mytablex.Fields("caja")
        mytabley.Fields("turno") = "" & mytablex.Fields("turno")
        mytabley.Fields("total") = Val("" & mytablex.Fields("total"))
        'ventas
        xcargo = 0
        xabono = 0

        If "" & mytablex.Fields("acu") = "A" Or "" & mytablex.Fields("acu") = "B" Or "" & mytablex.Fields("acu") = "C" Or "" & mytablex.Fields("acu") = "D" Or "" & mytablex.Fields("acu") = "G" Then  'ventas
            If "" & mytablex.Fields("moneda") = "S" Then
                xventas = xventas + Val("" & mytablex.Fields("total"))

            End If

            If "" & mytablex.Fields("moneda") = "D" Then
                xventad = xventad + Val("" & mytablex.Fields("total"))

            End If

            mytablez.Seek "=", "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo"), "" & mytablex.Fields("serie"), "" & mytablex.Fields("numero")

            If Not mytablez.NoMatch Then
                Do

                    If mytablez.EOF Then Exit Do
                    If "" & mytablez.Fields("local") = "" & mytablex.Fields("local") And "" & mytablez.Fields("tipo") = "" & mytablex.Fields("tipo") And "" & mytablez.Fields("serie") = "" & mytablex.Fields("serie") And "" & mytablez.Fields("numero") = "" & mytablex.Fields("numero") Then
                        '--------------------------------------------------
                        buf = busca_fpago("" & mytablez.Fields("fpago"))

                        'abono
                        If buf = "A" Or buf = "B" Or buf = "E" Then  'EFECTIVO
                            xabono = xabono + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "D" Or buf = "F" Or buf = "J" Then   'TARJETA
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "C" Or buf = "G" Then  'CREDITO O OLETRA
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "H" Then   'DEPOSITO BANCO
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If
         
                        If buf = "I" Then   'ADELANTO
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        '--------------------------------------------------
                        Else: Exit Do

                    End If

                    mytablez.MoveNext
                Loop

            End If

        End If

        If "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Or "" & mytablex.Fields("acu") = "P" Then  'compras

            'MsgBox "x"
            If "" & mytablex.Fields("moneda") = "S" Then
                xcompras = xcompras + Val("" & mytablex.Fields("total"))

            End If

            If "" & mytablex.Fields("moneda") = "D" Then
                xcomprad = xcomprad + Val("" & mytablex.Fields("total"))

            End If

            mytablez.Seek "=", "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo"), "" & mytablex.Fields("serie"), "" & mytablex.Fields("numero")

            If Not mytablez.NoMatch Then
                Do

                    If mytablez.EOF Then Exit Do
                    If "" & mytablez.Fields("local") = "" & mytablex.Fields("local") And "" & mytablez.Fields("tipo") = "" & mytablex.Fields("tipo") And "" & mytablez.Fields("serie") = "" & mytablex.Fields("serie") And "" & mytablez.Fields("numero") = "" & mytablex.Fields("numero") Then
                        '--------------------------------------------------
                        buf = busca_fpago("" & mytablez.Fields("fpago"))

                        'abono
                        If buf = "A" Or buf = "B" Or buf = "E" Then  'EFECTIVO
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "D" Or buf = "F" Or buf = "J" Then  'TARJETA
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "C" Or buf = "G" Then  'CREDITO O OLETRA
                            xabono = xabono + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "H" Then   'DEPOSITO BANCO
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "I" Then   'ADELANTO
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        '--------------------------------------------------
                        Else: Exit Do

                    End If

                    mytablez.MoveNext
                Loop

            End If

        End If

        mytabley.Fields("cargo") = xcargo
        mytabley.Fields("abono") = xabono
        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close

    '------------------ orden de trabajo ---------------------------------------
    buf = "select * from cpedidov where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario='" & extra_loquesea(cajero) & "'"

    End If

    If tipo = "%" Then
        buf = buf & " and tipo='60' "

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo='" & extra_loquesea(tipo) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja='" & extra_loquesea(caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno='" & extra_loquesea(turno) & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If tipoclie <> "%" Then
        buf = buf & " and tipoclie='" & tipoclie & "'"

    End If

    buf = buf & " and estado='2' "
    Set mytablex = mydbxglo.CreateSnapshot(buf)
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Frame1.Visible = False Then
            Exit Do

        End If

        mytabley.AddNew
        mytabley.Fields("local") = "" & mytablex.Fields("local")
        mytabley.Fields("tipo") = "" & mytablex.Fields("tipo")
        mytabley.Fields("tipoclie") = "" & mytablex.Fields("tipoclie")
        mytabley.Fields("nombret") = busca_tipo("" & mytablex.Fields("tipo"))
        mytabley.Fields("serie") = "" & mytablex.Fields("serie")
        mytabley.Fields("numero") = "" & mytablex.Fields("numero")
        mytabley.Fields("codigo") = "" & mytablex.Fields("codigo")
        mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
        mytabley.Fields("fecha") = "" & mytablex.Fields("fecha")
        mytabley.Fields("moneda") = "" & mytablex.Fields("moneda")
        mytabley.Fields("usuario") = "" & mytablex.Fields("usuario")
        mytabley.Fields("caja") = "" & mytablex.Fields("caja")
        mytabley.Fields("turno") = "" & mytablex.Fields("turno")
        mytabley.Fields("total") = "" & mytablex.Fields("total")
        'ventas
        xcargo = 0
        xabono = 0

        If "" & mytablex.Fields("moneda") = "S" Then
            xingresos = xingresos + Val("" & mytablex.Fields("total"))

        End If

        If "" & mytablex.Fields("moneda") = "D" Then
            xingresod = xingresod + Val("" & mytablex.Fields("total"))

        End If
   
        mytablez.Seek "=", "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo"), "" & mytablex.Fields("serie"), "" & mytablex.Fields("numero")

        If Not mytablez.NoMatch Then
            Do

                If mytablez.EOF Then Exit Do
                If "" & mytablez.Fields("local") = "" & mytablex.Fields("local") And "" & mytablez.Fields("tipo") = "" & mytablex.Fields("tipo") And "" & mytablez.Fields("serie") = "" & mytablex.Fields("serie") And "" & mytablez.Fields("numero") = "" & mytablex.Fields("numero") Then
                    '--------------------------------------------------
                    buf = busca_fpago("" & mytablez.Fields("fpago"))

                    'abono
                    If buf = "A" Or buf = "B" Or buf = "E" Then  'EFECTIVO
                        xabono = xabono + Val("" & mytablez.Fields("total"))

                    End If

                    If buf = "D" Or buf = "F" Or buf = "J" Then  'TARJETA
                        xcargo = xcargo + Val("" & mytablez.Fields("total"))

                    End If

                    If buf = "C" Or buf = "G" Then  'CREDITO O OLETRA
                        xcargo = xcargo + Val("" & mytablez.Fields("total"))

                    End If

                    If buf = "H" Then   'CRUCE DEPOSITO BANCO
                        xcargo = xcargo + Val("" & mytablez.Fields("total"))

                    End If

                    If buf = "I" Then   'CRUCE ADELANTO
                        xcargo = xcargo + Val("" & mytablez.Fields("total"))

                    End If

                    '--------------------------------------------------
                    Else: Exit Do

                End If

                mytablez.MoveNext
            Loop

        End If

        mytabley.Fields("cargo") = xcargo
        mytabley.Fields("abono") = xabono
        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    '------------------ ingresos egresos ----------------------------------------
    buf = "select * from recibo where "
    buf = buf & "  fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario='" & extra_loquesea(cajero) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo='" & extra_loquesea(tipo) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja='" & extra_loquesea(caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno='" & extra_loquesea(turno) & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If tipoclie <> "%" Then
        buf = buf & " and tipoclie='" & tipoclie & "'"

    End If

    buf = buf & " and estado='2' "
    Set mytablex = mydbxglo.CreateSnapshot(buf)
    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Frame1.Visible = False Then
            Exit Do

        End If

        mytabley.AddNew
        mytabley.Fields("local") = "" & mytablex.Fields("local")
        mytabley.Fields("tipo") = "" & mytablex.Fields("tipo")
        mytabley.Fields("tipoclie") = "" & mytablex.Fields("tipoclie")
        mytabley.Fields("nombret") = busca_tipo("" & mytablex.Fields("tipo"))
        mytabley.Fields("serie") = "" & mytablex.Fields("serie")
        mytabley.Fields("numero") = "" & mytablex.Fields("numero")
        mytabley.Fields("codigo") = "" & mytablex.Fields("codigo")
        mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
        mytabley.Fields("fecha") = "" & mytablex.Fields("fecha")
        mytabley.Fields("moneda") = "" & mytablex.Fields("moneda")
        mytabley.Fields("usuario") = "" & mytablex.Fields("usuario")
        mytabley.Fields("caja") = "" & mytablex.Fields("caja")
        mytabley.Fields("turno") = "" & mytablex.Fields("turno")
        mytabley.Fields("total") = "" & mytablex.Fields("total")
        'ventas
        xcargo = 0
        xabono = 0

        If "" & mytablex.Fields("acu") = "W" Then   'ingresos
            If "" & mytablex.Fields("moneda") = "S" Then
                xingresos = xingresos + Val("" & mytablex.Fields("total"))

            End If

            If "" & mytablex.Fields("moneda") = "D" Then
                xingresod = xingresod + Val("" & mytablex.Fields("total"))

            End If

            mytablez.Seek "=", "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo"), "" & mytablex.Fields("serie"), "" & mytablex.Fields("numero")

            If Not mytablez.NoMatch Then
                Do

                    If mytablez.EOF Then Exit Do
                    If "" & mytablez.Fields("local") = "" & mytablex.Fields("local") And "" & mytablez.Fields("tipo") = "" & mytablex.Fields("tipo") And "" & mytablez.Fields("serie") = "" & mytablex.Fields("serie") And "" & mytablez.Fields("numero") = "" & mytablex.Fields("numero") Then
                        '--------------------------------------------------
                        buf = busca_fpago("" & mytablez.Fields("fpago"))

                        'abono
                        If buf = "A" Or buf = "B" Or buf = "E" Then   'EFECTIVO
                            xabono = xabono + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "D" Or buf = "F" Then  'TARJETA
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "C" Or buf = "G" Then  'CREDITO O OLETRA
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "H" Then   'DEPOSITO BANCO
                            xabono = xabono + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "I" Then   'ADELANTO
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        '--------------------------------------------------
                        Else: Exit Do

                    End If

                    mytablez.MoveNext
                Loop

            End If

        End If

        If "" & mytablex.Fields("acu") = "V" Then   'egresos
            If "" & mytablex.Fields("moneda") = "S" Then
                xegresos = xegresos + Val("" & mytablex.Fields("total"))

            End If

            If "" & mytablex.Fields("moneda") = "D" Then
                xegresod = xegresod + Val("" & mytablex.Fields("total"))

            End If

            mytablez.Seek "=", "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo"), "" & mytablex.Fields("serie"), "" & mytablex.Fields("numero")

            If Not mytablez.NoMatch Then
                Do

                    If mytablez.EOF Then Exit Do
                    If "" & mytablez.Fields("local") = "" & mytablex.Fields("local") And "" & mytablez.Fields("tipo") = "" & mytablex.Fields("tipo") And "" & mytablez.Fields("serie") = "" & mytablex.Fields("serie") And "" & mytablez.Fields("numero") = "" & mytablex.Fields("numero") Then
                        '--------------------------------------------------
                        buf = busca_fpago("" & mytablez.Fields("fpago"))

                        'abono
                        If buf = "A" Or buf = "B" Or buf = "E" Then  'EFECTIVO
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "D" Or buf = "F" Then  'TARJETA
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "C" Or buf = "G" Then  'CREDITO O OLETRA
                            xabono = xabono + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "H" Then   'DEPOSITO BANCO
                            xcargo = xcargo + Val("" & mytablez.Fields("total"))

                        End If

                        If buf = "I" Then   'ADELANTO
                            xabono = xabono + Val("" & mytablez.Fields("total"))

                        End If

                        '--------------------------------------------------
                        Else: Exit Do

                    End If

                    mytablez.MoveNext
                Loop

            End If

        End If

        mytabley.Fields("cargo") = xcargo
        mytabley.Fields("abono") = xabono
        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    mytablez.Close
    Ventas = Format(xventas, "0.00")
    ventad = Format(xventad, "0.00")
    compras = Format(xcompras, "0.00")
    comprad = Format(xcomprad, "0.00")
    ingresos = Format(xingresos, "0.00")
    ingresod = Format(xingresod, "0.00")
    egresos = Format(xegresos, "0.00")
    egresod = Format(xegresod, "0.00")
    '------------------- fin de ingresos egresos---------------------------------

    Data2.Connect = "foxpro 2.5;"
    Data2.DatabaseName = globaldir
    Data2.RecordSource = "select * from _b" & gusuario
    Data2.refresh
    sumar_recibos
    Frame1.Visible = False

    Exit Sub
cmd38_err:
    Frame1.Visible = False
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub xborrar()

    On Error GoTo cmd112_err

    Data2.Database.Execute "DELETE FROM _b" & gusuario
    Exit Sub
cmd112_err:
    Exit Sub

End Sub

Sub sumar_recibos()

    Dim xcargos   As Double

    Dim xabonos   As Double

    Dim xcargod   As Double

    Dim xabonod   As Double

    Dim xcompras  As Double

    Dim xcomprad  As Double

    Dim xventas   As Double

    Dim xventad   As Double

    Dim xingresos As Double

    Dim xingresod As Double

    Dim xegresos  As Double

    Dim xegresod  As Double

    xcompras = 0
    xcomprad = 0
    xventas = 0
    xventad = 0
    xingresos = 0
    xingresod = 0
    xegresod = 0
    xegresos = 0

    xcargos = 0
    xabonos = 0
    xcargod = 0
    xabonod = 0

    Data2.refresh
    Do

        If Data2.Recordset.EOF Then Exit Do
        If "" & Data2.Recordset.Fields("moneda") = "S" Then
            xcargos = xcargos + Val("" & Data2.Recordset.Fields("cargo"))
            xabonos = xabonos + Val("" & Data2.Recordset.Fields("abono"))

        End If

        If "" & Data2.Recordset.Fields("moneda") = "D" Then
            xcargod = xcargod + Val("" & Data2.Recordset.Fields("cargo"))
            xabonod = xabonod + Val("" & Data2.Recordset.Fields("abono"))

        End If

        Data2.Recordset.MoveNext
    Loop
    cargos = Format(xcargos, "0.00")
    abonos = Format(xabonos, "0.00")
    saldos = Format(xabonos - xcargos, "0.00")
    cargod = Format(xcargod, "0.00")
    abonod = Format(xabonod, "0.00")
    saldod = Format(xabonod - xcargod, "0.00")

End Sub

Function busca_tipo(buf As String)

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("tipo")
    mytablex.Index = "tipo"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_tipo = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

Function busca_fpago(buf As String) As String

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("fpago")
    mytablex.Index = "fpago"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_fpago = "" & mytablex.Fields("tipo")

    End If

    mytablex.Close

End Function

Sub cabecera_documento1()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = "Movimiento de Clientes  "
    found = formateaa(buf, 90, 2, 0)
    
    found = formateaa("Lo", 3, 0, 0)
    found = formateaa("Tp", 3, 0, 0)
    found = formateaa("Srie", 5, 0, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("Nombre", 31, 0, 0)
    
    found = formateaa("Cargo ", 11, 0, 1)
    found = formateaa("Abono ", 11, 0, 1)
    found = formateaa("Saldo", 11, 2, 1)
        
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)

End Sub

Sub cuerpo_programa_documento1()

    Dim buf    As String

    Dim found  As Integer

    Dim xcargo As Double

    Dim xabono As Double

    On Error GoTo cmd78812_err

    xcargo = 0
    xabono = 0
    Data2.refresh
    Do

        If Data2.Recordset.EOF Then Exit Do
        buf = "" & Data2.Recordset.Fields("LOCAL")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & Data2.Recordset.Fields("tipo")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & Data2.Recordset.Fields("serie")
        found = formateaa(buf, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & Data2.Recordset.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & Data2.Recordset.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & Data2.Recordset.Fields("nombre")
        found = formateaa(buf, 30, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = Format(Val("" & Data2.Recordset.Fields("cargo")), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(Val("" & Data2.Recordset.Fields("abono")), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
      
        nlineas
      
        xcargo = xcargo + Val("" & Data2.Recordset.Fields("cargo"))
        xabono = xabono + Val("" & Data2.Recordset.Fields("abono"))
      
        Data2.Recordset.MoveNext
    Loop

    found = formateaa("", 65, 0, 0)
    buf = Format(xcargo, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(xabono, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(xcargo - xabono, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
      
    Exit Sub
cmd78812_err:
    Exit Sub

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > 45 Then
        cabecera_documento1

    End If

End Sub

Sub ejecuta(sw As Integer)

    On Error GoTo cmd7812_err

    Dim buf As String

    dbgrid3.Visible = False

    If opcion1 = "1" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from clientes"
        Else
            buf = "select Nombre,Codigo from clientes where  " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If opcion1 = "2" Then
        'If Len(buffer) = 0 Then
        buf = "select Producto,Descripcio,Unidad,Factor,Cantidad as Cant,Precio,Total,Local,Tipo,Serie,str(NUmero) from detalle where "
        'buf = buf & "  fecha=" & "DateValue('" & Data2.Recordset.Fields("fecha") & "'" & ")"
        buf = buf & " local='" & "" & Data2.Recordset.Fields("local") & "'"
        buf = buf & " and tipo='" & "" & Data2.Recordset.Fields("tipo") & "'"
        buf = buf & " and serie='" & "" & Data2.Recordset.Fields("serie") & "'"
        buf = buf & " and numero='" & "" & Data2.Recordset.Fields("numero") & "'"
        buf = buf & " and codigo='" & "" & Data2.Recordset.Fields("codigo") & "'"

        'Else
        'buf = "select Nombre,Codigo from clientes where  " & Combo1 & " like '" & buffer & "*'"
        'End If
    End If

    'MsgBox buf
    Data1.Connect = "foxpro 2.5;"
    Data1.DatabaseName = globaldir
    Data1.RecordSource = buf
    Data1.refresh

    If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
        Data1.Recordset.Close
        buffer.SetFocus
        Exit Sub

    End If

    If opcion1 = "1" Then
        DBGrid2.columns(0).Width = 2000
        DBGrid2.columns(1).Width = 1300

    End If

    If opcion1 = "2" Then
        DBGrid2.columns(0).Width = 1200
        DBGrid2.columns(1).Width = 4000
        DBGrid2.columns(2).Width = 700
        DBGrid2.columns(3).Width = 700
        DBGrid2.columns(4).Width = 900
        DBGrid2.columns(5).Width = 900
        DBGrid2.columns(6).Width = 900
        DBGrid2.columns(7).Width = 900
        DBGrid2.columns(8).Width = 900
        DBGrid2.columns(9).Width = 900
        DBGrid2.columns(10).Width = 900

    End If

    If sw = 1 Then
        DBGrid2.SetFocus

    End If

    Exit Sub
cmd7812_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub busca_fpagox(bxlocal As String, _
                 bxtipo As String, _
                 bxserie As String, _
                 bxnumero As String)

    Dim buf As String

    On Error GoTo cmd45_err

    dbgrid3.Visible = True
    buf = "Select Fpago,Descripcio,Total,Recibe from fpagov where local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'"
    Data3.Connect = "foxpro 2.5;"
    Data3.DatabaseName = globaldir
    Data3.RecordSource = buf
    Data3.refresh
    Exit Sub
cmd45_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub
               
End Sub

