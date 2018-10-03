VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tplanif 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plaanificacion"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reporte Orden Trabajo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Materiales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aprobar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden Trabajo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   7455
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   13150
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "ordentrabajo"
         Caption         =   "OrdenTrabajo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13322
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Fecha"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13322
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Fechae"
         Caption         =   "FechaEnt"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13322
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Estado"
         Caption         =   "Estado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13322
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "aprobado"
         Caption         =   "Aprobado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13322
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1530.142
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox mostrar 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
   Begin VB.ComboBox ordenado 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.ComboBox subtablapro 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label tablapro 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mostrar"
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ordenado Pro"
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
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TipoFormula"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu lfosa02 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tplanif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim txplanifica As New ADODB.Recordset

Private Sub Command1_Click()

    If subtablapro = "%" Then
        MsgBox "Debe elegir una opcion ", 48, "Aviso"
        Exit Sub

    End If

    ttordent.Show 1

End Sub

Private Sub Command6_Click()
    sql_inicio

End Sub

Private Sub Form_Activate()

    'sql_inicio
End Sub

Private Sub Form_Load()
    ordenado.Clear
    ordenado.AddItem "%"
    ordenado.AddItem "OrdenTrabajo"
    ordenado.AddItem "Fecha"
    ordenado.AddItem "FechaE"
    ordenado.ListIndex = 0

    mostrar.Clear
    mostrar.AddItem "%"
    mostrar.AddItem "PLANIFICACION"
    mostrar.AddItem "PRODUCCION"
    mostrar.AddItem "ANULADAS"
    mostrar.AddItem "TERMINADOS"
    mostrar.AddItem "ATRASADOS"
    mostrar.ListIndex = 0

    sql_inicio
    carga_tipo

End Sub

Sub sql_inicio()

    Dim buf As String

    buf = "select * from ordentrabajoc  "

    If txplanifica.State = 1 Then
        txplanifica.Close

    End If

    Set txplanifica = Nothing

    txplanifica.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txplanifica

End Sub

Sub carga_tipo()

    Dim mytablex As New ADODB.Recordset

    subtablapro.Clear
    subtablapro.AddItem "%"
    mytablex.Open "select * from subtablapro where tablapro='" & tablapro & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        subtablapro.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("subtablapro"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    subtablapro.ListIndex = 0

End Sub

Private Sub lfosa02_Click()
    tplanif.Hide
    Unload tplanif

End Sub
