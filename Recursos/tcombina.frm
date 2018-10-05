VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tcombina 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Combinaciones"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   14010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox observa 
      Height          =   495
      Left            =   5640
      MaxLength       =   40
      TabIndex        =   31
      Top             =   6600
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3975
      Left            =   960
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox cantidad 
         Height          =   375
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   3720
         TabIndex        =   28
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   2520
         TabIndex        =   27
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   1320
         TabIndex        =   26
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   3720
         TabIndex        =   24
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   2520
         TabIndex        =   23
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   1320
         TabIndex        =   22
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   3720
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   2520
         TabIndex        =   19
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   1320
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label numeros 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<Salir>"
         Height          =   495
         Left            =   5880
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<Select>"
         Height          =   495
         Left            =   5880
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   11245
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      HeadLines       =   2
      RowHeight       =   23
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
         DataField       =   "Descripciop"
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
      BeginProperty Column01 
         DataField       =   "Productop"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2954.835
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dbgrid3 
      Height          =   6375
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   11245
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
      ColumnCount     =   3
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "Descripciop"
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
         DataField       =   "productop"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3165.166
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Label ventana 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   6840
      Width           =   45
   End
   Begin VB.Label producto 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   6600
      Width           =   105
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Salir>"
      Height          =   495
      Left            =   10920
      TabIndex        =   11
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Graba>"
      Height          =   495
      Left            =   10920
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Borra>"
      Height          =   495
      Left            =   10920
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Baja>"
      Height          =   495
      Left            =   10920
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Sube>"
      Height          =   495
      Left            =   10920
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Cant>"
      Height          =   495
      Left            =   10920
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Baja>"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Sube>"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Select>"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label caja 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "tcombina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mytablexa As New ADODB.Recordset

Dim mytablexb As New ADODB.Recordset

Private Sub Form_Activate()

    Dim buf As String

    If ventana <> "S" Then
        buf = "select * from combina where combina.producto='" & producto & "' order by gcombina"

        If mytablexa.State = 1 Then
            mytablexa.Close
            Set mytablexa = Nothing

        End If

        mytablexa.Open buf, cn, adOpenStatic, adLockOptimistic
        Set DBGrid2.DataSource = mytablexa
        'cargar_tmcombina
        sql_combina

    End If

End Sub

Private Sub Label1_Click()

    Dim found As Integer

    If Frame1.Visible = True Then Exit Sub
    If mytablexa.RecordCount = 0 Then Exit Sub
    If mytablexa.EOF Or mytablexa.BOF Then Exit Sub
    found = sql_existe("" & mytablexa.Fields("productop"))

    If found = 1 Then
        MsgBox "Ya existe ", 48, "Aviso"
        Exit Sub

    End If

    sql_combina
    mytablexb.AddNew
    mytablexb.Fields("producto") = "" & mytablexa.Fields("producto")
    mytablexb.Fields("descripciop") = "" & mytablexa.Fields("descripciop")
    mytablexb.Fields("productop") = "" & mytablexa.Fields("productop")
    mytablexb.Fields("cantidad") = 1
    mytablexb.Update

End Sub

Private Sub Label11_Click()

    If mytablexb.RecordCount = 0 Then Exit Sub
    If Not IsNumeric(cantidad) Then
        cantidad = ""
        cantidad.SetFocus
        Exit Sub

    End If

    If Val(cantidad) <= 0 Then
        cantidad = "1"

    End If

    mytablexb.Fields("cantidad") = Val(cantidad)
    mytablexb.Update
    Frame1.Visible = False

End Sub

Private Sub Label12_Click()
    Frame1.Visible = False

End Sub

Private Sub Label2_Click()

    If Frame1.Visible = True Then Exit Sub
    If mytablexa.RecordCount = 0 Then Exit Sub
    If mytablexa.EOF Or mytablexa.BOF Then
        mytablexa.MoveFirst
        Exit Sub

    End If

    mytablexa.MovePrevious

End Sub

Private Sub Label3_Click()

    If Frame1.Visible = True Then Exit Sub
    If mytablexa.RecordCount = 0 Then Exit Sub
    If mytablexa.EOF Or mytablexa.BOF Then
        mytablexa.MoveFirst
        Exit Sub

    End If

    mytablexa.MoveNext

End Sub

Private Sub Label4_Click()

    If mytablexb.RecordCount = 0 Then Exit Sub
    Frame1.Visible = True
    cantidad = ""
    cantidad.SetFocus

End Sub

Private Sub Label5_Click()

    If Frame1.Visible = True Then Exit Sub
    If mytablexb.RecordCount = 0 Then Exit Sub
    If mytablexb.EOF Or mytablexb.BOF Then
        mytablexb.MoveFirst
        Exit Sub

    End If

    mytablexb.MovePrevious

End Sub

Private Sub Label6_Click()

    If Frame1.Visible = True Then Exit Sub
    If mytablexb.RecordCount = 0 Then Exit Sub
    If mytablexb.EOF Or mytablexb.BOF Then
        mytablexb.MoveFirst
        Exit Sub

    End If

    mytablexb.MoveNext

End Sub

Private Sub Label7_Click()

    Dim buf As String

    On Error GoTo cmd9012_err

    If mytablexb.RecordCount = 0 Then Exit Sub
    buf = "" & mytablexb.Fields("productop")

    If mytablexb.State = 1 Then
        mytablexb.Close
        Set mytablexb = Nothing

    End If

    cn.Execute ("delete from _c" & gusuario & " where producto='" & producto & "' and productop='" & buf & "'")
    sql_combina
    Exit Sub
cmd9012_err:
    Exit Sub

End Sub

Private Sub Label8_Click()

    Dim buf As String

    On Error GoTo cmd3_err

    If Frame1.Visible = True Then Exit Sub
    If mytablexb.RecordCount = 0 Then Exit Sub
    mytablexb.MoveFirst
    buf = ""
    Do

        If mytablexb.EOF Then Exit Do
        If Len("" & mytablexb.Fields("producto")) > 0 Then
            buf = buf & "/" & mytablexb.Fields("productop") & "(" & mytablexb.Fields("cantidad") & ")"

        End If

        mytablexb.MoveNext
    Loop
    observa = buf

    If MsgBox("Desea Grabar ", 1, "Aviso") <> 1 Then Exit Sub
    tptovta.Data2.Recordset.Edit
    tptovta.Data2.Recordset.Fields("observa4") = observa
    tptovta.Data2.Recordset.Update
    Label9_Click
    Exit Sub
cmd3_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label9_Click()
    tcombina.Hide
    Unload tcombina

End Sub

Sub sql_combina()

    Dim buf As String

    buf = "select * from " & "_c" & gusuario & " where producto='" & producto & "' order by gcombina"

    If mytablexb.State = 1 Then
        mytablexb.Close
        Set mytablexb = Nothing

    End If

    mytablexb.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = mytablexb

End Sub

Function sql_existe(buf1 As String)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    buf = "select * from " & "_c" & gusuario & " where producto='" & producto & "' and productop='" & buf1 & "' order by gcombina"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sql_existe = 1

    End If

    mytablex.Close

End Function

Private Sub numeros_Click(Index As Integer)

    If Index = 10 Then
        cantidad = ""
        Exit Sub

    End If

    cantidad = cantidad & numeros(Index).Caption

End Sub
