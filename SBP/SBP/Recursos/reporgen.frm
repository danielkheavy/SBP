VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form reporgen 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Reportes"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   14880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modifica"
      Height          =   5535
      Left            =   2880
      TabIndex        =   35
      Top             =   840
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton Command7 
         Caption         =   "&Ignorar"
         Height          =   615
         Left            =   4920
         TabIndex        =   40
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Grabar"
         Height          =   615
         Left            =   3600
         TabIndex        =   39
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox operacion 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1920
         Width           =   3975
      End
      Begin VB.TextBox tamano 
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   37
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox campo2 
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   36
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label tipo 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   47
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Operacion"
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   240
         TabIndex        =   45
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tamaño"
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Campo2"
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label campo1 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   42
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Campo1"
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criterio Busqueda"
      Height          =   4455
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   14535
      Begin VB.CommandButton Command13 
         Caption         =   "Ayuda"
         Height          =   615
         Left            =   10320
         TabIndex        =   49
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Ejecutar"
         Height          =   615
         Left            =   11640
         TabIndex        =   48
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Close"
         Height          =   615
         Left            =   12840
         TabIndex        =   34
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox elcriterio 
         Height          =   2415
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   480
         Width           =   14175
      End
      Begin VB.Label Label9 
         Height          =   1455
         Left            =   240
         TabIndex        =   50
         Top             =   2880
         Width           =   9735
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&VerCriterio"
      Height          =   375
      Left            =   9720
      TabIndex        =   31
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&GenerarExcell"
      Height          =   495
      Left            =   12720
      TabIndex        =   30
      Top             =   3720
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   4335
      Left            =   120
      TabIndex        =   29
      Top             =   4320
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   22
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
   Begin VB.CommandButton Command8 
      Caption         =   "Todos->"
      Height          =   615
      Left            =   3000
      TabIndex        =   28
      Top             =   1200
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   1815
      Left            =   4440
      TabIndex        =   26
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Excell"
      Height          =   375
      Left            =   11040
      TabIndex        =   25
      Top             =   120
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   12120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Ejecutar"
      Height          =   375
      Left            =   9720
      TabIndex        =   23
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Ignorar"
      Height          =   375
      Left            =   9720
      TabIndex        =   22
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Modifica"
      Height          =   375
      Left            =   9720
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo12 
      Height          =   315
      Left            =   7320
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3960
      Width           =   2055
   End
   Begin VB.ComboBox Combo11 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3960
      Width           =   2055
   End
   Begin VB.ComboBox Combo10 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3960
      Width           =   2055
   End
   Begin VB.ComboBox Combo9 
      Height          =   315
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   3120
      Width           =   2055
   End
   Begin VB.ComboBox Combo8 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3120
      Width           =   2055
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3120
      Width           =   2055
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7320
      TabIndex        =   10
      Top             =   2760
      Width           =   2055
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2760
      Width           =   2055
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2760
      Width           =   2055
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   9480
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   2400
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2400
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Borrar"
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Po&ner-->"
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox adicional 
      Height          =   375
      Left            =   5160
      TabIndex        =   51
      Top             =   3480
      Width           =   5415
   End
   Begin VB.Label activado 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3000
      TabIndex        =   27
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Refresca por defecto"
      Height          =   375
      Left            =   5160
      TabIndex        =   24
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Criterio"
      Height          =   375
      Left            =   3000
      TabIndex        =   20
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ordenar"
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label NAMETABLA 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin VB.Menu dlfdw 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "reporgen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rreport As New ADODB.Recordset

Dim tdiseno As New ADODB.Recordset

Dim mysnapx As New ADODB.Recordset

Private Sub Command1_Click()

    Dim found As Integer

    On Error GoTo cmd1_err

    If Len("" & List1.List(List1.ListIndex)) = 0 Then Exit Sub
    found = busca_registro("" & List1.List(List1.ListIndex))

    If found = 1 Then
        MsgBox "Registro ya Seleccionado", 24, "Aviso"
        Exit Sub

    End If

    ir_ultimo
    tdiseno.AddNew
    tdiseno.Fields("campo1") = List1.List(List1.ListIndex)
    tdiseno.Fields("campo2") = List1.List(List1.ListIndex)
    tdiseno.Fields("tipo") = "" & tdiseno.Fields("tipo").Type
    tdiseno.Fields("tamano") = Val("" & rreport.Fields(List1.ListIndex).DefinedSize)
    tdiseno.Update
    ir_ultimo

    Exit Sub
cmd1_err:

    If Err <> 3260 And Err <> 3186 And Err <> 3187 And Err <> 3158 And Err <> 3046 And Err <> 3202 And Err <> 3164 And Err <> 3188 And Err <> 3218 And Err <> 3006 And Err <> 3197 And Err <> 3189 Then
        MsgBox "Error, al agregar " & error$, 24, "Aviso"
    
        Exit Sub

    End If

    MsgBox mensaje_bloqueo & "ADG " & error(Err), 24, "AVISO DE NO ERROR"
    Resume

End Sub

Sub poner_todos()

    Dim found As Integer

    On Error GoTo cmd13_err

    If Len("" & List1.List(List1.ListIndex)) = 0 Then Exit Sub
    found = busca_registro("" & List1.List(List1.ListIndex))

    If found = 1 Then
        MsgBox "Registro ya Seleccionado", 24, "Aviso"
        Exit Sub

    End If

    ir_ultimo
    tdiseno.AddNew
    tdiseno.Fields("campo1") = List1.List(List1.ListIndex)
    tdiseno.Fields("campo2") = List1.List(List1.ListIndex)
    tdiseno.Fields("tipo") = "" & tdiseno.Fields("tipo").Type
    tdiseno.Fields("tamano") = Val("" & rreport.Fields(List1.ListIndex).DefinedSize)
    tdiseno.Update
    ir_ultimo

    Exit Sub
cmd13_err:

    If Err <> 3260 And Err <> 3186 And Err <> 3187 And Err <> 3158 And Err <> 3046 And Err <> 3202 And Err <> 3164 And Err <> 3188 And Err <> 3218 And Err <> 3006 And Err <> 3197 And Err <> 3189 Then
        MsgBox "Error, al agregar " & error$, 24, "Aviso"
    
        Exit Sub

    End If

    MsgBox mensaje_bloqueo & "ADG " & error(Err), 24, "AVISO DE NO ERROR"
    Resume

End Sub

Sub ir_ultimo()

    On Error GoTo cmd34_err

    tdiseno.MoveFirst
    Exit Sub
cmd34_err:
    Exit Sub

End Sub

Private Sub Command10_Click()

    Dim found    As Long

    Dim buf      As String

    Dim buf1     As String

    Dim sw       As Integer

    Dim contador As Integer

    sw = 0
    found = numero_registro()

    If found = 0 Then
        MsgBox "NO existe campos", 24, "Aviso"
        Exit Sub

    End If

    If Combo1.List(Combo1.ListIndex) <> "%" And Combo2.List(Combo2.ListIndex) <> "%" Then
        sw = 1
        buf = "" & Combo1.List(Combo1.ListIndex)
        buf = buf & poner_signo(Combo2.List(Combo2.ListIndex))
        buf = buf & "" & Text1.Text

    End If

    If Combo3.List(Combo3.ListIndex) <> "%" Then
        If Combo4.List(Combo4.ListIndex) <> "%" And Combo5.List(Combo5.ListIndex) <> "%" Then
            buf = buf & poner_signo(Combo3.List(Combo3.ListIndex))
            buf = buf & "" & Combo4.List(Combo4.ListIndex)
            buf = buf & poner_signo(Combo5.List(Combo5.ListIndex))
            buf = buf & "" & Text2.Text

        End If

    End If

    If Combo3.List(Combo3.ListIndex) <> "%" And Combo6.List(Combo6.ListIndex) <> "%" Then
        If Combo7.List(Combo7.ListIndex) <> "%" And Combo8.List(Combo8.ListIndex) <> "%" Then
            buf = buf & poner_signo(Combo6.List(Combo6.ListIndex))
            buf = buf & "" & Combo7.List(Combo7.ListIndex)
            buf = buf & poner_signo(Combo8.List(Combo8.ListIndex))
            buf = buf & "" & Text3.Text

        End If

    End If

    buf1 = "select "
    ir_ultimo
    contador = 0
    Do

        If tdiseno.EOF Then Exit Do
        If contador > 0 Then
            buf1 = buf1 & ","

        End If

        buf1 = buf1 & " " & tdiseno.Fields("campo1")
        contador = contador + 1
        tdiseno.MoveNext
    Loop
    buf1 = buf1 & " from  " & NAMETABLA

    If sw = 1 Then   'si hay datos
        buf1 = buf1 & " where  " & buf

    End If

    'MsgBox Combo10
    buf1 = buf1 & " " & Trim(adicional)

    If Combo10 <> "%" Then
        buf1 = buf1 & " order by " & Combo10

        If Combo11 <> "%" Then
            buf1 = buf1 & " , " & Combo11

        End If

        If Combo12 <> "%" Then
            buf1 = buf1 & " , " & Combo12

        End If

    End If

    Frame1.Visible = True
    elcriterio = buf1
    elcriterio.SetFocus

End Sub

Private Sub Command11_Click()
    Frame1.Visible = False

End Sub

Private Sub Command12_Click()

    If Len(Trim(elcriterio)) = 0 Then Exit Sub
    casillas elcriterio

    'Frame1.Visible = False
End Sub

Private Sub Command13_Click()
    Label9 = "SELECT [ALL | DISTINCT ]"
    Label9 = Label9 + "             <nombre_campo> [{,<nombre_campo>}]"
    Label9 = Label9 + " FROM <nombre_tabla>|<nombre_vista>"
    Label9 = Label9 + "         [{,<nombre_tabla>|<nombre_vista>}]"
    Label9 = Label9 + " [WHERE <condicion> [{ AND|OR <condicion>}]]"
    Label9 = Label9 + " [GROUP BY <nombre_campo> [{,<nombre_campo >}]]"
    Label9 = Label9 + " [HAVING <condicion>[{ AND|OR <condicion>}]]"
    Label9 = Label9 + " [ORDER BY <nombre_campo>|<indice_campo> [ASC | DESC]"
    Label9 = Label9 + "                 [{,<nombre_campo>|<indice_campo> [ASC | DESC ]}]]"

End Sub

Private Sub Command2_Click()

    On Error GoTo cmd78_err

    If tdiseno.RecordCount = 0 Then Exit Sub
    tdiseno.Delete
    'consulta_sql
    Exit Sub
cmd78_err:
    MsgBox "Seleccione un dato ", 24, "AVISO DE NO ERROR"
    Resume

End Sub

Private Sub Command3_Click()

    If tdiseno.RecordCount = 0 Then Exit Sub
    Frame2.Visible = True
    CAMPO1 = "" & tdiseno.Fields("campo1")
    CAMPO2 = "" & tdiseno.Fields("campo2")
    tamano = "" & tdiseno.Fields("tamano")
    tipo = "" & tdiseno.Fields("tipo")
    'operacion = "" & tdiseno.Fields("operacion")
    CAMPO2.SetFocus

End Sub

Private Sub Command4_Click()
    dlfdw_Click

End Sub

Private Sub Command5_Click()

    Dim found    As Long

    Dim buf      As String

    Dim buf1     As String

    Dim sw       As Integer

    Dim contador As Integer

    sw = 0
    found = numero_registro()

    If found = 0 Then
        MsgBox "NO existe campos", 24, "Aviso"
        Exit Sub

    End If

    If Combo1.List(Combo1.ListIndex) <> "%" And Combo2.List(Combo2.ListIndex) <> "%" Then
        sw = 1
        buf = "" & Combo1.List(Combo1.ListIndex)
        MsgBox buf
        buf = buf & poner_signo(Combo2.List(Combo2.ListIndex))
        buf = buf & "" & Text1.Text

    End If

    If Combo3.List(Combo3.ListIndex) <> "%" Then
        If Combo4.List(Combo4.ListIndex) <> "%" And Combo5.List(Combo5.ListIndex) <> "%" Then
            buf = buf & poner_signo(Combo3.List(Combo3.ListIndex))
            buf = buf & "" & Combo4.List(Combo4.ListIndex)
            buf = buf & poner_signo(Combo5.List(Combo5.ListIndex))
            buf = buf & "" & Text2.Text

        End If

    End If

    If Combo3.List(Combo3.ListIndex) <> "%" And Combo6.List(Combo6.ListIndex) <> "%" Then
        If Combo7.List(Combo7.ListIndex) <> "%" And Combo8.List(Combo8.ListIndex) <> "%" Then
            buf = buf & poner_signo(Combo6.List(Combo6.ListIndex))
            buf = buf & "" & Combo7.List(Combo7.ListIndex)
            buf = buf & poner_signo(Combo8.List(Combo8.ListIndex))
            buf = buf & "" & Text3.Text

        End If

    End If

    buf1 = "select "
    ir_ultimo
    contador = 0
    Do

        If tdiseno.EOF Then Exit Do
        If contador > 0 Then
            buf1 = buf1 & ","

        End If

        buf1 = buf1 & " " & tdiseno.Fields("campo1")
        contador = contador + 1
        tdiseno.MoveNext
    Loop
    buf1 = buf1 & " from  " & NAMETABLA

    If sw = 1 Then   'si hay datos
        buf1 = buf1 & " where  " & buf

    End If

    'MsgBox buf1
    buf1 = buf1 & " " & Trim(adicional)
    'MsgBox Combo10

    If Combo10 <> "%" Then
        buf1 = buf1 & " order by " & Combo10

        If Combo11 <> "%" Then
            buf1 = buf1 & " , " & Combo11

        End If

        If Combo12 <> "%" Then
            buf1 = buf1 & " , " & Combo12

        End If

    End If

    'MsgBox buf1
    casillas buf1
    Exit Sub

    '-------------------------------------
    If Check1.Value = 1 Then
        cuerpo_programa_excell buf1
        MsgBox ""
        Exit Sub

    End If

    borrar_archivo globaldir & "\temporal\" & gusuario & ".txt"
    Open globaldir & "\temporal\" & gusuario & ".txt" For Append As #1
    cabecera
    cuerpo_programa buf1
    Close #1
    '---------------------------
    visualiza_datos

End Sub

Sub visualiza_datos()
    'globalrepath = "e:\orion.v6\tmp"
    'globalrepa = "johnny.txt"
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Sub cabecera()

    Dim found As Integer

    Dim buf   As Integer

    tdiseno.MoveFirst
    Do

        If tdiseno.EOF Then Exit Do
        found = formateaa("" & tdiseno.Fields("campo2"), Val("" & tdiseno.Fields("tamano")), 0, 0)
        found = formateaa(" ", 1, 0, 0)
        tdiseno.MoveNext
    Loop
    found = formateaa("", 1, 2, 0)

End Sub

Sub cuerpo_programa(buf As String)

    On Error GoTo cmd7_err

    Dim found As Integer

    If mysnapx.State = 1 Then mysnapx.Close
    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic

    Do

        If mysnapx.EOF Then Exit Do
        '---------------------------
        tdiseno.MoveFirst
        Do

            If tdiseno.EOF Then Exit Do
            found = formateaa("" & mysnapx.Fields("" & tdiseno.Fields("campo1")), Val("" & tdiseno.Fields("tamano")), 0, 0)
            found = formateaa(" ", 1, 0, 0)
            tdiseno.MoveNext
        Loop
        found = formateaa("", 1, 2, 0)
        '---------------------------
        mysnapx.MoveNext
    Loop
    mysnapx.Close

    Exit Sub
cmd7_err:

    'mysnap.Close
End Sub

Function poner_signo(buf As String) As String

    Select Case buf

        Case "Igual"
            poner_signo = "="

        Case "Distinto"
            poner_signo = "<>"

        Case "Mayor"
            poner_signo = ">"

        Case "Menor"
            poner_signo = "<"

        Case "MayorIgual"
            poner_signo = ">="

        Case "MenorIgual"
            poner_signo = "<="

        Case "TodasPosibles"
            poner_signo = " Like "

        Case "Y"
            poner_signo = " and "

        Case "O"
            poner_signo = " or "

    End Select

End Function

Function numero_registro() As Long

    On Error GoTo cmd781_err

    'If tdiseno.EOF Then
    '   Exit Function
    'End If
    tdiseno.MoveLast
    numero_registro = tdiseno.RecordCount
    Exit Function
cmd781_err:
    Exit Function

End Function

Private Sub Command6_Click()

    On Error GoTo cmd3_err

    If Len("" & CAMPO2) > 0 And Val("" & tamano) > 0 Then
        'tdiseno.Edit
        tdiseno.Fields("campo2") = "" & CAMPO2
        tdiseno.Fields("tamano") = Val("" & tamano)
        tdiseno.Update
        Frame2.Visible = False

    End If

    Exit Sub
cmd3_err:
    MsgBox "No se puede Grabar ", 24, "AVISO DE NO ERROR"

End Sub

Private Sub Command7_Click()
    Frame2.Visible = False

End Sub

Private Sub Command9_Click()
    
    Dim xlApp     As Excel.Application

    Dim xlBook    As Excel.Workbook

    Dim xlSheet   As Excel.Worksheet

    Dim sFileName As String

    On Error GoTo PROC_ERR

    'MsgBox "Please format Date column to Date and Time column to time in Excel.", vbInformation, "Message"
    If mysnapx.RecordCount = 0 Then
        MsgBox "No existen Datos ", 48, "Aviso"
        Exit Sub

    End If
    
    sFileName = App.path & "\Time Log as of " & CStr(Format(Now, "mm-dd-yyyy")) & ".xls"

    ExportRecordSetToExcel mysnapx, sFileName, "", "TimeLog"

    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(sFileName)
    xlApp.Application.Visible = True

PROC_EXIT:
    Set mysnapx = Nothing
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Exit Sub

PROC_ERR:
    MsgBox "Primero Ejecutar: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub

Private Sub dlfdw_Click()
    reporgen.Hide
    Unload reporgen

End Sub

Private Sub Form_Activate()
   
    Frame1.Top = 0: Frame1.Left = 0
   
    Dim cad As String

    If activado <> "S" Then
        cad = "SELECT * FROM " & NAMETABLA & " WHERE 1=2"

        If rreport.State = 1 Then rreport.Close
        rreport.Open cad, cn, adOpenStatic, adLockOptimistic
        abre_tabla
        consulta_sql
        activado = "S"

    End If

End Sub

Function busca_registro(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from reporte where campo1='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_registro = 1

    End If

    mytablex.Close

End Function

Sub consulta_sql()

    If tdiseno.State = 1 Then tdiseno.Close
    tdiseno.Open "select * from reporte ", cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = tdiseno
    pone_tamano

End Sub

Sub pone_tamano()
    Exit Sub
    dbGrid1.columns(0).Width = 2000
    dbGrid1.columns(1).Width = 1000
    dbGrid1.columns(2).Width = 1000
    dbGrid1.columns(3).Width = 1000
    dbGrid1.columns(4).Width = 1000

End Sub

Sub abre_tabla()

    Dim I   As Integer

    Dim cad As String
   
    List1.Clear  'nombre tabla
    Combo1.Clear 'como debe salir el nombre
    Combo2.Clear
    Combo2.AddItem "%"
    Combo2.AddItem "Igual"
    Combo2.AddItem "Distinto"
    Combo2.AddItem "Mayor"
    Combo2.AddItem "Menor"
    Combo2.AddItem "MayorIgual"
    Combo2.AddItem "MenorIgual"
    Combo5.AddItem "%"
    Combo5.AddItem "Igual"
    Combo5.AddItem "Distinto"
    Combo5.AddItem "Mayor"
    Combo5.AddItem "Menor"
    Combo5.AddItem "MayorIgual"
    Combo5.AddItem "MenorIgual"
    Combo8.AddItem "%"
    Combo8.AddItem "Igual"
    Combo8.AddItem "Distinto"
    Combo8.AddItem "Mayor"
    Combo8.AddItem "Menor"
    Combo8.AddItem "MayorIgual"
    Combo8.AddItem "MenorIgual"
    Combo3.Clear
    Combo4.Clear '
    Combo7.Clear
    Combo10.Clear
    Combo11.Clear
    Combo12.Clear
   
    Combo10.AddItem "%"
    Combo11.AddItem "%"
    Combo12.AddItem "%"

    Combo1.AddItem "%"
    Combo4.AddItem "%"
    Combo7.AddItem "%"
   
    Combo3.AddItem "%"
    Combo6.AddItem "%"
    Combo9.AddItem "%"
   
    Combo3.AddItem "Y"
    Combo6.AddItem "Y"
    Combo9.AddItem "Y"
   
    Combo3.AddItem "O"
    Combo6.AddItem "O"
    Combo9.AddItem "O"
   
    Combo2.AddItem "TodasPosibles"
    Combo5.AddItem "TodasPosibles"
    Combo8.AddItem "TodasPosibles"

    'MsgBox Trim(rreport.Fields(0).DefinedSize)
    For I = 0 To rreport.Fields.count - 1
        List1.AddItem Trim(rreport.Fields(I).Name)
        Combo1.AddItem Trim(rreport.Fields(I).Name)
        Combo4.AddItem Trim(rreport.Fields(I).Name)
        Combo7.AddItem Trim(rreport.Fields(I).Name)
        Combo10.AddItem Trim(rreport.Fields(I).Name)
        Combo11.AddItem Trim(rreport.Fields(I).Name)
        Combo12.AddItem Trim(rreport.Fields(I).Name)
    Next I

    Label8_Click

End Sub

Private Sub Form_Load()
    cn.Execute ("delete from reporte")

End Sub

Private Sub Label8_Click()
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
    Combo4.ListIndex = 0
    Combo5.ListIndex = 0
    Combo6.ListIndex = 0
    Combo7.ListIndex = 0
    Combo8.ListIndex = 0
    Combo9.ListIndex = 0
    Combo10.ListIndex = 0
    Combo11.ListIndex = 0
    Combo12.ListIndex = 0

End Sub

Private Sub List1_DblClick()
    Command1_Click

End Sub

Sub cuerpo_programa_excell(buf As String)

    Dim v, h As Long

    Dim Heading(80) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim I           As Integer

    Dim j           As Long

    On Error GoTo cmd7_err

    Dim found As Integer

    I = 0
    'cabecera
    tdiseno.MoveFirst
    Do

        If tdiseno.EOF Then Exit Do
        I = I + 1
        Heading(I) = "" & tdiseno.Fields("campo2")
        tdiseno.MoveNext
    Loop

    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(I, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    v = 5
    h = 1

    If mysnapx.State = 1 Then mysnapx.Close
    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic
    j = 0
    Do

        If mysnapx.EOF Then Exit Do
        '---------------------------
        tdiseno.MoveFirst
        Do

            If tdiseno.EOF Then Exit Do
            objExcel.ActiveSheet.Cells(v, h + j) = "'" & mysnapx.Fields("" & tdiseno.Fields("campo1"))
            j = j + 1
            tdiseno.MoveNext
        Loop
        j = 0
        v = v + 1
        mysnapx.MoveNext
    Loop
    mysnapx.Close
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd7_err:

    'mysnap.Close
End Sub

Sub casillas(buf As String)

    On Error GoTo cmd9012_err

    If mysnapx.State = 1 Then mysnapx.Close
    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = mysnapx
    DBGrid2.refresh
    Exit Sub
cmd9012_err:
    MsgBox "Formato Consulta no Valido " + error$, 48, "Aviso"
    Exit Sub
 
End Sub
