VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tcaprod 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambia de Productos"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox campose 
      Height          =   495
      Left            =   8520
      MaxLength       =   120
      TabIndex        =   17
      Top             =   600
      Width           =   4575
   End
   Begin VB.ComboBox ordenado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   120
      Width           =   5655
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   6720
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hacer Cambio"
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
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "tcaprod.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6720
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox cambiar 
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   7200
      Width           =   2895
   End
   Begin VB.ComboBox criterio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   480
      Width           =   3855
   End
   Begin VB.ComboBox campo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   3855
   End
   Begin VB.TextBox buscar 
      Height          =   495
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "%"
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton Command21 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "tcaprod.frx":07AE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   4455
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Campos Ver (Opcional)"
      Height          =   495
      Left            =   6960
      TabIndex        =   18
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ordenado"
      Height          =   495
      Left            =   6960
      TabIndex        =   16
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label tabla 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Campo"
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cambiar Por"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Campo"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label21 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Que Buscar"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nota. Por cambios mal realizados no nos hacemos responsables. Debe hacerse por personal que tenga conocimiento de Sistemas."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   4
      Top             =   6720
      Width           =   8055
   End
   Begin VB.Label Label23 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Criterio Busqueda"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Menu m8912 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcaprod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mysnapx As New ADODB.Recordset

Private Sub Command1_Click()

    On Error GoTo cmd9078_err

    If Combo1 = "%" Then Exit Sub

    If MsgBox("Desea realizar el proceso ", 1, "Aviso") <> 1 Then Exit Sub
    If Combo1 = "%" Then Exit Sub
    buf = "update " & tabla & " set " & Combo1 & "='" & cambiar & "'"
    buf = buf & " where "
    buf = buf & "" & campo.List(campo.ListIndex)
    buf = buf & poner_signo(criterio.List(criterio.ListIndex))
    buf = buf & " '" & buscar.Text & "'"
    cn.Execute (buf)
    MsgBox "Proceso Realizado ", 48, "Aviso"
    Exit Sub
cmd9078_err:
    MsgBox "No se realizado en proceso ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command21_Click()

    Dim buf As String

    If campo <> "%" And criterio <> "%" Then
        If Len(Trim(campose)) = 0 Then
            buf = "select * from " & tabla & " where "
        Else
            buf = "select " & campose & " from " & tabla & " where "

        End If

        buf = buf & "" & campo.List(campo.ListIndex)
        buf = buf & poner_signo(criterio.List(criterio.ListIndex))
        buf = buf & " '" & buscar.Text & "'"

        If ordenado <> "%" Then
            buf = buf & " order by " & ordenado

        End If

        casillas buf

    End If

End Sub

Private Sub Form_Activate()
    sql_cabeza
    Command21_Click

End Sub

Private Sub Form_Load()
    criterio.Clear
    criterio.AddItem "%"
    criterio.AddItem "TodasPosibles"
    criterio.AddItem "Igual"
    criterio.AddItem "Distinto"
    criterio.AddItem "Mayor"
    criterio.AddItem "Menor"
    criterio.AddItem "MayorIgual"
    criterio.AddItem "MenorIgual"
    criterio.ListIndex = 1
   
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

Private Sub m8912_Click()
    tcaprod.Hide
    Unload tcaprod

End Sub

Sub sql_cabeza()

    Dim rmytablex As New ADODB.Recordset

    Dim I         As Integer

    Dim buf       As String

    campo.Clear
    campo.AddItem "%"

    Combo1.Clear
    Combo1.AddItem "%"

    ordenado.Clear
    ordenado.AddItem "%"

    buf = "SELECT * FROM " & tabla & " WHERE 1=2"
    rmytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    For I = 0 To rmytablex.Fields.count - 1
        campo.AddItem UCase(Trim(rmytablex.Fields(I).Name))
        Combo1.AddItem UCase(Trim(rmytablex.Fields(I).Name))
        ordenado.AddItem UCase(Trim(rmytablex.Fields(I).Name))
    Next I

    campo.ListIndex = 1
    Combo1.ListIndex = 0
    ordenado.ListIndex = 0
    rmytablex.Close

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
