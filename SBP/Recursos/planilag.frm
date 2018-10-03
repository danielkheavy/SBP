VERSION 5.00
Begin VB.Form planilag 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generacion Planilla"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox periodo 
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   120
      Width           =   4455
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Paso 3"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Paso 2"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Paso 1"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar"
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
      Left            =   6360
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu ldo34 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "planilag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim found As Integer

    Dim sw    As Integer

    If MsgBox("Si existe Proceso Anterior,se Reemplazara por el Actual " & Chr$(10) & Chr$(10) & "Desea Procesar,", 1, "Aviso") <> 1 Then Exit Sub
    found = borrar_periodo()

    If found = 0 Then
        MsgBox "Error al Borrar Proceso ", 48, "Aviso"
        ldo34_Click

    End If

    found = generar_planilla1()

    If found = 1 Then
        Check1.Value = 1

    End If

    found = generar_planilla2()

    If found = 1 Then
        Check2.Value = 1

    End If

    found = generar_planilla3()

    If found = 1 Then
        Check3.Value = 1

    End If

    'found = generar_periodo()
    If Check1.Value = 1 And Check2.Value = 1 And Check3.Value = 1 Then
        MsgBox "Proceso Finalizado ,exitoso ", 48, "Aviso"
        ldo34_Click
        Exit Sub

    End If

    MsgBox "Proceso Finalizado ,Con errores ", 48, "Aviso"
    Exit Sub

End Sub

Function generar_planilla1()

    On Error GoTo cmd67_err

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim xmes     As String

    Dim I        As Integer
   
    mytablex.Open "select * from remune01  ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If
   
    Do

        If mytablex.EOF Then Exit Do
        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open "select * from remune02  where tipopla='" & mytablex.Fields("tipopla") & "' and codigo='" & mytablex.Fields("codigo") & "' and tipo='" & mytablex.Fields("tipo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            mytabley.AddNew
            graba_campo mytablex, mytabley, "" & extra_loquesea(periodo)
            mytabley.Update
        Else
            graba_campo mytablex, mytabley, "" & extra_loquesea(periodo)
            mytabley.Update

        End If

        mytabley.Close
        mytablex.MoveNext
    Loop
    mytablex.Close
    
    generar_planilla1 = 1
    Exit Function
cmd67_err:
    MsgBox "Aviso en generar plantilla1 " + error$, 48, "Aviso"
    
    Exit Function

End Function

Function generar_planilla2()

    On Error GoTo cmd671_err

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim xmes     As String

    Dim I        As Integer
   
    mytablex.Open "select * from descue01  ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If
   
    Do

        If mytablex.EOF Then Exit Do
        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open "select * from descue02  where tipopla='" & mytablex.Fields("tipopla") & "' and codigo='" & mytablex.Fields("codigo") & "' and tipo='" & mytablex.Fields("tipo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            mytabley.AddNew
            graba_campo mytablex, mytabley, "" & extra_loquesea(periodo)
            mytabley.Update
        Else
            graba_campo mytablex, mytabley, "" & extra_loquesea(periodo)
            mytabley.Update

        End If

        mytabley.Close
        mytablex.MoveNext
    Loop
    mytablex.Close
    
    generar_planilla2 = 1
    Exit Function
cmd671_err:
    
    Exit Function

End Function

Sub graba_campo(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset, xbuf As String)

    On Error GoTo cmd8912_err

    mytabley.Fields("tipopla") = "" & mytablex.Fields("tipopla")
    mytabley.Fields("codigo") = "" & mytablex.Fields("codigo")
    mytabley.Fields("tipo") = "" & mytablex.Fields("tipo")
    mytabley.Fields("periodo") = xbuf
    mytabley.Fields("concepto") = "" & mytablex.Fields("concepto")
    mytabley.Fields("importe") = Val("" & mytablex.Fields("importe"))
    Exit Sub
cmd8912_err:
    MsgBox "Aviso en graba Campo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function generar_planilla3()

    On Error GoTo cmd1267_err

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim xmes     As String

    Dim I        As Integer
   
    mytablex.Open "select * from aporta01  ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If
   
    Do

        If mytablex.EOF Then Exit Do
        If mytabley.State = 1 Then mytabley.Close
   
        mytabley.Open "select * from aporta02  where tipopla='" & mytablex.Fields("tipopla") & "' and codigo='" & mytablex.Fields("codigo") & "' and tipo='" & mytablex.Fields("tipo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            mytabley.AddNew
            graba_campo mytablex, mytabley, "" & extra_loquesea(periodo)
            mytabley.Update
        Else
            graba_campo mytablex, mytabley, "" & extra_loquesea(periodo)
            mytabley.Update

        End If

        mytabley.Close
        mytablex.MoveNext
    Loop
    mytablex.Close
    
    generar_planilla3 = 1
    Exit Function
cmd1267_err:
    
    Exit Function

End Function

Private Sub Command2_Click()
    ldo34_Click

End Sub

Private Sub Form_Load()

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    fecha = Format(Now, "dd/mm/yyyy")
    periodo.Clear
    periodo.AddItem "%"
    mytablex.Open "select * from plaperiodo ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            periodo.AddItem Trim("" & mytablex.Fields("periodo")) & "|" & mytablex.Fields("descripcio")
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    periodo.ListIndex = 0

End Sub

Private Sub ldo34_Click()
    planilag.Hide
    Unload planilag

End Sub

Function borrar_periodo()

    On Error GoTo cmd89_err

    cn.Execute ("delete from remune02 where periodo='" & extra_loquesea(periodo) & "'")
    cn.Execute ("delete from descue02 where periodo='" & extra_loquesea(periodo) & "'")
    cn.Execute ("delete from aporta02 where periodo='" & extra_loquesea(periodo) & "'")
    borrar_periodo = 1
    Exit Function
cmd89_err:
 
    Exit Function

End Function

