VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form tingper 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada Salida Personal"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   5400
      Picture         =   "tingper.frx":0000
      ScaleHeight     =   2475
      ScaleWidth      =   2115
      TabIndex        =   22
      Top             =   3480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtUserID 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8865
      TabIndex        =   5
      Top             =   1620
      Width           =   3150
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   8880
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2115
      Width           =   3150
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10080
      TabIndex        =   4
      Top             =   7200
      Width           =   2670
   End
   Begin VB.CommandButton cmdAdministration 
      Caption         =   "&Administracion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      TabIndex        =   3
      Top             =   7200
      Width           =   2670
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   225
      Top             =   900
   End
   Begin VB.OptionButton optTimeIn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Entrada"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7470
      TabIndex        =   2
      Top             =   765
      Value           =   -1  'True
      Width           =   2445
   End
   Begin VB.OptionButton optTimeOut 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9900
      TabIndex        =   1
      Top             =   765
      Width           =   2445
   End
   Begin ChamaleonButton.ChameleonBtn cmdOK 
      Height          =   705
      Left            =   8880
      TabIndex        =   23
      Top             =   2640
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1244
      BTYPE           =   4
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tingper.frx":25C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   30
      TabIndex        =   10
      Top             =   1485
      Width           =   7305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Control de Asistencia Personal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   15
      TabIndex        =   9
      Top             =   -15
      Width           =   12795
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clave :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   8055
      TabIndex        =   8
      Top             =   2115
      Width           =   735
   End
   Begin VB.Label lblLastLogin 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ultimo Ingreso :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   7
      Top             =   6720
      Width           =   2370
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   855
      Width           =   7305
   End
   Begin VB.Menu flo9343 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tingper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    Dim ssql As String

    Dim rs   As New ADODB.Recordset

    On Error GoTo PROC_ERR

    If Len(txtpassword) = 0 Then Exit Sub
    'set sql query for password lookup
    ssql = "SELECT * FROM vendedor WHERE clave='" & txtpassword & "' and (clavere='1' or clavere='2') "
    'pass result to recordset
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic
    '    MsgBox sSQL, vbInformation, "SQL Query"
        
    If rs.RecordCount > 0 Then
        'save log to tblLogin
        txtUserID = "" & rs.Fields("codigo")
        Call SaveToLog

        'display last user
        If optTimeIn.Value = True Then
            lblLastLogin.Caption = "Ultima Entrada: [" & txtUserID.Text & "] " & rs![nombre] & " (" & lblTime.Caption & ")"
        ElseIf optTimeOut.Value = True Then
            lblLastLogin.Caption = "Ultima Salida: [" & txtUserID.Text & "] " & rs![nombre] & " (" & lblTime.Caption & ")"

        End If
        
        'clear text box entry
        txtUserID.Text = ""
        txtpassword.Text = ""
    Else
        'inform for error
        MsgBox "No existe usuario ", vbExclamation, "Message"

    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub

Private Sub Command1_Click(Index As Integer)

    If Index = 10 Then
        txtpassword = ""
        Exit Sub

    End If

    txtpassword = txtpassword & Command1(Index).Caption

End Sub

Private Sub flo9343_Click()
    tingper.Hide
    Unload tingper

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM parame where  codigo='01'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        If "" & mytablex.Fields("solohuella") = "S" Then
            txtpassword.Enabled = False

        End If

    End If

    mytablex.Close

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 21/10/2006 19:26
' Author    :
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    On Error GoTo PROC_ERR

    'set time and date
    lblDate.Caption = Format(Now, "dd/mm/yyyy")
    lblTime.Caption = Format(Now, "hh:mm:ss")
    
    'toggle color of Time Option
    Call ToggleTimeOption

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub

Private Sub optTimeIn_Click()
    Call ToggleTimeOption

End Sub

Private Sub optTimeOut_Click()
    Call ToggleTimeOption

End Sub

Private Sub Picture1_Click()
    control_huelladigital

End Sub

Private Sub Timer1_Timer()

    On Error GoTo PROC_ERR

    ' This Code Helps for date and time
    lblDate.Caption = Format(Now, "dd/mm/yyyy")
    lblTime.Caption = Format(Now, "hh:mm:ss")
    
PROC_EXIT:
    Exit Sub

PROC_ERR:

    Resume PROC_EXIT

End Sub

Private Sub cmdAdministration_Click()

    Dim rs   As New ADODB.Recordset

    Dim ssql As String

    On Error GoTo PROC_ERR

    'If Len(txtPassword) = 0 Then Exit Sub
    ssql = "SELECT * FROM vendedor WHERE clave = '" & txtpassword.Text & "' and clavere='2' "
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic
    '    MsgBox sSQL, vbInformation, "SQL Query"
    
    If rs.RecordCount > 0 Then
        txtpassword = ""
        tsegper.Show 1
    Else
        MsgBox "Acceso no Autorizado !", vbCritical, "Message"

    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub

Private Sub cmdExit_Click()

    Dim ssql As String

    Dim rs   As New ADODB.Recordset

    On Error GoTo PROC_ERR

    If Len(txtpassword) = 0 Then Exit Sub
    ssql = "SELECT * FROM vendedor WHERE clave = '" & txtpassword.Text & "' and clavere='2' "
    rs.Open ssql, cn, adOpenStatic, adLockOptimistic
    '    MsgBox sSQL, vbInformation, "SQL Query"
    
    If rs.RecordCount > 0 Then
        'end program
        tingper.Hide
        Unload tingper
        Exit Sub
    Else
        MsgBox "Usted no esta autorizado para cerrar el programa !", vbCritical, "Mensage"

    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SaveToLog
' DateTime  : 21/10/2006 19:25
' Author    :
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SaveToLog()

    Dim ssql As String

    Dim rs   As New ADODB.Recordset

    On Error GoTo PROC_ERR

    'set sql query for insert type: Log in or Log out
    rs.Open "select * from ingper where codigo='" & txtUserID.Text & "'", cn, adOpenStatic, adLockOptimistic
        
    If optTimeIn.Value = True Then
        rs.AddNew
        rs.Fields("codigo") = Trim(txtUserID.Text)
        rs.Fields("fecha") = Format(lblDate.Caption, "dd/mm/yyyy")
        rs.Fields("timein") = lblTime.Caption
        rs.Update
        'ssql = "INSERT INTO ingper(codigo,fecha,TimeIn) VALUES ('" & txtUserID.Text & "','" & Format(lblDate.Caption, "dd/mm/yyyy") & "','" & lblTime.Caption & "')"
    ElseIf optTimeOut.Value = True Then
        'ssql = "INSERT INTO ingper(codigo,fecha,TimeOut) VALUES ('" & txtUserID.Text & "','" & Format(lblDate.Caption, "dd/mm/yyyy") & "','" & lblTime.Caption & "')"
        rs.AddNew
        rs.Fields("codigo") = Trim(txtUserID.Text)
        rs.Fields("fecha") = Format(lblDate.Caption, "dd/mm/yyyy")
        rs.Fields("timeout") = lblTime.Caption
        rs.Update

    End If

    '   MsgBox sSQL, vbInformation, "SQL Query"
         
    'cn.Execute ssql

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub

Private Sub ToggleTimeOption()

    On Error GoTo PROC_ERR

    If optTimeIn.Value = True Then
        optTimeIn.ForeColor = vbFontBlack
        optTimeOut.ForeColor = vbFontGray
        ''03/07/2017 KENYO Diseño ocutar opcion  Picture1 de huela. Frm tingper
        ' cmdOK.Caption = "&Entrada"
        ''03/07/2017 KENYO Diseño ocutar opcion  Picture1 de huela. Frm tingper
    ElseIf optTimeOut.Value = True Then
        optTimeOut.ForeColor = vbFontBlack
        optTimeIn.ForeColor = vbFontGray

        ''03/07/2017 KENYO Diseño ocutar opcion  Picture1 de huela. Frm tingper
        'cmdOK.Caption = "&Salida"
        ''03/07/2017 KENYO Diseño ocutar opcion  Picture1 de huela. Frm tingper
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    cmdOK_Click

End Sub

Sub control_huelladigital()

    Dim mytablex As New ADODB.Recordset

    codigohuella = ""
    thuellat.tipo = "personal"
    thuellat.codigo = ""
    thuellat.nombre = ""
    thuellat.Show 1

    'MsgBox codigohuella
    If Len(Trim(codigohuella)) > 0 Then
        mytablex.Open "select * from vendedor where codigo='" & Trim(codigohuella) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            txtpassword = Trim("" & mytablex.Fields("clave"))
            MsgBox txtpassword
            mytablex.Close
            cmdOK_Click
            Exit Sub

        End If

        mytablex.Close

    End If

End Sub
