VERSION 5.00
Begin VB.Form tconecta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conectividad"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   8670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label procesos 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label estado1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Elija un Recurso"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu dloo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tconecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cnn      As New ADODB.Connection

Dim cnremote As New ADODB.Connection

Private Sub Command1_Click()

    Dim vr

    Dim found    As Integer

    Dim AdoCmd   As New ADODB.Command

    Dim ufile    As String

    Dim dfile    As String

    Dim SQL      As String

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    On Error GoTo ErrorHandler

    If Combo1 = "EnvioProductos" Then
        If InputBox("Llave de Paso ", "Control", "") <> "CUIDADO" Then Exit Sub
        found = conecta_productos()

        If found = 0 Then
            MsgBox "No existe conexion ", 48, "Aviso"
            Exit Sub

        End If

        MsgBox "Conexion establecia"
        procesos = "Familia"
        vr = DoEvents
   
        mytablex.Open "select * from familia where familia like '%'", cnn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Sub

        End If
   
        cn.Execute ("delete from familia")
        mytabley.Open "select * from familia where familia like '%'", cn, adOpenStatic, adLockOptimistic
   
        Do

            If mytablex.EOF Then Exit Do
   
            mytabley.AddNew

            For I = 0 To mytablex.Fields.count - 1
                mytabley.Fields(I) = mytablex.Fields(I)
            Next I

            mytabley.Update
            mytablex.MoveNext
        Loop
        mytablex.Close
        mytabley.Close
   
        'productos
        procesos = "Producto"
        vr = DoEvents
   
        mytablex.Open "select * from producto where producto like '%'", cnn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Sub

        End If

        cn.Execute ("delete from producto where producto like '%'")
        mytabley.Open "select * from producto where producto like '%'", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablex.EOF Then Exit Do
            mytabley.AddNew

            For I = 0 To mytablex.Fields.count - 1
                mytabley.Fields(I) = mytablex.Fields(I)
            Next I

            mytabley.Update
            mytablex.MoveNext
        Loop
        mytablex.Close
        mytabley.Close
   
        'subfamilia
        procesos = "Subfamilia"
        vr = DoEvents
   
        mytablex.Open "select * from subfamil where subfamilia like '%'", cnn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Sub

        End If

        cn.Execute ("delete from subfamil")
        mytabley.Open "select * from subfamil where subfamilia like '%'", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablex.EOF Then Exit Do
            mytabley.AddNew

            For I = 0 To mytablex.Fields.count - 1
                mytabley.Fields(I) = mytablex.Fields(I)
            Next I

            mytabley.Update
            mytablex.MoveNext
        Loop
        mytablex.Close
        mytabley.Close
   
        'marca
        procesos = "Marca"
        vr = DoEvents
   
        mytablex.Open "select * from marca where marca like '%'", cnn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Sub

        End If

        cn.Execute ("delete from marca")
        mytabley.Open "select * from marca where marca like '%'", cn, adOpenStatic, adLockOptimistic
        Do

            If mytablex.EOF Then Exit Do
            mytabley.AddNew

            For I = 0 To mytablex.Fields.count - 1
                mytabley.Fields(I) = mytablex.Fields(I)
            Next I

            mytabley.Update
            mytablex.MoveNext
        Loop
        mytablex.Close
        mytabley.Close
        cnn.Close
        MsgBox "Proceso Terminado ", 48, "Aviso"
        Exit Sub

    End If

    If Combo1 = "CopiaSeguridad" Then
        If MsgBox("Desea Procesar..", 1, "Aviso") <> 1 Then Exit Sub
        'found = conectar_remoto()
        'If found = 0 Then
        '   MsgBox "Coneccion No establecida ", 48, "Aviso"
        '   Exit Sub
        'End If
        '   MsgBox "Coneccion Exitosa,Realizando copia seguridad", 48, "Aviso"
        ufile = Trim(extra_loquesea1(menup.gempresa))
        dfile = globalpath & "\bk" & Format(Now, "ddmmyy") & ".bak"

        If Len(Dir(Trim(dfile))) > 0 Then
            Kill dfile

        End If

        'sql = "BACKUP DATABASE [CMH] TO DISK = '" & Trim(dfile) & "'"
        'cn.Execute sql
        'MsgBox "pase"
        'Exit Sub
   
        Screen.MousePointer = vbHourglass
        strexecute = "BACKUP DATABASE [" & Trim(ufile) & "] "
        strexecute = strexecute & "TO DISK=N'" & Trim(dfile) & "' "
        strexecute = strexecute & "WITH FORMAT,INIT,STATS=10 "

        'MsgBox strexecute
        With AdoCmd
            .ActiveConnection = cnremote
            .CommandType = adCmdText
            .CommandTimeout = 0
            .CommandText = "use " & ufile
            .Execute
            .CommandText = strexecute
            .Execute
            Set .ActiveConnection = Nothing

        End With
    
        Set AdoCmd = Nothing
        MsgBox "Backup Completado.", vbInformation
        Screen.MousePointer = vbDefault

    End If

    Exit Sub
ErrorHandler:
    Set AdoCmd = Nothing
    MsgBox "Aviso en copia " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub dloo33_Click()
    tconecta.Hide
    Unload tconecta

End Sub

Private Sub Form_Load()
    Combo1.AddItem "%"
    Combo1.AddItem "CopiaSeguridad"
    'Combo1.AddItem "EnvioProductos"
    Combo1.ListIndex = 0

End Sub

Public Function conectar_remoto()

    Dim buf As String

    On Error GoTo cmd5454_err

    buf = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=;Data Source=" & Trim(extra_loquesea1(menup.gempresa))

    With cnremote

        If .State = adStateOpen Then .Close
        .ConnectionString = buf
        '        ConnectionTimeout = 30
        '        CursorLocation = adUseClient
        .Open

    End With

    conectar_remoto = 1
    Exit Function
cmd5454_err:
    Exit Function

End Function

Function conecta_productos()

    On Error GoTo cmd8912_err

    cnn.CursorLocation = adUseClient
    cnn.Open "Driver={SQL Server};Server=hh.zapto.org;Database=" & extra_loquesea1(menup.gempresa) & ";uid=sa"
    conecta_productos = 1
    Exit Function
cmd8912_err:
    MsgBox "No se conecta con precios " + error$, 48, "Aviso"
    Exit Function

End Function
 
