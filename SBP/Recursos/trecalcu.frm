VERSION 5.00
Begin VB.Form trecalcu 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recalculos"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox local1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox bodega 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Proceso Terminado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox fechaf 
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
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox fechai 
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
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox producto 
      Enabled         =   0   'False
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
      Left            =   2040
      MaxLength       =   15
      TabIndex        =   1
      Text            =   "%"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ComboBox familia 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Almacen"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label contador 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label productop 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      Caption         =   "Estado Proceso"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInv.Inicial"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Menu proki23 
      Caption         =   "&Procesar"
   End
   Begin VB.Menu lo3434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trecalcu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bodega_Click()

    Dim found As Integer

    found = busca_parame(extra_loquesea(bodega))

End Sub

Private Sub bodega_DblClick()

    Dim found As Integer

    found = busca_parame(extra_loquesea(bodega))

End Sub

Private Sub bodega_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    found = busca_parame(extra_loquesea(bodega))

End Sub

Private Sub bodega_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    found = busca_parame(extra_loquesea(bodega))

End Sub

Private Sub bodega_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    found = busca_parame(extra_loquesea(bodega))

End Sub

Private Sub Command2_Click()
    Command2.Visible = False

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from familia", cn, adOpenStatic, adLockOptimistic
    familia.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        familia.AddItem "" & mytablex.Fields("familia")
        mytablex.MoveNext
    Loop
    familia.ListIndex = 0
    mytablex.Close

    mytablex.Open "select * from bodega", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    bodega.ListIndex = 0
    mytablex.Close

    mytablex.Open "select * from tlocal", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

    mytablex.Close
 
    'found = busca_parame(extra_loquesea(bodega))
    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")

End Sub

Function busca_parame(buf As String)

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    fechai = ""
    'MsgBox buf
    mytablex.Open "select * from bodega where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        fechai = "" & mytablex.Fields("fecha")
        busca_parame = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Private Sub lo3434_Click()

    If Command2.Visible = True Then
        Command2.Visible = False
        Exit Sub

    End If

    trecalcu.Hide
    Unload trecalcu

End Sub

Private Sub proki23_Click()

    Dim found     As Integer

    Dim mytablex  As New ADODB.Recordset

    Dim mytablez  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    Dim mytableb  As New ADODB.Recordset

    Dim buf       As String

    Dim signo     As Double

    Dim xt1       As Double

    Dim xt2       As Double

    Dim xt3       As Double

    Dim xt4       As Double

    Dim xt5       As Double

    Dim xt6       As Double

    Dim xt7       As Double

    Dim xt8       As Double

    Dim xt9       As Double

    Dim xt10      As Double

    Dim xt11      As Double

    Dim xt12      As Double

    Dim xt13      As Double

    Dim xt14      As Double

    Dim xt15      As Double

    Dim xt16      As Double

    Dim mytablera As New ADODB.Recordset

    'Dim found As Integer
    Dim vr

    Dim sdxt As Double

    If Command2.Visible = True Then Exit Sub

    If Len(fechai) = 0 Then
        bodega.SetFocus
        Exit Sub

    End If

    If Len(fechai) <> 10 Then
        bodega.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechai) Then
        bodega.SetFocus
        Exit Sub

    End If

    Command2.Visible = True

    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    If CVDate(fechaf) < CVDate(fechai) Then Exit Sub
    'sql producto
    actualiza_kardex
    Exit Sub
    sdxt = 0
    buf = "select * from producto where descripcio like '%'"

    If producto <> "%" Then
        buf = buf & " and producto like '" & producto & "'"

    End If

    If familia <> "%" Then
        buf = buf & " and familia like '" & familia & "'"

    End If

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    '----primero borrando los datos----
    suma1 = 0
    Do

        If mytablex.EOF Then Exit Do
        suma1 = suma1 + 1
        contador = Format(suma1, "0")
        vr = DoEvents()

        If Command2.Visible = False Then Exit Do
        saldoini = 0
        xt1 = 0
        xt2 = 0
        xt3 = 0
        xt4 = 0
        xt5 = 0
        xt6 = 0
        xt7 = 0
        xt8 = 0
        xt9 = 0
        xt10 = 0
        xt11 = 0
        xt12 = 0
        xt13 = 0
        xt14 = 0
        xt15 = 0
        xt16 = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "Select * from dsaldoini where local='" & extra_loquesea(local1) & "' and producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "' and fecha='" & Format(fechai, "YYYYMMDD") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            saldoini = Val("" & mytablez.Fields("cantidad")) * Val("" & mytablez.Fields("factor"))
            xt1 = Val("" & mytablez.Fields("t1"))
            xt2 = Val("" & mytablez.Fields("t2"))
            xt3 = Val("" & mytablez.Fields("t3"))
            xt4 = Val("" & mytablez.Fields("t4"))
            xt5 = Val("" & mytablez.Fields("t5"))
            xt6 = Val("" & mytablez.Fields("t6"))
            xt7 = Val("" & mytablez.Fields("t7"))
            xt8 = Val("" & mytablez.Fields("t8"))
            xt9 = Val("" & mytablez.Fields("t9"))
            xt10 = Val("" & mytablez.Fields("t10"))
            xt11 = Val("" & mytablez.Fields("t11"))
            xt12 = Val("" & mytablez.Fields("t12"))
            xt13 = Val("" & mytablez.Fields("t13"))
            xt14 = Val("" & mytablez.Fields("t14"))
            xt15 = Val("" & mytablez.Fields("t15"))
            xt16 = Val("" & mytablez.Fields("t16"))

        End If

        mytablez.Close

        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            'mytabley.Edit
            sdxt = sdxt + saldoini
            mytabley.Fields("saldo") = saldoini
            mytabley.Fields("t1") = xt1
            mytabley.Fields("t2") = xt2
            mytabley.Fields("t3") = xt3
            mytabley.Fields("t4") = xt4
            mytabley.Fields("t5") = xt5
            mytabley.Fields("t6") = xt6
            mytabley.Fields("t7") = xt7
            mytabley.Fields("t8") = xt8
            mytabley.Fields("t9") = xt9
            mytabley.Fields("t10") = xt10
            mytabley.Fields("t11") = xt11
            mytabley.Fields("t12") = xt12
            mytabley.Fields("t13") = xt13
            mytabley.Fields("t14") = xt14
            mytabley.Fields("t15") = xt15
            mytabley.Fields("t16") = xt16
            mytabley.Update
        Else
            mytabley.AddNew
            sdxt = sdxt + saldoini
            mytabley.Fields("local") = extra_loquesea(local1)
            mytabley.Fields("producto") = "" & mytablex.Fields("producto")
            mytabley.Fields("bodega") = extra_loquesea(bodega)
            mytabley.Fields("saldo") = saldoini
            mytabley.Fields("t1") = xt1
            mytabley.Fields("t2") = xt2
            mytabley.Fields("t3") = xt3
            mytabley.Fields("t4") = xt4
            mytabley.Fields("t5") = xt5
            mytabley.Fields("t6") = xt6
            mytabley.Fields("t7") = xt7
            mytabley.Fields("t8") = xt8
            mytabley.Fields("t9") = xt9
            mytabley.Fields("t10") = xt10
            mytabley.Fields("t11") = xt11
            mytabley.Fields("t12") = xt12
            mytabley.Fields("t13") = xt13
            mytabley.Fields("t14") = xt14
            mytabley.Fields("t15") = xt15
            mytabley.Fields("t16") = xt16
            mytabley.Update

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close

    'ahora ver las transacciones y sumarlos al saldo
    buf = "select * from detalle where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If familia <> "%" Then
        buf = buf & " and familia like '" & familia & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and producto like '" & producto & "'"

    End If

    buf = buf & " and local='" & extra_loquesea(local1) & "'"
    buf = buf & " and bodega='" & extra_loquesea(bodega) & "'"
    buf = buf & " and (acu='S' or acu='T' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' OR acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='E')"
    'buf = buf & " and (len(acu1)=0 or acu1=null)"
    buf = buf & " and estado='2'"
    'buf = buf & " group by producto,bodega,flage,acu1"

    If mytableb.State = 1 Then mytableb.Close
    mytableb.Open buf, cn, adOpenStatic, adLockOptimistic
    suma1 = 0

    If Command2.Visible = False Then Exit Sub
    Do
        vr = DoEvents()

        If Command2.Visible = False Then Exit Do
        If mytableb.EOF Then Exit Do

        'aqui validamos si se puede actualizar
        If mytablera.State = 1 Then mytablera.Close
        mytablera.Open "select tipo1 from factura where local='" & "" & mytableb.Fields("local") & "' and tipo='" & "" & mytableb.Fields("tipo") & "' and serie='" & mytableb.Fields("serie") & "' and numero='" & "" & mytableb.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablera.RecordCount > 0 Then
            found = ve_descarga("" & mytablera.Fields("tipo1"))

            If found = 1 Then 'qe no se descarge
                mytablera.Close
                GoTo sisin

            End If

        End If

        mytablera.Close

        suma1 = suma1 + 1
        contador = Format(suma1, "0")
        productop = "" & mytableb.Fields("producto")
        signo = 1

        If "" & mytableb.Fields("acu") = "T" Or "" & mytableb.Fields("acu") = "A" Or "" & mytableb.Fields("acu") = "B" Or "" & mytableb.Fields("acu") = "C" Or "" & mytableb.Fields("acu") = "D" Or "" & mytableb.Fields("acu") = "G" Or "" & mytableb.Fields("acu") = "N" Then
            signo = -1

        End If

        If "" & mytableb.Fields("acu") = "S" Or "" & mytableb.Fields("acu") = "J" Or "" & mytableb.Fields("acu") = "K" Or "" & mytableb.Fields("acu") = "L" Or "" & mytableb.Fields("acu") = "M" Or "" & mytableb.Fields("acu") = "P" Or "" & mytableb.Fields("acu") = "E" Then
            signo = 1

        End If

        'If Val("" & mytableb.Fields("cantidad")) < 0 Then
        '   signo = 1
        'End If
        'ahora en almacenes
        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open "Select * from almacen where local='" & extra_loquesea(local1) & "' and producto='" & "" & mytableb.Fields("producto") & "' and bodega='" & "" & mytableb.Fields("bodega") & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            'mytabley.Edit
            sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytableb.Fields("cantidad")) * Val("" & mytableb.Fields("factor"))
            mytabley.Fields("saldo") = sdx

            'sdxt = sdxt + sdx
            If Len("" & mytableb.Fields("linea")) > 0 Then
                sdx = Val("" & mytabley.Fields("T1")) + signo * Val("" & mytableb.Fields("T1"))
                mytabley.Fields("t1") = sdx
                sdx = Val("" & mytabley.Fields("T2")) + signo * Val("" & mytableb.Fields("T2"))
                mytabley.Fields("t2") = sdx
                sdx = Val("" & mytabley.Fields("T3")) + signo * Val("" & mytableb.Fields("T3"))
                mytabley.Fields("t3") = sdx
                sdx = Val("" & mytabley.Fields("T4")) + signo * Val("" & mytableb.Fields("T4"))
                mytabley.Fields("t4") = sdx
                sdx = Val("" & mytabley.Fields("T5")) + signo * Val("" & mytableb.Fields("T5"))
                mytabley.Fields("t5") = sdx
                sdx = Val("" & mytabley.Fields("T6")) + signo * Val("" & mytableb.Fields("T6"))
                mytabley.Fields("t6") = sdx
                sdx = Val("" & mytabley.Fields("T7")) + signo * Val("" & mytableb.Fields("T7"))
                mytabley.Fields("t7") = sdx
                sdx = Val("" & mytabley.Fields("T8")) + signo * Val("" & mytableb.Fields("T8"))
                mytabley.Fields("t8") = sdx
                sdx = Val("" & mytabley.Fields("T9")) + signo * Val("" & mytableb.Fields("T9"))
                mytabley.Fields("t9") = sdx
                sdx = Val("" & mytabley.Fields("T10")) + signo * Val("" & mytableb.Fields("T10"))
                mytabley.Fields("t10") = sdx
                sdx = Val("" & mytabley.Fields("T11")) + signo * Val("" & mytableb.Fields("T11"))
                mytabley.Fields("t11") = sdx
                sdx = Val("" & mytabley.Fields("T12")) + signo * Val("" & mytableb.Fields("T12"))
                mytabley.Fields("t12") = sdx
                sdx = Val("" & mytabley.Fields("T13")) + signo * Val("" & mytableb.Fields("T13"))
                mytabley.Fields("t13") = sdx
                sdx = Val("" & mytabley.Fields("T14")) + signo * Val("" & mytableb.Fields("T14"))
                mytabley.Fields("t14") = sdx
                sdx = Val("" & mytabley.Fields("T15")) + signo * Val("" & mytableb.Fields("T15"))
                mytabley.Fields("t15") = sdx
                sdx = Val("" & mytabley.Fields("T16")) + signo * Val("" & mytableb.Fields("T16"))
                mytabley.Fields("t16") = sdx

            End If

            mytabley.Update
        Else
            mytabley.AddNew
            mytabley.Fields("producto") = "" & mytableb.Fields("producto")
            mytabley.Fields("bodega") = "" & mytableb.Fields("bodega")
            mytabley.Fields("local") = "" & mytableb.Fields("local")
            sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytableb.Fields("cantidad")) * Val("" & mytableb.Fields("factor"))
            mytabley.Fields("saldo") = sdx

            'sdxt = sdxt + sdx
            If Len("" & mytableb.Fields("linea")) > 0 Then
                sdx = Val("" & mytabley.Fields("T1")) + signo * Val("" & mytableb.Fields("T1"))
                mytabley.Fields("t1") = sdx
                sdx = Val("" & mytabley.Fields("T2")) + signo * Val("" & mytableb.Fields("T2"))
                mytabley.Fields("t2") = sdx
                sdx = Val("" & mytabley.Fields("T3")) + signo * Val("" & mytableb.Fields("T3"))
                mytabley.Fields("t3") = sdx
                sdx = Val("" & mytabley.Fields("T4")) + signo * Val("" & mytableb.Fields("T4"))
                mytabley.Fields("t4") = sdx
                sdx = Val("" & mytabley.Fields("T5")) + signo * Val("" & mytableb.Fields("T5"))
                mytabley.Fields("t5") = sdx
                sdx = Val("" & mytabley.Fields("T6")) + signo * Val("" & mytableb.Fields("T6"))
                mytabley.Fields("t6") = sdx
                sdx = Val("" & mytabley.Fields("T7")) + signo * Val("" & mytableb.Fields("T7"))
                mytabley.Fields("t7") = sdx
                sdx = Val("" & mytabley.Fields("T8")) + signo * Val("" & mytableb.Fields("T8"))
                mytabley.Fields("t8") = sdx
                sdx = Val("" & mytabley.Fields("T9")) + signo * Val("" & mytableb.Fields("T9"))
                mytabley.Fields("t9") = sdx
                sdx = Val("" & mytabley.Fields("T10")) + signo * Val("" & mytableb.Fields("T10"))
                mytabley.Fields("t10") = sdx
                sdx = Val("" & mytabley.Fields("T11")) + signo * Val("" & mytableb.Fields("T11"))
                mytabley.Fields("t11") = sdx
                sdx = Val("" & mytabley.Fields("T12")) + signo * Val("" & mytableb.Fields("T12"))
                mytabley.Fields("t12") = sdx
                sdx = Val("" & mytabley.Fields("T13")) + signo * Val("" & mytableb.Fields("T13"))
                mytabley.Fields("t13") = sdx
                sdx = Val("" & mytabley.Fields("T14")) + signo * Val("" & mytableb.Fields("T14"))
                mytabley.Fields("t14") = sdx
                sdx = Val("" & mytabley.Fields("T15")) + signo * Val("" & mytableb.Fields("T15"))
                mytabley.Fields("t15") = sdx
                sdx = Val("" & mytabley.Fields("T16")) + signo * Val("" & mytableb.Fields("T16"))
                mytabley.Fields("t16") = sdx

            End If

            mytabley.Update

        End If

sisin:
        mytableb.MoveNext
    Loop
    'MsgBox sdxt
    mytableb.Close
    mytabley.Close
    Command2.Visible = False

End Sub

Function ve_descarga(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe

        Select Case "" & mytablex.Fields("tipodoc")

            Case "A", "B", "C", "D", "G", "E", "F", "J", "K", "L", "M", "P", "N", "O", "S", "T"
                ve_descarga = 1

        End Select

    End If

    mytablex.Close

End Function

Function actualiza_kardex()

    Dim found As Integer

    found = kardexactualiza(extra_loquesea(local1), "" & producto, extra_loquesea(bodega), fechai, fechaf)

End Function

