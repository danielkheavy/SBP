VERSION 5.00
Begin VB.Form CENTRADI 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centralizaciones"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   10965
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10905
      TabIndex        =   13
      Top             =   0
      Width           =   10965
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "CENTRADI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      TabIndex        =   5
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7680
      TabIndex        =   4
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox caja 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox fechai 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label internos1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label otros1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label otros 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label internos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label anulados 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label nregistro 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Procesados"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Menu dfo223 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "CENTRADI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CAJA_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command1.SetFocus

End Sub

Private Sub cmdExit_Click()
    dfo223_Click

End Sub

Private Sub Command1_Click()

    Dim found As Integer

    If Len(fechai) <> 10 Then
        fechai.SetFocus
        Exit Sub

    End If

    If Not IsDate(fechai) Then
        fechai = ""
        fechai.SetFocus
        Exit Sub

    End If

    fechai = Format(fechai, "dd/mm/yyyy")
    found = busca_caja()

    If found = 0 Then
        MsgBox "NO EXISTE NUMERO DE CAJA", 24, "AVISO"
        caja.SetFocus
        Exit Sub

    End If

    'If Len(ruta) = 0 Then
    '   MsgBox "Ruta no Definido ", 48, "Aviso"
    '   caja.SetFocus
    '   Exit Sub
    'End If
    'If Len(alocal) = 0 Then Exit Sub
    If MsgBox("DESEA PROCESAR..", 1, "AVISO") <> 1 Then Exit Sub
    Command1.Enabled = False
    found = busca_producto()

    If found = 0 Then
        Command1.Enabled = True
        Exit Sub

    End If

    fechai.Enabled = False
    Command1.Enabled = False
    caja.Enabled = False
    Command2.Enabled = False
    found = copia_cabecera()
    fechai.Enabled = True
    Command1.Enabled = True
    caja.Enabled = True
    Command2.Enabled = True
    MsgBox "Proceso Terminado ", 48, "Aviso"

End Sub

Private Sub Command2_Click()
    dfo223_Click

End Sub

Private Sub dfo223_Click()

    If Command1.Enabled = False Then Exit Sub
    CENTRADI.Hide
    Unload CENTRADI

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)

End Sub

Function busca_caja()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_caja = 1

    End If

    mytablex.Close

End Function

Function busca_producto()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select fecha from cadiario where fecha='" & Format(fechai, "YYYYMMDD") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_producto = 1

    End If

    mytablex.Close
    Exit Function

End Function

Function copia_cabecera()

    Dim found    As Integer

    Dim I        As Integer

    Dim vr       As Integer

    Dim num_rec  As Long

    Dim banula   As Integer

    Dim pasos    As Double

    Dim sw       As Integer

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim mytablem As New ADODB.Recordset

    Dim mytablev As New ADODB.Recordset

    Dim mytableT As New ADODB.Recordset

    On Error GoTo cmd15_err

    sum1 = 0
    banula = 0
    sum2 = 0
    sum3 = 0
    sum4 = 0
    sum5 = 0
    sum6 = 0
    pasos = 0
    otros = ""
    internos = ""
    otros1 = ""
    internos1 = ""
    sw = 0
    'Set mytablex = mydbxglo.OpenTable("factura")
    'mytablex.Index = "tfactura"
    'Set mytablez = mydbxglo.OpenTable("detalle")
    'Set mytablem = mydbxglo.OpenTable("fpagov")
    'Set mydby = OpenDatabase(RUTA, False, False, "foxpro 2.5;")
    'Set mytabley = mydby.OpenTable("cadiario")
    'Set mytables = mydby.OpenTable("dediario")
    'Set mytablen = mydby.OpenTable("fpdiario")
    
    'mytablez.Open "select * from detalle ", cn, adOpenStatic, adLockOptimistic
    
    mytablev.Open "select * from detalle where tipo='s' ", cn, adOpenStatic, adLockOptimistic
    mytableT.Open "select * from fpagov where tipo='s' ", cn, adOpenStatic, adLockOptimistic
   
    mytabley.Open "select * from cadiario where fecha='" & Format(fechai, "YYYYMMDD") & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        Do

            If mytabley.EOF Or mytabley.BOF Then Exit Do
            sum1 = sum1 + 1
            nregistro = Format(sum1, "00000")
            vr = DoEvents()
            Set mytablex = Nothing
            sw = 0

            If mytablex.State = 1 Then mytablex.Close
            mytablex.Open "select * from factura where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

            If mytablex.RecordCount > 0 Then

                '-------------------------------
                'GoTo pepito
                If "" & mytabley.Fields("estado") = "1" Then
                    If "" & mytabley.Fields("TIPO") = "5" Then
                        sum5 = sum5 + 1
                    Else
                        sum6 = sum6 + 1

                    End If

                    sum4 = sum4 + 1
                    banula = banula + 1
                    anulados = Format(banula, "00000")
                    otros = Format(sum6, "00000")
                    internos = Format(sum5, "00000")
                    vr = DoEvents()
                    'mytablex.Edit
                    mytablex.Fields("estado") = "1"
                    mytablex.Update
                    'ahora en detalle
                    cn.Execute ("update detalle set estado='1' where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'")
                    'ahora en fpagov
                    sw = 2
                    cn.Execute ("update fpagov set estado='1' where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'")

                End If

                'MsgBox "xxx"
pepito:
            Else
                '-------------COPIANDO CABECERA-------------------------
                
                If "" & mytabley.Fields("TIPO") = "5" Then
                    sum2 = sum2 + 1
                Else
                    sum3 = sum3 + 1

                End If

                otros1 = Format(sum3, "00000")
                internos1 = Format(sum2, "00000")
                mytablex.AddNew

                For I = 0 To mytabley.Fields.count - 1
                    mytablex.Fields(I) = mytabley.Fields(I)
                Next I

                mytablex.Update
                'found = copiar_video(mytabley)
                found = copia_detalle(mytabley, mytablev)
                found = copia_fpago(mytabley, mytableT)
                'found = copia_grafico()
                found = 1

                '--------------------------------------
            End If

            mytablex.Close
            Set mytablex = Nothing
            mytabley.MoveNext
        Loop

    End If

    '----------------------
    mytabley.Close
    Set mytabley = Nothing
    
    found = copia_ingreso(mytableT)
    mytablev.Close
    mytableT.Close
    copia_cabecera = 1
    Exit Function
cmd15_err:
    MsgBox "Aviso en copia cabecera " + Str(sw) + " " + error$, 24, "AVISO DE NO ERROR"
    Exit Function

End Function

Function copia_detalle(mytabley As ADODB.Recordset, mytablec As ADODB.Recordset)

    Dim I        As Integer

    Dim mytablea As New ADODB.Recordset

    On Error GoTo cmd16_err

    mytablea.Open "select * from dediario where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount > 0 Then
        Do

            If mytablea.EOF Or mytablea.BOF Then Exit Do
            '-------------COPIANDO detalle-------------------------
            mytablec.AddNew

            For I = 0 To mytablea.Fields.count - 1
                mytablec.Fields(I) = mytablea.Fields(I)
            Next I

            mytablec.Update
            found = 1
            mytablea.MoveNext
        Loop

    End If

    mytablea.Close
    Exit Function
cmd16_err:
    MsgBox "Aviso en copia detalle " + error$, 24, "AVISO DE NO ERROR"
    Exit Function

End Function

Function copia_fpago(mytabley As ADODB.Recordset, mytablec As ADODB.Recordset)

    Dim I        As Integer

    Dim mytablea As New ADODB.Recordset

    On Error GoTo cmd17_err

    mytablea.Open "select * from fpdiario where local='" & mytabley.Fields("local") & "' and tipo='" & mytabley.Fields("tipo") & "' and serie='" & mytabley.Fields("serie") & "' and numero='" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablea.RecordCount > 0 Then
        Do

            If mytablea.EOF Or mytablea.BOF Then Exit Do
            '-------------COPIANDO detalle-------------------------
            mytablec.AddNew

            For I = 0 To mytablea.Fields.count - 1
                mytablec.Fields(I) = mytablea.Fields(I)
            Next I

            mytablec.Update
            found = 1
            mytablea.MoveNext
        Loop

    End If

    mytablea.Close
    
    Exit Function
cmd17_err:
    MsgBox "Aviso en copia fpagov " + error$, 24, "AVISO DE NO ERROR"
    Exit Function

End Function

Function copia_ingreso(mytableT As ADODB.Recordset)

    Dim mytable5 As New ADODB.Recordset
 
    Dim found    As Integer

    Dim I        As Integer

    Dim vr       As Integer

    On Error GoTo cmd216_err

    mytable5.Open "select * from recibo where fecha='" & Format(fechai, "YYYYMMDD") & "'", cn, adOpenStatic, adLockOptimistic

    If mytable5.RecordCount > 0 Then
        Do

            If mytable5.EOF Or mytable5.BOF Then Exit Do
            found = copia_fpago(mytable5, mytableT)
            mytable5.MoveNext
        Loop

    End If

    mytable5.Close
    Exit Function
cmd216_err:
    MsgBox "Aviso en copia ingreso", 24, "AVISO DE NO ERROR"
    Exit Function

End Function

Private Sub fechai_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechai) = 0 Then
        fechai = Format(Now, "dd/mm/yyyy")

    End If

    caja.SetFocus

End Sub

Private Sub Form_Load()
    fechai = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub local1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    fechai.SetFocus

End Sub

Function copiar_video(mytablex As ADODB.Recordset)

    Dim found As Integer

    'MsgBox globaldat & "\cavideo\" & mytablex.Fields("numero")
    'found = existearchivo(globaldat & "\video\" & mytablex.Fields("numero"))
    'If found = 1 Then
    copia_video globaldat & "\cavideo\" & mytablex.Fields("serie") & "-" & mytablex.Fields("numero"), globaldat & "\video\" & mytablex.Fields("serie") & "-" & mytablex.Fields("numero")

    'End If
End Function

Sub copia_video(buf As String, buf1 As String)

    On Error GoTo cmd777111_err

    FileCopy buf, buf1
    Exit Sub
cmd777111_err:
    Exit Sub

End Sub

