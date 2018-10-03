VERSION 5.00
Begin VB.Form tcenremo 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "centralizaciones Remotos"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   10560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3840
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.TextBox clave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox caja 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox ruta 
      Enabled         =   0   'False
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
      Left            =   2280
      MaxLength       =   60
      TabIndex        =   4
      Text            =   "\\192.168.1.34\servidor (D)\rp_orion.v2\001d\01"
      Top             =   1320
      Width           =   8175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   3
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      TabIndex        =   2
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox FECHA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clave Responsable"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOCAL"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label tipo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RUTA REMOTO"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "REGISTROS"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label registro 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Menu flo34341 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcenremo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CAJA_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = busca_caja()
If found = 0 Then
   caja.SetFocus
   Exit Sub
End If


clave.SetFocus
End Sub

Private Sub clave_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Command1.SetFocus
End Sub

Private Sub Command1_Click()
If Len(fecha) <> 10 Then
   fecha = ""
   fecha.SetFocus
   Exit Sub
End If
If Not IsDate(fecha) Then
   fecha = ""
   fecha.SetFocus
   Exit Sub
End If

found = busca_caja()
If found = 0 Then
   caja = ""
   caja.SetFocus
   Exit Sub
End If
If clave <> "CENTRAL" Then
   clave = ""
   clave.SetFocus
   Exit Sub
End If
If Len(ruta) = 0 Then
   MsgBox "Error en Ruta ", 48, "Aviso"
   caja = ""
   caja.SetFocus
   Exit Sub
End If
If Trim(tipo) = "PRODUCTO" Then
   If MsgBox("Esta seguro de procesar", 1, "Aviso") <> 1 Then Exit Sub
   importa_productos
End If
If Trim(tipo) = "VENTAS" Then
   found = busca_producto()
   If found = 0 Then
      MsgBox "No hay Datos ", 48, "Aviso"
      Command3.Visible = False
    registro = "0"
      Exit Sub
   End If
   found = copia_cabecera()
   If found = 0 Then
      MsgBox "No se pudo centralizar,intente mas tarde", 48, "Aviso"
      Command3.Visible = False
    registro = "0"
      Exit Sub
   End If
End If
Command3.Visible = False
    registro = "0"
End Sub
Sub importa_productos()
Dim mytablex As Table
Dim mytabley As Table
Dim mydbx As Database
Dim vr
sdx = 0
Command3.Visible = True
registro = "0"
If Len(ruta) = 0 Then
   'ruta.SetFocus
   Exit Sub
   End If
Set mydbx = OpenDatabase(ruta, False, False, "foxpro 2.5;")
   Set mytabley = mydbxglo.OpenTable("producto")
   Set mytablex = mydbx.OpenTable("producto")
   mytabley.Index = "producto"
   Do
   If mytablex.EOF Then Exit Do
      '----------------------------------------
      sdx = sdx + 1
      vr = DoEvents()
      registro = "Producto " & sdx
      If Command3.Visible = False Then
         Exit Do
      End If
      mytabley.Seek "=", "" & mytablex.Fields("producto")
      If Not mytabley.NoMatch Then
         mytabley.Edit
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      If mytabley.NoMatch Then
         mytabley.AddNew
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      '----------------------------------------
      mytablex.MoveNext
   Loop
   mytablex.Close
   'familia
   sdx = 0
   Set mytabley = mydbxglo.OpenTable("familia")
   Set mytablex = mydbx.OpenTable("familia")
   mytabley.Index = "familia"
   Do
   If mytablex.EOF Then Exit Do
      '----------------------------------------
      sdx = sdx + 1
      vr = DoEvents()
      registro = "Familia " & sdx
      If Command3.Visible = False Then
         Exit Do
      End If
      mytabley.Seek "=", "" & mytablex.Fields("familia")
      If Not mytabley.NoMatch Then
         mytabley.Edit
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      If mytabley.NoMatch Then
         mytabley.AddNew
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      '----------------------------------------
      mytablex.MoveNext
   Loop
   mytablex.Close
   'subfamilia
   sdx = 0
Set mytabley = mydbxglo.OpenTable("subfamil")
   Set mytablex = mydbx.OpenTable("subfamil")
   mytabley.Index = "subfamilia"
   Do
   If mytablex.EOF Then Exit Do
      '----------------------------------------
      sdx = sdx + 1
      vr = DoEvents()
      registro = "Subfamilia " & sdx
      If Command3.Visible = False Then
         Exit Do
      End If
      mytabley.Seek "=", "" & mytablex.Fields("familia"), "" & mytablex.Fields("subfamilia")
      If Not mytabley.NoMatch Then
         mytabley.Edit
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      If mytabley.NoMatch Then
         mytabley.AddNew
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      '----------------------------------------
      mytablex.MoveNext
   Loop
   mytablex.Close
   'seccion
   sdx = 0
   Set mytabley = mydbxglo.OpenTable("seccion")
   Set mytablex = mydbx.OpenTable("seccion")
   mytabley.Index = "seccion"
   Do
   If mytablex.EOF Then Exit Do
      '----------------------------------------
      sdx = sdx + 1
      vr = DoEvents()
      registro = "Seccion " & sdx
      If Command3.Visible = False Then
         Exit Do
      End If
      mytabley.Seek "=", "" & mytablex.Fields("seccion")
      If Not mytabley.NoMatch Then
         mytabley.Edit
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      If mytabley.NoMatch Then
         mytabley.AddNew
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      '----------------------------------------
      mytablex.MoveNext
   Loop
   mytablex.Close
   'marca
   sdx = 0
   Set mytabley = mydbxglo.OpenTable("marca")
   Set mytablex = mydbx.OpenTable("marca")
   mytabley.Index = "marca"
   Do
   If mytablex.EOF Then Exit Do
      '----------------------------------------
      sdx = sdx + 1
      vr = DoEvents()
      registro = "Marca " & sdx
      If Command3.Visible = False Then
         Exit Do
      End If
      mytabley.Seek "=", "" & mytablex.Fields("marca")
      If Not mytabley.NoMatch Then
         mytabley.Edit
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      If mytabley.NoMatch Then
         mytabley.AddNew
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      '----------------------------------------
      mytablex.MoveNext
   Loop
   mytablex.Close
   'categoria
   sdx = 0
   Set mytabley = mydbxglo.OpenTable("categoria")
   Set mytablex = mydbx.OpenTable("categoria")
   mytabley.Index = "categoria"
   Do
   If mytablex.EOF Then Exit Do
      '----------------------------------------
      sdx = sdx + 1
      vr = DoEvents()
      registro = "Categoria " & sdx
      If Command3.Visible = False Then
         Exit Do
      End If
      mytabley.Seek "=", "" & mytablex.Fields("categoria")
      If Not mytabley.NoMatch Then
         mytabley.Edit
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      If mytabley.NoMatch Then
         mytabley.AddNew
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      '----------------------------------------
      mytablex.MoveNext
   Loop
   mytablex.Close
   'linea
   sdx = 0
   Set mytabley = mydbxglo.OpenTable("linea")
   Set mytablex = mydbx.OpenTable("linea")
   mytabley.Index = "linea"
   Do
   If mytablex.EOF Then Exit Do
      '----------------------------------------
      sdx = sdx + 1
      vr = DoEvents()
      registro = "Linea " & sdx
      If Command3.Visible = False Then
         Exit Do
      End If
      mytabley.Seek "=", "" & mytablex.Fields("linea")
      If Not mytabley.NoMatch Then
         mytabley.Edit
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      If mytabley.NoMatch Then
         mytabley.AddNew
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      '----------------------------------------
      mytablex.MoveNext
   Loop
   mytablex.Close
   'color
   sdx = 0
   Set mytabley = mydbxglo.OpenTable("color")
   Set mytablex = mydbx.OpenTable("color")
   mytabley.Index = "color"
   Do
   If mytablex.EOF Then Exit Do
      '----------------------------------------
      sdx = sdx + 1
      vr = DoEvents()
      registro = "Color " & sdx
      If Command3.Visible = False Then
         Exit Do
      End If
      mytabley.Seek "=", "" & mytablex.Fields("color")
      If Not mytabley.NoMatch Then
         mytabley.Edit
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      If mytabley.NoMatch Then
         mytabley.AddNew
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      '----------------------------------------
      mytablex.MoveNext
   Loop
   mytablex.Close
   'fabrica
   sdx = 0
   Set mytabley = mydbxglo.OpenTable("fabrica")
   Set mytablex = mydbx.OpenTable("fabrica")
   mytabley.Index = "codigo"
   Do
   If mytablex.EOF Then Exit Do
      '----------------------------------------
      sdx = sdx + 1
      vr = DoEvents()
      registro = "Fabrica " & sdx
      If Command3.Visible = False Then
         Exit Do
      End If
      mytabley.Seek "=", "" & mytablex.Fields("codigo")
      If Not mytabley.NoMatch Then
         mytabley.Edit
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      If mytabley.NoMatch Then
         mytabley.AddNew
         pone_registros mytablex, mytabley
         mytabley.Update
      End If
      '----------------------------------------
      mytablex.MoveNext
   Loop
   mytablex.Close
   mydbx.Close
   Command3.Visible = False
   MsgBox "Proceso Terminado", 48, "Aviso"
   Command3.Visible = False

End Sub

Function busca_caja()
Dim mytablex As Table
   ruta = ""
   Set mytablex = mydbxglo.OpenTable("tlocal")
   mytablex.Index = "codigo"
   mytablex.Seek "=", caja
   If Not mytablex.NoMatch Then
      If tipo = "PRODUCTO" Then
      busca_caja = 1
      ruta = "" & mytablex.Fields("rutaprod")
      End If
      If tipo = "VENTAS" Then
      busca_caja = 1
      ruta = "" & mytablex.Fields("rutaVTA")
      End If
   End If
   mytablex.Close
End Function

Private Sub Command3_Click()
If Command3.Visible = True Then
   Command3.Visible = False
End If
End Sub

Private Sub FECHA_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
caja.SetFocus
End Sub

Private Sub flo34341_Click()
If Command3.Visible = True Then
   Command3.Visible = False
   Exit Sub
End If

tcenremo.Hide
Unload tcenremo
End Sub

Private Sub Form_Load()
fecha = Format(Now, "dd/mm/yyyy")
End Sub
Sub pone_registros(mytablex As Table, mytabley As Table)
Dim i As Integer
For i = 0 To mytablex.Fields.Count - 1
           mytabley.Fields(i) = mytablex.Fields(i)
       Next i
End Sub

Function copia_cabecera()
Dim found  As Integer
Dim i As Integer
Dim vr As Integer
Dim num_rec As Long
Dim banula As Integer
Dim pasos As Double
Dim mytablex As Table
Dim mytabley As Table
Dim mytablez As Table
Dim mytablea As Table
Dim mytablem As Table
Dim mytablen As Table
Dim mytables As Table
Dim mydbx As Database
Dim mydby As Database
Dim mydbz As Database
Dim mydba As Database
Dim mydbm As Database
Dim mydbs As Database
Dim mydbn As Database
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
    sdx = 0
    Command3.Visible = True
    registro = "0"
    If Len(ruta) = 0 Then
       'ruta.SetFocus
       Exit Function
    End If

    
    
    Set mytablex = mydbxglo.OpenTable("factura")
    mytablex.Index = "tfactura"
    Set mytablez = mydbxglo.OpenTable("detalle")
    Set mytablem = mydbxglo.OpenTable("fpagov")
    
    Set mydby = OpenDatabase(ruta, False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("factura")
    Set mytables = mydby.OpenTable("detalle")
    Set mytablen = mydby.OpenTable("fpagov")

    mytabley.Index = "fecha"
    mytabley.Seek "=", fecha
    If mytabley.NoMatch Then
       MsgBox "No existen Datos", 24, "Aviso"
       GoTo al4
    End If
    If Not mytabley.NoMatch Then
       Do
          If mytabley.EOF Then Exit Do
             If Not IsDate("" & mytabley.Fields("fecha")) Then
                MsgBox "Fecha erronea corrija", 24, "Aviso"
             End If
             If CVDate("" & mytabley.Fields("fecha")) = CVDate(fecha) Then
                'If Len("" & mytabley.Fields("tipo")) = 0 Or Len("" & mytabley.Fields("numero")) = 0 Then
                '   GoTo seguimos1
                'End If
                sum1 = sum1 + 1
                registro = Format(sum1, "00000")
                vr = DoEvents()
                mytablex.Seek "=", "" & mytabley.Fields("tipo"), "" & mytabley.Fields("serie"), "" & mytabley.Fields("numero")   '& "A"
                If Not mytablex.NoMatch Then
                   '-------------------------------
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
                      mytablex.Edit
                      mytablex.Fields("estado") = "1"
                      mytablex.Update
                      'ahora en detalle
                      mytablez.Index = "TDETALLE"
                      mytablez.Seek "=", "" & mytabley.Fields("tipo"), "" & mytabley.Fields("serie"), "" & mytabley.Fields("numero")
                      If Not mytablez.NoMatch Then
                         Do
                         If mytablez.EOF Then Exit Do
                         If "" & mytablez.Fields("tipo") = "" & mytabley.Fields("tipo") And "" & mytablez.Fields("serie") = "" & mytabley.Fields("serie") And "" & mytablez.Fields("numero") = "" & mytabley.Fields("numero") Then
                         mytablez.Edit
                         mytablez.Fields("estado") = "1"
                         mytablez.Update
                         Else: Exit Do
                         End If
                         mytablez.MoveNext
                         Loop
                      End If
                      'ahora en fpagov
                      mytablem.Index = "FPAGOV"
                      mytablem.Seek "=", "" & mytabley.Fields("tipo"), "" & mytabley.Fields("serie"), "" & mytabley.Fields("numero")
                      If Not mytablem.NoMatch Then
                         Do
                         If mytablem.EOF Then Exit Do
                         If "" & mytablem.Fields("tipo") = "" & mytabley.Fields("tipo") And "" & mytablem.Fields("serie") = "" & mytabley.Fields("serie") And "" & mytablem.Fields("numero") = "" & mytabley.Fields("numero") Then
                            '----
                            mytablem.Edit
                            mytablem.Fields("estado") = "1"
                            mytablem.Update
                            '----
                            Else: Exit Do
                         End If
                         mytablem.MoveNext
                         Loop
                      End If
                End If
                '-------------------------------
             End If
             If mytablex.NoMatch Then
                '-------------COPIANDO CABECERA-------------------------
                   If "" & mytabley.Fields("TIPO") = "5" Then
                      sum2 = sum2 + 1
                      Else
                      sum3 = sum3 + 1
                   End If
                   otros1 = Format(sum3, "00000")
                   internos1 = Format(sum2, "00000")
                   mytablex.AddNew
                   For i = 0 To mydby("factura").Fields.Count - 1
                       mytablex.Fields(i) = mytabley.Fields(i)
                   Next i
                   mytablex.Fields("local") = local_1
                   mytablex.Update
                   found = copia_detalle(mytablez, mytabley, mydby, mytables)
                   found = copia_fpago(mytablem, mytabley, mydby, mytablen)
                   'found = copia_grafico()
                   found = 1
                   '--------------------------------------
             End If
             Else: Exit Do
          End If
seguimos1:
          mytabley.MoveNext
     Loop
 End If
    '-------borrando
al4:
    '---------------------------------
    '---------------------------------
    mytabley.Close
    mytablex.Close
    mytablez.Close
    mytables.Close
    '----------------------
    found = copia_ingreso(mytablem, mydby, mytablen)
    mytablen.Close
    mytablem.Close
    mydby.Close
    '----------------------
    copia_cabecera = 1
    '----------------------
    
    Exit Function
cmd15_err:
If Err <> 3260 And Err <> 3186 And Err <> 3187 And Err <> 3158 And Err <> 3046 And Err <> 3202 Then
   MsgBox "MENSAJE, ERROR EN GRABA CABECERA " & i & " " & error$, 24, "AVISO"
   mytabley.Close
    
   mytablex.Close
    
   mytablez.Close
   mytablem.Close
   mydby.Close
   Exit Function
End If
MsgBox mensaje_bloqueo, 24, "AVISO DE NO ERROR"
Resume

End Function
Function copia_detalle(mytablez As Table, mytabley As Table, mydba As Database, mytablea As Table)
Dim found  As Integer
Dim i As Integer
Dim vr As Integer
On Error GoTo cmd16_err
    mytablea.Index = "tDETALLE"
    mytablea.Seek "=", "" & mytabley.Fields("tipo"), "" & mytabley.Fields("serie"), "" & mytabley.Fields("numero")
    If Not mytablea.NoMatch Then
       Do
       If mytablea.EOF Then
          Exit Do
       End If
       'If Len("" & mytablea.Fields("tipo")) = 0 Or Len("" & mytablea.Fields("numero")) = 0 Then
       '  GoTo al2
       'End If
      If "" & mytablea.Fields("tipo") = "" & mytabley.Fields("tipo") And "" & mytablea.Fields("serie") = "" & mytabley.Fields("serie") And "" & mytablea.Fields("numero") = "" & mytabley.Fields("numero") Then
        '-------------COPIANDO detalle-------------------------

        mytablez.AddNew
        For i = 0 To mydba("detalle").Fields.Count - 1
          mytablez.Fields(i) = mytablea.Fields(i)
        Next i
        mytablez.Fields("local") = local_1
        mytablez.Update
        found = 1
        '--------------------------------------
          Else: GoTo al2
      End If
amk1:
      mytablea.MoveNext
       Loop
    End If
    '-------borrando
al2:
    Exit Function
cmd16_err:
If Err <> 3260 And Err <> 3186 And Err <> 3187 And Err <> 3158 And Err <> 3046 And Err <> 3202 Then
   MsgBox "MENSAJE, ERROR EN GRABA DETALLE " & error$, 24, "AVISO"
   Exit Function
End If
MsgBox mensaje_bloqueo, 24, "AVISO DE NO ERROR"
Resume

End Function
Function copia_fpago(mytablem As Table, mytabley As Table, mydbn As Database, mytablen As Table)
Dim found  As Integer
Dim i As Integer
Dim vr As Integer
On Error GoTo cmd17_err
    mytablen.Index = "fpagov"
    mytablen.Seek "=", "" & mytabley.Fields("tipo"), "" & mytabley.Fields("serie"), "" & mytabley.Fields("numero")
    If Not mytablen.NoMatch Then
       Do
       'MsgBox "Hola"
       If mytablen.EOF Then Exit Do
          'If Len("" & mytablen.Fields("tipo")) = 0 Or Len("" & mytablen.Fields("numero")) = 0 Then
          '   GoTo al3
          'End If
          If "" & mytablen.Fields("tipo") = "" & mytabley.Fields("tipo") And "" & mytablen.Fields("serie") = "" & mytabley.Fields("serie") And "" & mytablen.Fields("numero") = "" & mytabley.Fields("numero") Then
                '-------------COPIANDO detalle-------------------------
                mytablem.AddNew
                For i = 0 To mydbn("fpagov").Fields.Count - 1  '- 4
                  mytablem.Fields(i) = mytablen.Fields(i)
                Next i
                mytablem.Fields("local") = local_1
                mytablem.Update
                found = 1
                '--------------------------------------
              Else: Exit Do
          End If
          mytablen.MoveNext
       Loop
    End If
    '-------borrando
    Exit Function
cmd17_err:
If Err <> 3260 And Err <> 3186 And Err <> 3187 And Err <> 3158 And Err <> 3046 And Err <> 3202 Then
   MsgBox "1.MENSAJE, ERROR EN GRABA FPAGO " & error$, 24, "AVISO"
   Exit Function
End If
MsgBox mensaje_bloqueo, 24, "AVISO DE NO ERROR"
Resume

End Function

Function copia_ingreso(mytablem As Table, mydbn As Database, mytablen As Table)
 Dim mytable5 As Table
 
Dim found  As Integer
Dim i As Integer
Dim vr As Integer
On Error GoTo cmd216_err
    Set mytable5 = mydbxglo.OpenTable("recibo")
    mytable5.Index = "FECHA"
    mytable5.Seek "=", fechai
    If Not mytable5.NoMatch Then
       Do
       If mytable5.EOF Then Exit Do
          If CVDate("" & mytable5.Fields("fecha")) = CVDate(fechai) Then
             found = copia_fpago(mytablem, mytable5, mydbn, mytablen)
             found = 1
             Else: Exit Do
          End If
      mytable5.MoveNext
       Loop
    End If
    '-------borrando
    mytable5.Close
    
    Exit Function
cmd216_err:
If Err <> 3260 And Err <> 3186 And Err <> 3187 And Err <> 3158 And Err <> 3046 And Err <> 3202 Then
   MsgBox "MENSAJE, ERROR EN GRABA INGRESO " & error$, 24, "AVISO"
   mytable5.Close
   
   Exit Function
End If
MsgBox mensaje_bloqueo, 24, "AVISO DE NO ERROR"
Resume

End Function
Function busca_producto()
Dim mytablex As Table
Dim mydbx As Database
Dim sw As Integer
   On Error GoTo cmd34_err
   sw = 0
   Set mydbx = OpenDatabase(ruta, False, False, "foxpro 2.5;")
   Set mytablex = mydbx.OpenTable("factura")
   sw = 1
   mytablex.Index = "fecha"
   mytablex.Seek "=", fecha
   If Not mytablex.NoMatch Then
      busca_producto = 1
   End If
   mytablex.Close
   mydbx.Close
   Exit Function
cmd34_err:
MsgBox "NO SE PUEDE CONECTAR A LA CAJA", 24, "AVISO"
If sw = 0 Then Exit Function
   mytablex.Close
   mydbx.Close
Exit Function
End Function


