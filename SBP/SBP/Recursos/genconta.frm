VERSION 5.00
Begin VB.Form genconta 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interfase Sistema Comercial"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   3855
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
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   3
      Top             =   600
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
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu juerl12 
      Caption         =   "&Ejecuta"
   End
   Begin VB.Menu ldso232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "genconta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
fechai = "01/" & (Format(Month(Now), "00")) & "/" & (Format(Year(Now), "0000"))
fechaf = "30/" & (Format(Month(Now), "00")) & "/" & (Format(Year(Now), "0000"))
End Sub
Sub gencontable()

End Sub

Function pasa_las_compras()
Dim buf As String
Dim xorigen As String
Dim xvoucher As String
Dim xmespro As String
Dim xbuf As String

Dim cod_asien As Double
Dim mytabled As New ADODB.Recordset
Dim mytablec As New ADODB.Recordset
Dim rbusca As New ADODB.Recordset

Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim mytablez As New ADODB.Recordset
Dim vr
On Error GoTo cmd98991221_err
'if Len(fechai) <> 10 Then Exit Function
'If Len(fechaf) <> 10 Then Exit Function
'If Not IsDate(fechai) Then Exit Function
'If Not IsDate(fechaf) Then Exit Function
'----verificamos si existen parametris
   mytablec.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
   If mytablec.RecordCount = 0 Then
      mytablec.Close
      Exit Function
   End If
   mytablez.Open "Select * from enlace_contabled where enlace='COMPRAS' and len(cuenta)>0 ", cn, adOpenStatic, adLockOptimistic
     If mytablez.RecordCount = 0 Then
        mytablez.Close
        Exit Function
   End If
'---borrando el voucher---------------------------------------
'buf = "delete  from asientos where "
'buf = buf & " fecha_asi>='" & Format(fechai, "YYYYMMDD") & "'"
'buf = buf & " and  fecha_asi<='" & Format(fechaf, "YYYYMMDD") & "'"
'buf = buf & " and orion='S'"
'MsgBox buf
'cn.Execute (buf)
'-------------------------------------------------------------
buf = "select * from factura where "
buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
buf = buf & " and  fecha<='" & Format(fechaf, "YYYYMMDD") & "'"
buf = buf & " and (acu='J' or acu='K' or  acu='L' OR acu='J') "
buf = buf & "order by fecha,tipo,serie,str(numero)"
mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Function
End If
   
mytabley.Open "Select * from asientos ", cn, adOpenStatic, adLockOptimistic
sdx = 0
Do
If mytablex.EOF Then Exit Do
     Command1.Caption = "Procesando...." & sdx
       vr = DoEvents()
     'NUMERACION DE LOS ASIENTOS
     sdx = Val("" & mytablec.Fields("asientos")) + 1
amiga:
   If rbusca.State = 1 Then
      rbusca.Close
      Set rbusca = Nothing
   End If
   rbusca.Open "select cod_asien from asientos where cod_asien=" & sdx & "", cn, adOpenStatic, adLockOptimistic
   If rbusca.RecordCount > 0 Then
      sdx = sdx + 1
      GoTo amiga
   End If
   rbusca.Close
   cod_asien = "" & sdx
   mytablec.Fields("asientos") = "" & cod_asien
   mytablec.Update
     '--------------------------------
     mytablez.MoveFirst
     Do
     If mytablez.EOF Then Exit Do
        mytabley.AddNew
        mytabley.Fields("cod_asien") = cod_asien
        mytabley.Fields("fecha_asi") = Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy")
        mytabley.Fields("nro_seq") = ""
        mytabley.Fields("cuenta") = "" & mytablez.Fields("cuenta")
        '--------cuenta
        If mytabled.State = 1 Then
           mytabled.Close
           Set mytabled = Nothing
        End If
        mytabled.Open "select * from cuentas where codcta='" & Trim("" & mytablez.Fields("cuenta")) & "'", cn, adOpenStatic, adLockOptimistic
        If mytabled.RecordCount > 0 Then
           mytabley.Fields("descripcio") = "" & mytabled.Fields("descripcio")
        End If
        mytabled.Close
        '----------------
        If mytabled.State = 1 Then
           mytabled.Close
           Set mytabled = Nothing
        End If
        mytabled.Open "select * from tipo where tipo='" & Trim("" & mytablex.Fields("tipo")) & "'", cn, adOpenStatic, adLockOptimistic
        If mytabled.RecordCount > 0 Then
           mytabley.Fields("tipo") = "" & mytabled.Fields("sunat")
        End If
        mytabled.Close

        '----------------
        If "" & mytablez.Fields("debito") = "S" Then
        mytabley.Fields("tipo_cta") = "D"
        End If
        If "" & mytablez.Fields("credito") = "S" Then
        mytabley.Fields("tipo_cta") = "H"
        End If
        If Trim("" & mytablez.Fields("tipo")) = "IGV" Then
        mytabley.Fields("CANTIDAD") = Val(Format(Val("" & mytablex.Fields("IMPUESTO")), "0.00"))
        End If
        If Trim("" & mytablez.Fields("tipo")) = "SUBTOTAL" Then
        mytabley.Fields("CANTIDAD") = Val(Format(Val("" & mytablex.Fields("subtotal")), "0.00"))
        End If
        If Trim("" & mytablez.Fields("tipo")) = "TOTAL" Then
        mytabley.Fields("CANTIDAD") = Val(Format(Val("" & mytablex.Fields("total")), "0.00"))
        End If
        mytabley.Fields("cod_libro") = 0
        mytabley.Fields("motivo") = "Centralizacion" '& mytablex.Fields("observa")
        mytabley.Fields("referencia") = "Centralizacion"
        mytabley.Fields("comproba") = "" & mytablex.Fields("serie") & "-" & mytablex.Fields("numero")
        mytabley.Fields("fuente") = "" & mytablez.Fields("fuente")
        mytabley.Fields("nro_ruc") = "" & mytablex.Fields("codigo")
        mytabley.Fields("vrbase") = 0
        mytabley.Fields("cod_cdec") = 0
        'mytabley.Fields("descripcio") = ""
        mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
        mytabley.Fields("orion") = "S"
        mytabley.Update
    mytablez.MoveNext
    Loop
    '--------------------------------
mytablex.MoveNext
Loop
mytablex.Close
mytablez.Close
MsgBox "Proceso Terminado " & sdx, 48, "Aviso"
Exit Function
cmd98991221_err:
MsgBox "Error en Pasa las Compras " & error$, 48, "Aviso"
Exit Function

End Function

Private Sub juerl12_Click()
Dim found As Integer
Command1.Visible = True
found = pasa_las_ventas()
'found = pasa_las_compras()
Command1.Visible = False
End Sub

Private Sub ldso232_Click()
If Command1.Visible = True Then
   Command1.Visible = False
   Exit Sub
End If
genconta.Hide
Unload genconta
End Sub
Function verifica_asiento_destino()
verifica_asiento_destino = 1
End Function
Function pasa_las_ventas()
Dim buf As String
Dim xorigen As String
Dim xvoucher As String
Dim xmespro As String
Dim xbuf As String

Dim pago_contado As Double
Dim pago_credito As Double
Dim pago_letra As Double

Dim cod_asien As Double
Dim mytablexx As New ADODB.Recordset
Dim mytabled As New ADODB.Recordset
Dim mytablec As New ADODB.Recordset
Dim rbusca As New ADODB.Recordset

Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim mytablez As New ADODB.Recordset
Dim vr
On Error GoTo cmd8991221_err
If Len(fechai) <> 10 Then Exit Function
If Len(fechaf) <> 10 Then Exit Function
If Not IsDate(fechai) Then Exit Function
If Not IsDate(fechaf) Then Exit Function
'----verificamos si existen parametris
   mytablec.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
   If mytablec.RecordCount = 0 Then
      mytablec.Close
      Exit Function
   End If
   mytablez.Open "Select * from enlace_contabled where enlace='FACTURA' and len(cuenta)>0 ", cn, adOpenStatic, adLockOptimistic
     If mytablez.RecordCount = 0 Then
        mytablez.Close
        Exit Function
   End If
'---borrando el voucher---------------------------------------
buf = "delete  from asientos where "
buf = buf & " fecha_asi>='" & Format(fechai, "YYYYMMDD") & "'"
buf = buf & " and  fecha_asi<='" & Format(fechaf, "YYYYMMDD") & "'"
buf = buf & " and orion='S'"
'MsgBox buf
cn.Execute (buf)
'-------------------------------------------------------------
buf = "select * from factura where "
buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
buf = buf & " and  fecha<='" & Format(fechaf, "YYYYMMDD") & "'"
buf = buf & " and (acu='A' or acu='B' or  acu='C' OR acu='D') "
buf = buf & "order by fecha,tipo,serie,str(numero)"
mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Function
End If
   
mytabley.Open "Select * from asientos ", cn, adOpenStatic, adLockOptimistic
sdx = 0
Do
If mytablex.EOF Then Exit Do
     Command1.Caption = "Procesando...." & sdx
       vr = DoEvents()
     'NUMERACION DE LOS ASIENTOS
     sdx = Val("" & mytablec.Fields("asientos")) + 1
amiga:
   If rbusca.State = 1 Then
      rbusca.Close
      Set rbusca = Nothing
   End If
   rbusca.Open "select cod_asien from asientos where cod_asien=" & sdx & "", cn, adOpenStatic, adLockOptimistic
   If rbusca.RecordCount > 0 Then
      sdx = sdx + 1
      GoTo amiga
   End If
   rbusca.Close
   cod_asien = "" & sdx
   mytablec.Fields("asientos") = "" & cod_asien
   mytablec.Update
     '--------------------------------
     'venta neta   70
     'contado      10
     'plazos       121
     'letras       122
     'dscto        71
     'ingreso financiero   77
     'igv          4011
     'isc          4012
     
     
     buf = "select * from fpagov where local='" & "" & mytablex.Fields("local") & "'"
     buf = buf & " and tipo='" & "" & mytablex.Fields("tipo") & "'"
     buf = buf & " and serie='" & "" & mytablex.Fields("serie") & "'"
     buf = buf & " and numero='" & "" & mytablex.Fields("numero") & "'"

mytablexx.Open buf, cn, adOpenStatic, adLockOptimistic
pago_contado = 0
pago_credito = 0
pago_letra = 0
If mytablexx.RecordCount > 0 Then
   Do
   If mytablexx.EOF Then Exit Do
   If "" & mytablexx.Fields("acufp") = "A" Then  'SOLES
      pago_contado = pago_contado + Val("" & mytablexx.Fields("recibe"))
   End If
   If "" & mytablexx.Fields("acufp") = "B" Then  'DOLARES
      pago_contado = pago_contado + Val("" & mytablexx.Fields("recibe"))
   End If
   If "" & mytablexx.Fields("acufp") = "C" Then  'credito
      pago_credito = pago_credito + Val("" & mytablexx.Fields("recibe"))
   End If
   If "" & mytablexx.Fields("acufp") = "D" Then  'TARJETA CREDITO
      pago_contado = pago_contado + Val("" & mytablexx.Fields("recibe"))
   End If
   If "" & mytablexx.Fields("acufp") = "E" Then  'EUROS
      pago_contado = pago_contado + Val("" & mytablexx.Fields("recibe"))
   End If
   If "" & mytablexx.Fields("acufp") = "F" Then  'TARJETA DEBITO
      pago_contado = pago_contado + Val("" & mytablexx.Fields("recibe"))
   End If
   If "" & mytablexx.Fields("acufp") = "G" Then  'letra
      pago_letra = pago_letra + Val("" & mytablexx.Fields("recibe"))
   End If
   mytablexx.MoveNext
   Loop
End If
mytablexx.Close

    
     
     
     
     mytablez.MoveFirst
     Do
     If mytablez.EOF Then Exit Do
        mytabley.AddNew
        mytabley.Fields("cod_asien") = cod_asien
        mytabley.Fields("fecha_asi") = Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy")
        mytabley.Fields("nro_seq") = ""
        mytabley.Fields("cuenta") = "" & mytablez.Fields("cuenta")
        '--------cuenta
        If mytabled.State = 1 Then
           mytabled.Close
           Set mytabled = Nothing
        End If
        mytabled.Open "select * from cuentas where codcta='" & Trim("" & mytablez.Fields("cuenta")) & "'", cn, adOpenStatic, adLockOptimistic
        If mytabled.RecordCount > 0 Then
           mytabley.Fields("descripcio") = "" & mytabled.Fields("descripcio")
        End If
        mytabled.Close
        '----------------
        If mytabled.State = 1 Then
           mytabled.Close
           Set mytabled = Nothing
        End If
        mytabled.Open "select * from tipo where tipo='" & Trim("" & mytablex.Fields("tipo")) & "'", cn, adOpenStatic, adLockOptimistic
        If mytabled.RecordCount > 0 Then
           mytabley.Fields("tipo") = "" & mytabled.Fields("sunat")
        End If
        mytabled.Close

        '----------------
        If "" & mytablez.Fields("debito") = "S" Then
        mytabley.Fields("tipo_cta") = "D"
        End If
        If "" & mytablez.Fields("credito") = "S" Then
        mytabley.Fields("tipo_cta") = "H"
        End If
        If Trim("" & mytablez.Fields("tipo")) = "IGV" Then
        mytabley.Fields("CANTIDAD") = Val(Format(Val("" & mytablex.Fields("IMPUESTO")), "0.00"))
        End If
        If Trim("" & mytablez.Fields("tipo")) = "SUBTOTAL" Then
        mytabley.Fields("CANTIDAD") = Val(Format(Val("" & mytablex.Fields("subtotal")), "0.00"))
        End If
        If Trim("" & mytablez.Fields("tipo")) = "TOTAL" Then
        mytabley.Fields("CANTIDAD") = Val(Format(Val("" & mytablex.Fields("total")), "0.00"))
        End If
        If Trim("" & mytablez.Fields("tipo")) = "PAGOEFECTIVO" Then
        mytabley.Fields("CANTIDAD") = pago_contado
        End If
        If Trim("" & mytablez.Fields("tipo")) = "PAGOCREDITO" Then
        mytabley.Fields("CANTIDAD") = pago_credito
        End If
        If Trim("" & mytablez.Fields("tipo")) = "PAGOLETRAS" Then
        mytabley.Fields("CANTIDAD") = pago_letras
        End If
        If Trim("" & mytablez.Fields("tipo")) = "DESCUENTO" Then
        mytabley.Fields("CANTIDAD") = Val(Format(Val("" & mytablex.Fields("descuento")), "0.00"))
        End If
        If Trim("" & mytablez.Fields("tipo")) = "ISC" Then
        mytabley.Fields("CANTIDAD") = Val(Format(Val("" & mytablex.Fields("ISC")), "0.00"))
        End If
        mytabley.Fields("cod_libro") = 0
        mytabley.Fields("motivo") = "Centralizacion" '& mytablex.Fields("observa")
        mytabley.Fields("referencia") = "Centralizacion"
        mytabley.Fields("comproba") = "" & mytablex.Fields("serie") & "-" & mytablex.Fields("numero")
        mytabley.Fields("fuente") = "" & mytablez.Fields("fuente")
        mytabley.Fields("nro_ruc") = "" & mytablex.Fields("codigo")
        mytabley.Fields("vrbase") = 0
        mytabley.Fields("cod_cdec") = 0
        'mytabley.Fields("descripcio") = ""
        mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
        mytabley.Fields("orion") = "S"
        mytabley.Update
    mytablez.MoveNext
    Loop
    '--------------------------------
mytablex.MoveNext
Loop
mytablex.Close
mytablez.Close
MsgBox "Proceso Terminado " & sdx, 48, "Aviso"
Exit Function
cmd8991221_err:
MsgBox "Error en Pasa las Ventas " & error$, 48, "Aviso"
Exit Function
End Function
Function pasa_las_ventasfpago()
Dim buf As String
Dim xorigen As String
Dim xvoucher As String
Dim xmespro As String
Dim xbuf As String

Dim cod_asien As Double
Dim mytabled As New ADODB.Recordset
Dim mytablec As New ADODB.Recordset
Dim rbusca As New ADODB.Recordset

Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim mytablez As New ADODB.Recordset
Dim vr
On Error GoTo cmd18991221_err
'If Len(fechai) <> 10 Then Exit Function
'If Len(fechaf) <> 10 Then Exit Function
'If Not IsDate(fechai) Then Exit Function
'If Not IsDate(fechaf) Then Exit Function
'----verificamos si existen parametris
   mytablec.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
   If mytablec.RecordCount = 0 Then
      mytablec.Close
      Exit Function
   End If
   mytablez.Open "Select * from enlace_contabled where enlace='FPAGO' and len(cuenta)>0 ", cn, adOpenStatic, adLockOptimistic
     If mytablez.RecordCount = 0 Then
        mytablez.Close
        Exit Function
   End If
'---borrando el voucher---------------------------------------
'buf = "delete  from asientos where "
'buf = buf & " fecha_asi>='" & Format(fechai, "YYYYMMDD") & "'"
'buf = buf & " and  fecha_asi<='" & Format(fechaf, "YYYYMMDD") & "'"
'buf = buf & " and orion='S'"
'MsgBox buf
'cn.Execute (buf)
'-------------------------------------------------------------
buf = "select * from fpagov where "
buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
buf = buf & " and  fecha<='" & Format(fechaf, "YYYYMMDD") & "'"
buf = buf & " and (acu='A' or acu='B' or  acu='C' OR acu='D') "
buf = buf & "order by fecha,tipo,serie,str(numero)"
mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Function
End If
   
mytabley.Open "Select * from asientos ", cn, adOpenStatic, adLockOptimistic
sdx = 0
Do
If mytablex.EOF Then Exit Do
     Command1.Caption = "Procesando...." & sdx
       vr = DoEvents()
     'NUMERACION DE LOS ASIENTOS
     sdx = Val("" & mytablec.Fields("asientos")) + 1
amiga:
   If rbusca.State = 1 Then
      rbusca.Close
      Set rbusca = Nothing
   End If
   rbusca.Open "select cod_asien from asientos where cod_asien=" & sdx & "", cn, adOpenStatic, adLockOptimistic
   If rbusca.RecordCount > 0 Then
      sdx = sdx + 1
      GoTo amiga
   End If
   rbusca.Close
   cod_asien = "" & sdx
   mytablec.Fields("asientos") = "" & cod_asien
   mytablec.Update
     '--------------------------------
     mytablez.MoveFirst
     Do
     If mytablez.EOF Then Exit Do
        mytabley.AddNew
        mytabley.Fields("cod_asien") = cod_asien
        mytabley.Fields("fecha_asi") = Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy")
        mytabley.Fields("nro_seq") = ""
        mytabley.Fields("cuenta") = "" & mytablez.Fields("cuenta")
        '--------cuenta
        If mytabled.State = 1 Then
           mytabled.Close
           Set mytabled = Nothing
        End If
        mytabled.Open "select * from cuentas where codcta='" & Trim("" & mytablez.Fields("cuenta")) & "'", cn, adOpenStatic, adLockOptimistic
        If mytabled.RecordCount > 0 Then
           mytabley.Fields("descripcio") = "" & mytabled.Fields("descripcio")
        End If
        mytabled.Close
        '----------------
        If mytabled.State = 1 Then
           mytabled.Close
           Set mytabled = Nothing
        End If
        mytabled.Open "select * from tipo where tipo='" & Trim("" & mytablex.Fields("tipo")) & "'", cn, adOpenStatic, adLockOptimistic
        If mytabled.RecordCount > 0 Then
           mytabley.Fields("tipo") = "" & mytabled.Fields("sunat")
        End If
        mytabled.Close

        '----------------
        If "" & mytablez.Fields("debito") = "S" Then
        mytabley.Fields("tipo_cta") = "D"
        End If
        If "" & mytablez.Fields("credito") = "S" Then
        mytabley.Fields("tipo_cta") = "H"
        End If
        'If Trim("" & mytablez.Fields("tipo")) = "IGV" Then
        'mytabley.Fields("CANTIDAD") = Val(Format(Val("" & mytablex.Fields("IMPUESTO")), "0.00"))
        'End If
        'If Trim("" & mytablez.Fields("tipo")) = "SUBTOTAL" Then
        'mytabley.Fields("CANTIDAD") = Val(Format(Val("" & mytablex.Fields("subtotal")), "0.00"))
        'End If
        'If Trim("" & mytablez.Fields("tipo")) = "TOTAL" Then
        mytabley.Fields("CANTIDAD") = Val(Format(Val("" & mytablex.Fields("recibe")), "0.00"))
        'End If
        mytabley.Fields("cod_libro") = 0
        mytabley.Fields("motivo") = "Centralizacion" '& mytablex.Fields("observa")
        mytabley.Fields("referencia") = "Centralizacion"
        mytabley.Fields("comproba") = "" & mytablex.Fields("serie") & "-" & mytablex.Fields("numero")
        mytabley.Fields("fuente") = "" & mytablez.Fields("fuente")
        mytabley.Fields("nro_ruc") = "" & mytablex.Fields("codigo")
        mytabley.Fields("vrbase") = 0
        mytabley.Fields("cod_cdec") = 0
        'mytabley.Fields("descripcio") = ""
        mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
        mytabley.Fields("orion") = "S"
        mytabley.Update
    mytablez.MoveNext
    Loop
    '--------------------------------
mytablex.MoveNext
Loop
mytablex.Close
mytablez.Close
'MsgBox "Proceso Terminado " & sdx, 48, "Aviso"
Exit Function
cmd18991221_err:
MsgBox "Error en Pasa las Ventas " & error$, 48, "Aviso"
Exit Function
End Function


Function busca_asiento(xorigen As String, xmes As String) As String
Dim mytablex As Table
Dim sdx As Double
Set mytablex = mydbzglo.OpenTable("origen")
mytablex.Index = "origen"
mytablex.Seek "=", xorigen
If mytablex.NoMatch Then
   busca_asiento = "-1"
End If
If Not mytablex.NoMatch Then
   sdx = 0
   Select Case Mid$(xmes, 1, 2)
          Case "01"
               sdx = Val("" & mytablex.Fields("enero")) + 1
          Case "02"
          sdx = Val("" & mytablex.Fields("febrero")) + 1
          Case "03"
          sdx = Val("" & mytablex.Fields("marzo")) + 1
          Case "04"
          sdx = Val("" & mytablex.Fields("abril")) + 1
          Case "05"
          sdx = Val("" & mytablex.Fields("mayo")) + 1
          Case "06"
          sdx = Val("" & mytablex.Fields("junio")) + 1
          Case "07"
          sdx = Val("" & mytablex.Fields("julio")) + 1
          Case "08"
          sdx = Val("" & mytablex.Fields("agosto")) + 1
          Case "09"
          sdx = Val("" & mytablex.Fields("setiembre")) + 1
          Case "10"
          sdx = Val("" & mytablex.Fields("octubre")) + 1
          Case "11"
          sdx = Val("" & mytablex.Fields("noviembre")) + 1
          Case "12"
          sdx = Val("" & mytablex.Fields("diciembre")) + 1
   End Select
   busca_asiento = Format(sdx, "0")
End If
mytablex.Close
End Function
Function graba_origen(xorigen As String, xmes As String, xnum As String)
Dim mytablex As Table
Dim sdx As Double
Set mytablex = mydbzglo.OpenTable("origen")
mytablex.Index = "origen"
mytablex.Seek "=", xorigen
If Not mytablex.NoMatch Then
   mytablex.Edit
   Select Case Mid$(xmes, 1, 2)
          Case "01"
               mytablex.Fields("enero") = xnum
          Case "02"
               mytablex.Fields("febrero") = xnum
          Case "03"
               mytablex.Fields("marzo") = xnum
          Case "04"
               mytablex.Fields("abril") = xnum
          Case "05"
               mytablex.Fields("mayo") = xnum
          Case "06"
               mytablex.Fields("junio") = xnum
          Case "07"
               mytablex.Fields("julio") = xnum
          Case "08"
               mytablex.Fields("agosto") = xnum
          Case "09"
               mytablex.Fields("setiembre") = xnum
          Case "10"
               mytablex.Fields("octubre") = xnum
          Case "11"
          mytablex.Fields("noviembre") = xnum
          Case "12"
          mytablex.Fields("diciembre") = xnum
   End Select
   mytablex.Update
   graba_origen = 1
End If
mytablex.Close
End Function
Function busca_cuenta(buf As String) As String
Dim mytablex As Table
Set mytablex = mydbzglo.OpenTable("mdh_plan")
mytablex.Index = "mdh_plan"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_cuenta = "" & mytablex.Fields("nombre")
End If
mytablex.Close
End Function
Function un_documento(mytabley As ADODB.Recordset)
End Function


