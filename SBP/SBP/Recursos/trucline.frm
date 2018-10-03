VERSION 5.00
Begin VB.Form TRUCLINE 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Ruc en Linea"
   ClientHeight    =   6495
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdir 
      Height          =   405
      Left            =   2280
      MaxLength       =   150
      TabIndex        =   11
      Top             =   1800
      Width           =   9255
   End
   Begin VB.TextBox txtcon 
      Height          =   405
      Left            =   2280
      MaxLength       =   150
      TabIndex        =   10
      Top             =   1440
      Width           =   9255
   End
   Begin VB.TextBox txtest 
      Height          =   405
      Left            =   2280
      MaxLength       =   150
      TabIndex        =   9
      Top             =   1080
      Width           =   9255
   End
   Begin VB.TextBox txtrazsoc 
      Height          =   405
      Left            =   2280
      MaxLength       =   150
      TabIndex        =   8
      Top             =   720
      Width           =   9255
   End
   Begin VB.CommandButton btnCon 
      Caption         =   "Consulta En Linea"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtruc 
      Height          =   405
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label viene 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9960
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copiar"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label lbl5 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lbl4 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lbl3 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label LBL2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ruc"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu flo4443 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "TRUCLINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xDat     As String

Dim xRazSoc  As String, xEst As String, xCon As String, xdir As String

Dim xRazSocX As Long, xEstX As Long, xConX As Long, xDirX As Long

Dim xRazSocY As Long, xEstY As Long, xConY As Long, xDirY As Long

Private Sub btnCon_Click()

    If Trim(txtRuc.Text) = "" Then
        MsgBox "Ingrese número del RUC"
        txtRuc.SetFocus
        Exit Sub

    End If

    If IsNumeric(txtRuc.Text) = True Then
        If Len(txtRuc.Text) < 11 Then
            Limpiar
            MsgBox "Ingrese los 11 números del RUC"
            txtRuc.SetFocus
            Exit Sub

        End If

        If Val(Mid(Trim(txtRuc.Text), 2, 9)) = 0 Or Trim(txtRuc.Text) = "23333333333" Then
            Limpiar
            MsgBox "Verificar número del RUC"
            txtRuc.SetFocus
            Exit Sub

        End If

        If Verificar_ruc(txtRuc.Text) = False Then
            Limpiar
            MsgBox "El número del RUC no es válido"
            txtRuc.SetFocus
            Exit Sub

        End If

        'Rruc txtruc.Text
        OTRO txtRuc.Text
    Else
        Limpiar
        MsgBox "Solo se aceptan números"
        txtRuc.SetFocus

    End If

End Sub

Private Sub Rruc(ByVal xnum As String)

    On Error Resume Next

    Dim xWml As New XMLHTTP

    xWml.Open "POST", "http://www.sunat.gob.pe/w/wapS01Alias?ruc=" & xnum, False
    xWml.send

    If xWml.Status = 200 Then
        Limpiar
        xDat = xWml.responseText

        If Len(xDat) <= 635 Then
            Habilitar False
            MsgBox "El numero Ruc ingresado no existe en la Base de datos de la SUNAT"
            Set xWml = Nothing
            txtRuc.SetFocus
            Exit Sub

        End If

        Habilitar True
        xDat = Replace(xDat, "N&#xFA;mero Ruc. </b> " & xnum & " - ", "RazonSocial:")
        xDat = Replace(xDat, "Estado.</b>", "Estado:")
        xDat = Replace(xDat, "Agente Retenci&#xF3;n IGV.", "ARIGV:")
        
        xDat = Replace(xDat, "Situaci&#xF3;n.<b> ", "Situacion:")
        xDat = Replace(xDat, "Direcci&#xF3;n.</b><br/>", "Direccion:")
        xDat = Replace(xDat, "     ", " ")
        xDat = Replace(xDat, "    ", " ")
        xDat = Replace(xDat, "   ", " ")
        xDat = Replace(xDat, "  ", " ")
        xDat = Replace(xDat, "( ", "(")
        xDat = Replace(xDat, " )", ")")
       
        'MsgBox xDat
        xRazSocX = InStr(1, xDat, "RazonSocial:", vbTextCompare)
        xRazSocY = InStr(1, xDat, " <br/></small>", vbTextCompare)
        xRazSocX = xRazSocX + 12
        xRazSoc = Mid(xDat, xRazSocX, (xRazSocY - xRazSocX))
        MsgBox xRazSoc

        xEstX = InStr(1, xDat, "Estado:", vbTextCompare)
        xEstY = InStr(1, xDat, "ARIGV:", vbTextCompare)
        xEstX = xEstX + 7
        xEst = Mid(xDat, xEstX, ((xEstY - 34) - xEstX))
       
        xConX = InStr(1, xDat, "Situacion:", vbTextCompare)
        xConY = InStr(1, xDat, "</b></small><br/>", vbTextCompare)
        xDirY = xConX - 23
        xConX = xConX + 10
        xCon = Mid(xDat, xConX, (xConY - xConX))
   
        xDirX = InStr(1, xDat, "Direccion:", vbTextCompare)
        xDirX = xDirX + 10
        xdir = Mid(xDat, xDirX, (xDirY - xDirX))
       
        xRazSoc = Replace(xRazSoc, "&#209;", "Ñ")
        xRazSoc = Replace(xRazSoc, "&#xD1;", "Ñ")
        xRazSoc = Replace(xRazSoc, "&#193;", "Á")
        xRazSoc = Replace(xRazSoc, "&#201;", "É")
        xRazSoc = Replace(xRazSoc, "&#205;", "Í")
        xRazSoc = Replace(xRazSoc, "&#211;", "Ó")
        xRazSoc = Replace(xRazSoc, "&#218;", "Ú")
        xRazSoc = Replace(xRazSoc, "&#xC1;", "Á")
        xRazSoc = Replace(xRazSoc, "&#xC9;", "É")
        xRazSoc = Replace(xRazSoc, "&#xCD;", "Í")
        xRazSoc = Replace(xRazSoc, "&#xD3;", "Ó")
        xRazSoc = Replace(xRazSoc, "&#xDA;", "Ú")
       
        xdir = Replace(xdir, "&#209;", "Ñ")
        xdir = Replace(xdir, "&#xD1;", "Ñ")
        xdir = Replace(xdir, "&#193;", "Á")
        xdir = Replace(xdir, "&#201;", "É")
        xdir = Replace(xdir, "&#205;", "Í")
        xdir = Replace(xdir, "&#211;", "Ó")
        xdir = Replace(xdir, "&#218;", "Ú")
        xdir = Replace(xdir, "&#xC1;", "Á")
        xdir = Replace(xdir, "&#xC9;", "É")
        xdir = Replace(xdir, "&#xCD;", "Í")
        xdir = Replace(xdir, "&#xD3;", "Ó")
        xdir = Replace(xdir, "&#xDA;", "Ú")
       
        txtrazsoc.Text = xRazSoc
        txtest.Text = xEst
        txtcon.Text = xCon
        txtdir.Text = xdir
    Else
        Habilitar False
        Limpiar
        MsgBox "No responde el servicio de la SUNAT"

    End If

    Set xWml = Nothing

End Sub

Private Sub OTRO(ByVal xnum As String)

    On Error Resume Next

    Dim xWml As New XMLHTTP

    xWml.Open "POST", "http://www.sunat.gob.pe/w/wapS01Alias?ruc=" & xnum, False
    xWml.send

    If xWml.Status = 200 Then
        Limpiar
        xDat = xWml.responseText

        If Len(xDat) <= 635 Then
            Habilitar False
            MsgBox "El numero Ruc ingresado no existe en la Base de datos de la SUNAT"
            Set xWml = Nothing
            txtRuc.SetFocus
            Exit Sub

        End If

        Habilitar True

        Dim xTabla() As String
       
        xDat = Replace(xDat, "     ", " ")
        xDat = Replace(xDat, "    ", " ")
        xDat = Replace(xDat, "   ", " ")
        xDat = Replace(xDat, "  ", " ")
        xDat = Replace(xDat, "( ", "(")
        xDat = Replace(xDat, " )", ")")
       
        xTabla = Split(xDat, "<small>")
      
        xTabla(1) = Replace(xTabla(1), "<b>N&#xFA;mero Ruc. </b> " & xnum & " - ", "")
        xTabla(1) = Replace(xTabla(1), " <br/></small>", "")
       
        xTabla(4) = Replace(xTabla(4), "<b>Estado.</b>", "")
        xTabla(4) = Replace(xTabla(4), "</small><br/>", "")
       
        xTabla(7) = Replace(xTabla(7), "<b>Direcci&#xF3;n.</b><br/>", "")
        xTabla(7) = Replace(xTabla(7), "</small><br/>", "")
       
        xTabla(8) = Replace(xTabla(8), "Situaci&#xF3;n.<b> ", "")
        xTabla(8) = Replace(xTabla(8), "</b></small><br/>", "")
       
        xRazSoc = CStr(xTabla(1))
        xEst = CStr(xTabla(4))
        xdir = CStr(xTabla(7))
        xCon = CStr(xTabla(8))
       
        xRazSoc = Replace(xRazSoc, "&#209;", "Ñ")
        xRazSoc = Replace(xRazSoc, "&#xD1;", "Ñ")
        xRazSoc = Replace(xRazSoc, "&#193;", "Á")
        xRazSoc = Replace(xRazSoc, "&#201;", "É")
        xRazSoc = Replace(xRazSoc, "&#205;", "Í")
        xRazSoc = Replace(xRazSoc, "&#211;", "Ó")
        xRazSoc = Replace(xRazSoc, "&#218;", "Ú")
        xRazSoc = Replace(xRazSoc, "&#xC1;", "Á")
        xRazSoc = Replace(xRazSoc, "&#xC9;", "É")
        xRazSoc = Replace(xRazSoc, "&#xCD;", "Í")
        xRazSoc = Replace(xRazSoc, "&#xD3;", "Ó")
        xRazSoc = Replace(xRazSoc, "&#xDA;", "Ú")
       
        xRazSoc = Mid(xRazSoc, 1, Len(xRazSoc) - 3)
        MsgBox xRazSoc
        xdir = Replace(xdir, "&#209;", "Ñ")
        xdir = Replace(xdir, "&#xD1;", "Ñ")
        xdir = Replace(xdir, "&#193;", "Á")
        xdir = Replace(xdir, "&#201;", "É")
        xdir = Replace(xdir, "&#205;", "Í")
        xdir = Replace(xdir, "&#211;", "Ó")
        xdir = Replace(xdir, "&#218;", "Ú")
        xdir = Replace(xdir, "&#xC1;", "Á")
        xdir = Replace(xdir, "&#xC9;", "É")
        xdir = Replace(xdir, "&#xCD;", "Í")
        xdir = Replace(xdir, "&#xD3;", "Ó")
        xdir = Replace(xdir, "&#xDA;", "Ú")
       
        xEst = Mid(xEst, 1, Len(xEst) - 6)
        xCon = Mid(xCon, 1, Len(xCon) - 3)
        xdir = Mid(xdir, 1, Len(xdir) - 3)
       
        txtrazsoc.Text = xRazSoc
        txtest.Text = xEst
        txtcon.Text = xCon
        txtdir.Text = xdir
    Else
        Habilitar False
        Limpiar
        MsgBox "No responde el servicio de la SUNAT"

    End If

    Set xWml = Nothing

End Sub

Private Sub Limpiar()
    xRazSoc = ""
    xEst = ""
    xCon = ""
    xdir = ""
    txtrazsoc.Text = ""
    txtest.Text = ""
    txtcon.Text = ""
    txtdir.Text = ""

End Sub

Private Sub Habilitar(ByVal xOpc As Boolean)
    LBL2.Visible = xOpc
    lbl3.Visible = xOpc
    lbl4.Visible = xOpc
    lbl5.Visible = xOpc
    txtrazsoc.Visible = xOpc
    txtest.Visible = xOpc
    txtcon.Visible = xOpc
    txtdir.Visible = xOpc

End Sub

Private Sub flo4443_Click()
    TRUCLINE.Hide
    Unload TRUCLINE

End Sub

Private Sub Form_Load()
    Habilitar False

End Sub

Private Sub Label2_Click()

    If viene = "CODIGO" Then
        tdeliver.codigo = Trim(txtRuc)
        tdeliver.nombre = Mid$("" & txtrazsoc, 1, 60)

    End If

    If viene = "XRUC" Then
        tdeliver.xruc = Trim(txtRuc)
        tdeliver.xnombre = Mid$("" & txtrazsoc, 1, 60)
        tdeliver.xdireccion = Mid$("" & txtdir, 1, 60)

    End If

End Sub

Private Sub txtruc_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    btnCon_Click

End Sub

Function OTROPOSX(ByVal xnum As String) As String

    Dim xDat    As String

    Dim xRazSoc As String

    On Error Resume Next

    Dim xWml As New XMLHTTP

    xWml.Open "POST", "http://www.sunat.gob.pe/w/wapS01Alias?ruc=" & xnum, False
    xWml.send

    If xWml.Status = 200 Then
        Limpiar
        xDat = xWml.responseText

        If Len(xDat) <= 635 Then
            Habilitar False
            MsgBox "El numero Ruc ingresado no existe en la Base de datos de la SUNAT"
            Set xWml = Nothing
            'txtruc.SetFocus
            Exit Function

        End If

        Dim xTabla() As String
       
        xDat = Replace(xDat, "     ", " ")
        xDat = Replace(xDat, "    ", " ")
        xDat = Replace(xDat, "   ", " ")
        xDat = Replace(xDat, "  ", " ")
        xDat = Replace(xDat, "( ", "(")
        xDat = Replace(xDat, " )", ")")
       
        xTabla = Split(xDat, "<small>")
      
        xTabla(1) = Replace(xTabla(1), "<b>N&#xFA;mero Ruc. </b> " & xnum & " - ", "")
        xTabla(1) = Replace(xTabla(1), " <br/></small>", "")
       
        xTabla(4) = Replace(xTabla(4), "<b>Estado.</b>", "")
        xTabla(4) = Replace(xTabla(4), "</small><br/>", "")
       
        xTabla(7) = Replace(xTabla(7), "<b>Direcci&#xF3;n.</b><br/>", "")
        xTabla(7) = Replace(xTabla(7), "</small><br/>", "")
       
        xTabla(8) = Replace(xTabla(8), "Situaci&#xF3;n.<b> ", "")
        xTabla(8) = Replace(xTabla(8), "</b></small><br/>", "")
       
        xRazSoc = CStr(xTabla(1))
        'xEst = CStr(xTabla(4))
        'xDir = CStr(xTabla(7))
        'xCon = CStr(xTabla(8))
       
        xRazSoc = Replace(xRazSoc, "&#209;", "Ñ")
        xRazSoc = Replace(xRazSoc, "&#xD1;", "Ñ")
        xRazSoc = Replace(xRazSoc, "&#193;", "Á")
        xRazSoc = Replace(xRazSoc, "&#201;", "É")
        xRazSoc = Replace(xRazSoc, "&#205;", "Í")
        xRazSoc = Replace(xRazSoc, "&#211;", "Ó")
        xRazSoc = Replace(xRazSoc, "&#218;", "Ú")
        xRazSoc = Replace(xRazSoc, "&#xC1;", "Á")
        xRazSoc = Replace(xRazSoc, "&#xC9;", "É")
        xRazSoc = Replace(xRazSoc, "&#xCD;", "Í")
        xRazSoc = Replace(xRazSoc, "&#xD3;", "Ó")
        xRazSoc = Replace(xRazSoc, "&#xDA;", "Ú")
       
        xRazSoc = Mid(xRazSoc, 1, Len(xRazSoc) - 3)
        OTROPOSX = Trim(xRazSoc)
        
    Else
        MsgBox "No responde el servicio de la SUNAT"

    End If

    Set xWml = Nothing

End Function

