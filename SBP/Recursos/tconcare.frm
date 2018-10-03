VERSION 5.00
Begin VB.Form tconcare 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exporta Contabilidad Concar"
   ClientHeight    =   5655
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   480
      Width           =   3135
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Procesar"
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Año Proceso"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mes Proceso"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu dfoo9922 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tconcare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim vr

    Dim mytabler     As Table

    Dim mytables     As Table

    Dim mytableT     As Table

    Dim mydbr        As Database

    Dim mydbs        As Database

    Dim mydbt        As Database

    Dim j            As Double

    Dim xcorrelativo As Long

    Dim I            As Integer

    Dim found        As Integer

    'found = verifica_mes()
    If found = 0 Then Exit Sub
    '--------------- familias -------------
    xcorrelativo = 0
    j = 0
    Set mydbr = OpenDatabase("\orion.v2\001D\2002", False, False, "foxpro 2.5;")
    Set mytabler = mydbr.OpenTable("CABEZA")

    Set mydbs = OpenDatabase("\orion.v2\CONCAR", False, False, "foxpro 2.5;")
    Set mytables = mydbs.OpenTable("cd000403")

    Do

        If mytables.EOF Then Exit Do
        mytables.Delete
        mytables.MoveNext
    Loop

    Set mydbt = OpenDatabase("\orion.v2\CONCAR", False, False, "foxpro 2.5;")
    Set mytableT = mydbt.OpenTable("Cc000403")

    Do

        If mytableT.EOF Then Exit Do
        mytableT.Delete
        mytableT.MoveNext
    Loop
    Do
        vr = DoEvents()

        If mytabler.EOF Then Exit Do
        j = j + 1
        xnro = Format(j, "00000")

        If Year("" & mytabler.Fields("fecha")) = Val(Combo2) And Month("" & mytabler.Fields("fecha")) = Val(xmes) Then
            If Val("" & mytabler.Fields("tipo")) >= 1 And Val("" & mytabler.Fields("tipo")) <= 4 Then
                xcorrelativo = xcorrelativo + 1

                For I = 1 To 3

                    If I = 1 Then
                        mytableT.AddNew
                        mytableT.Fields("CSUBDIA") = Format(xmes, "00")
                        mytableT.Fields("CCOMPRO") = Format(xmes, "00") & Format(xcorrelativo, "0000")
                        mytableT.Fields("CFECCOM") = Format("" & mytabler.Fields("fecha"), "yymmdd")
                        mytableT.Fields("CCODMON") = "MN"
                        mytableT.Fields("CSITUA") = "F"
                        mytableT.Fields("CTIPCAM") = ""
                        mytableT.Fields("CGLOSA") = "" & mytabler.Fields("NOMBREB")
                        mytableT.Fields("CTOTAL") = Format(Val("" & mytabler.Fields("TOTAL")), "0.0000")
                        mytableT.Fields("CTIPO") = "V"
                        mytableT.Fields("CFLAG") = "N"
                        mytableT.Fields("CDATE") = Format(Now, "dd/mm/yyyy")
                        mytableT.Fields("CHORA") = Format(Now, "hhmmss")
                        mytableT.Fields("CFECCAM") = ""
                        mytableT.Fields("Cuser") = "" & mytabler.Fields("usuario")
                        mytableT.Fields("corig") = ""
                        mytableT.Fields("Cform") = ""
                        mytableT.Fields("ctipcom") = ""
                        mytableT.Fields("cextor") = ""

                        If Val("" & mytabler.Fields("estado")) = 1 Then
                            mytableT.Fields("ctotal") = 0

                        End If

                        mytableT.Update

                    End If

                    If I = 1 Then
                        mytables.AddNew
                        mytables.Fields("dsubdia") = xmes
                        mytables.Fields("dcompro") = xmes & Format(xcorrelativo, "0000")
                        mytables.Fields("dsecue") = Format(I, "0000")
                        mytables.Fields("dfeccom") = Format("" & mytabler.Fields("fecha"), "yymmdd")
                        mytables.Fields("dcodane") = "00009"

                        If Val("" & mytabler.Fields("tipo")) = 2 Or Val("" & mytabler.Fields("tipo")) = 4 Then
                            mytables.Fields("dcodane") = "" & mytabler.Fields("ruc")

                        End If

                        If Val("" & mytabler.Fields("estado")) = 1 Then
                            mytables.Fields("dcodane") = "00001"

                        End If

                        mytables.Fields("dcencos") = ""
                        mytables.Fields("dcodmon") = "MN"
                        mytables.Fields("DDH") = "D"
                        mytables.Fields("DIMPORT") = Val("" & mytabler.Fields("TOTAL"))

                        If Val("" & mytabler.Fields("estado")) = 1 Then
                            mytables.Fields("dimport") = 0

                        End If
            
                        If "" & mytabler.Fields("tipo") = "1" Then
                            mytables.Fields("dtipdoc") = "TK"
                            mytables.Fields("dcuenta") = "12109"

                        End If

                        If "" & mytabler.Fields("tipo") = "2" Then
                            mytables.Fields("dtipdoc") = "TK"
                            mytables.Fields("dcuenta") = "12109"

                        End If

                        If "" & mytabler.Fields("tipo") = "3" Then
                            mytables.Fields("dtipdoc") = "BV"
                            mytables.Fields("dcuenta") = "12105"

                        End If

                        If "" & mytabler.Fields("tipo") = "4" Then
                            mytables.Fields("dtipdoc") = "FT"
                            mytables.Fields("dcuenta") = "12101"

                        End If

                        mytables.Fields("dnumdoc") = "" & mytabler.Fields("numero")
                        mytables.Fields("dfecdoc") = Format("" & mytabler.Fields("fecha"), "yymmdd")
                        mytables.Fields("dfecven") = ""
                        mytables.Fields("darea") = ""
                        mytables.Fields("dflag") = "N"
                        mytables.Fields("dxglosa") = "" & mytabler.Fields("nombreb")
                        mytables.Fields("ddate") = Format(Now, "yy/mm/dd")
                        mytables.Fields("dcodane2") = ""
                        mytables.Fields("dusimpor") = ""
                        mytables.Fields("dmnimpor") = ""
                        mytables.Fields("dcodarc") = ""
                        'mytables.Fields("dnuasiento") = ""
                        'mytables.Fields("druc") = "" & mytabler.Fields("ruc")
                        'mytables.Fields("dnulibro") = ""
                        mytables.Update

                    End If

                    If I = 2 Then
                        mytables.AddNew
                        mytables.Fields("dsubdia") = xmes
                        mytables.Fields("dcompro") = xmes & Format(xcorrelativo, "0000")
                        mytables.Fields("dsecue") = Format(I, "0000")
                        mytables.Fields("dfeccom") = Format("" & mytabler.Fields("fecha"), "yymmdd")
                        mytables.Fields("dcuenta") = "40101"
                        mytables.Fields("dcodane") = "00009"

                        If Val("" & mytabler.Fields("tipo")) = 2 Or Val("" & mytabler.Fields("tipo")) = 4 Then
                            mytables.Fields("dcodane") = "" & mytabler.Fields("ruc")

                        End If

                        If Val("" & mytabler.Fields("estado")) = 1 Then
                            mytables.Fields("dcodane") = "00001"

                        End If
            
                        mytables.Fields("dcencos") = ""
                        mytables.Fields("dcodmon") = "MN"
                        mytables.Fields("DDH") = "H"
                        mytables.Fields("DIMPORT") = Val("" & mytabler.Fields("impuesto"))

                        If Val("" & mytabler.Fields("estado")) = 1 Then
                            mytables.Fields("dimport") = 0

                        End If

                        If "" & mytabler.Fields("tipo") = "1" Then
                            mytables.Fields("dtipdoc") = "TK"

                        End If

                        If "" & mytabler.Fields("tipo") = "2" Then
                            mytables.Fields("dtipdoc") = "TK"

                        End If

                        If "" & mytabler.Fields("tipo") = "3" Then
                            mytables.Fields("dtipdoc") = "BV"

                        End If

                        If "" & mytabler.Fields("tipo") = "4" Then
                            mytables.Fields("dtipdoc") = "FT"

                        End If
            
                        mytables.Fields("dnumdoc") = "" & mytabler.Fields("numero")
                        mytables.Fields("dfecdoc") = Format("" & mytabler.Fields("fecha"), "yymmdd")
                        mytables.Fields("dfecven") = ""
                        mytables.Fields("darea") = ""
                        mytables.Fields("dflag") = "N"
                        mytables.Fields("dxglosa") = "" & mytabler.Fields("nombreb")
                        mytables.Fields("ddate") = Format(Now, "yy/mm/dd")
                        mytables.Fields("dcodane2") = ""
                        mytables.Fields("dusimpor") = ""
                        mytables.Fields("dmnimpor") = ""
                        mytables.Fields("dcodarc") = ""
                        mytables.Update

                        '
                    End If

                    If I = 3 Then
                        mytables.AddNew
                        mytables.Fields("dsubdia") = xmes
                        mytables.Fields("dcompro") = xmes & Format(xcorrelativo, "0000")
                        mytables.Fields("dsecue") = Format(I, "0000")
                        mytables.Fields("dfeccom") = Format("" & mytabler.Fields("fecha"), "yymmdd")
                        mytables.Fields("dcuenta") = "70101"
                        mytables.Fields("dcodane") = "00009"

                        If Val("" & mytabler.Fields("tipo")) = 2 Or Val("" & mytabler.Fields("tipo")) = 4 Then
                            mytables.Fields("dcodane") = "" & mytabler.Fields("ruc")

                        End If

                        If Val("" & mytabler.Fields("estado")) = 1 Then
                            mytables.Fields("dcodane") = "00001"

                        End If
            
                        mytables.Fields("dcencos") = ""
                        mytables.Fields("dcodmon") = "MN"
                        mytables.Fields("DDH") = "H"
                        mytables.Fields("DIMPORT") = Val("" & mytabler.Fields("SUBTOTAL"))

                        If Val("" & mytabler.Fields("estado")) = 1 Then
                            mytables.Fields("dimport") = 0

                        End If

                        If "" & mytabler.Fields("tipo") = "1" Then
                            mytables.Fields("dtipdoc") = "TK"

                        End If

                        If "" & mytabler.Fields("tipo") = "2" Then
                            mytables.Fields("dtipdoc") = "TK"

                        End If

                        If "" & mytabler.Fields("tipo") = "3" Then
                            mytables.Fields("dtipdoc") = "BV"

                        End If

                        If "" & mytabler.Fields("tipo") = "4" Then
                            mytables.Fields("dtipdoc") = "FT"

                        End If

                        mytables.Fields("dnumdoc") = "" & mytabler.Fields("numero")
                        mytables.Fields("dfecdoc") = Format("" & mytabler.Fields("fecha"), "yymmdd")
                        mytables.Fields("dfecven") = ""
                        mytables.Fields("darea") = ""
                        mytables.Fields("dflag") = "N"
                        mytables.Fields("dxglosa") = "" & mytabler.Fields("nombreb")
                        mytables.Fields("ddate") = Format(Now, "yy/mm/dd")
                        mytables.Fields("dcodane2") = ""
                        mytables.Fields("dusimpor") = ""
                        mytables.Fields("dmnimpor") = ""
                        mytables.Fields("dcodarc") = ""
                        mytables.Update

                        If Val("" & mytabler.Fields("tax1")) > 0 Then 'inafecto
                            mytables.AddNew
                            mytables.Fields("dsubdia") = xmes
                            mytables.Fields("dcompro") = xmes & Format(xcorrelativo, "0000")
                            mytables.Fields("dsecue") = Format(I, "0000")
                            mytables.Fields("dfeccom") = Format("" & mytabler.Fields("fecha"), "yymmdd")
                            mytables.Fields("dcuenta") = "70101"
                            mytables.Fields("dcodane") = "00009"

                            If Val("" & mytabler.Fields("tipo")) = 2 Or Val("" & mytabler.Fields("tipo")) = 4 Then
                                mytables.Fields("dcodane") = "" & mytabler.Fields("ruc")

                            End If

                            If Val("" & mytabler.Fields("estado")) = 1 Then
                                mytables.Fields("dcodane") = "" & mytabler.Fields("00001")

                            End If
            
                            mytables.Fields("dcencos") = ""
                            mytables.Fields("dcodmon") = "MN"
                            mytables.Fields("DDH") = "H"
                            mytables.Fields("DIMPORT") = Val("" & mytabler.Fields("tax1"))

                            If Val("" & mytabler.Fields("estado")) = 1 Then
                                mytables.Fields("dimport") = 0

                            End If

                            If "" & mytabler.Fields("tipo") = "1" Then
                                mytables.Fields("dtipdoc") = "TK"

                            End If

                            If "" & mytabler.Fields("tipo") = "2" Then
                                mytables.Fields("dtipdoc") = "TK"

                            End If

                            If "" & mytabler.Fields("tipo") = "3" Then
                                mytables.Fields("dtipdoc") = "BV"

                            End If

                            If "" & mytabler.Fields("tipo") = "4" Then
                                mytables.Fields("dtipdoc") = "FT"

                            End If
            
                            mytables.Fields("dnumdoc") = "" & mytabler.Fields("numero")
                            mytables.Fields("dfecdoc") = Format("" & mytabler.Fields("fecha"), "yymmdd")
                            mytables.Fields("dfecven") = ""
                            mytables.Fields("darea") = ""
                            mytables.Fields("dflag") = "N"
                            mytables.Fields("dxglosa") = "" & mytabler.Fields("nombreb")
                            mytables.Fields("ddate") = Format(Now, "yy/mm/dd")
                            mytables.Fields("dcodane2") = ""
                            mytables.Fields("dusimpor") = ""
                            mytables.Fields("dmnimpor") = ""
                            mytables.Fields("dcodarc") = ""
                            mytables.Update
            
                        End If

                    End If

                Next I

            End If

        End If

        mytabler.MoveNext
    Loop
    mytableT.Close
    mydbt.Close
    mytables.Close
    mydbs.Close
    mytabler.Close
    mydbr.Close
    MsgBox "Proceso Terminado", 24, "Aviso"

End Sub

Private Sub Form_Load()

    Dim I As Integer

    Combo3.Clear       'sub serie nombre
    Combo3.AddItem "30 001 Miraflores manual"
    Combo3.AddItem "07 002 Dutty Free"
    Combo3.AddItem "05 003 Ferias"
    Combo3.AddItem "30 004 Miraflores Ticket"
    Combo3.AddItem "48 005 "
    Combo3.AddItem "05 006 Ferias"
    Combo3.AddItem "37 007 Nacional"
    Combo3.AddItem "08 009 El comercio"
    Combo3.AddItem "36 010 San Miguel"
    Combo3.AddItem "10 011 Cusco"
    Combo3.AddItem "15 012 "
    Combo3.AddItem "09 013 Dutty Free Manuales"
    Combo3.AddItem "03 014 Larco Mar"
    Combo3.AddItem "04 015 Jockey"
    Combo3.AddItem "20 016 Mezanine"
    Combo3.AddItem "07 017 Arequipa"
    Combo3.AddItem "26 018 Independencia"
    Combo3.Text = "30 001 Miraflores manual"

    Combo1.Clear
    Combo1.AddItem "ENERO"
    Combo1.AddItem "FEBRERO"
    Combo1.AddItem "MARZO"
    Combo1.AddItem "ABRIL"
    Combo1.AddItem "MAYO"
    Combo1.AddItem "JUNIO"
    Combo1.AddItem "JULIO"
    Combo1.AddItem "AGOSTO"
    Combo1.AddItem "SETIEMBRE"
    Combo1.AddItem "OCTUBRE"
    Combo1.AddItem "NOVIEMBRE"
    Combo1.AddItem "DICIEMBRE"
    Combo1.Text = "ENERO"
    'If Month(Now) = 1 Then
    Combo1.ListIndex = Month(Now) - 1
   
    'End If
    Combo2.Clear

    For I = 2000 To 2010
        Combo2.AddItem Format(I, "0000")
    Next I

    Combo2.AddItem Format(Year(Now), "0000")

End Sub
