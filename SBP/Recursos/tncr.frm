VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tncr 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Precios"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   13710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox producto 
      Height          =   855
      Left            =   2040
      MaxLength       =   15
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin MSDBGrid.DBGrid DBGrid4 
      Height          =   6375
      Left            =   120
      OleObjectBlob   =   "tncr.frx":0000
      TabIndex        =   5
      Top             =   1560
      Width           =   12255
   End
   Begin VB.Label moneda 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   8040
      Width           =   255
   End
   Begin VB.Label local1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label descripcio 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   12255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ingrese Codigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu b1212 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tncr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type campo_precio

    unidad As String
    factor As String
    precio As String

End Type

Dim campo_precios(12) As campo_precio

Private Sub b1212_Click()
    tncr.Hide
    Unload tncr

End Sub

Function busca_equiva(buf As String) As Integer

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM productb where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1
        mytablex.Close
        Exit Function

    End If

    mytablex.Close
   
    mytablex.Open "SELECT * FROM producto where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1

    End If

    mytablex.Close

End Function

Private Sub DBGrid4_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, _
                                    StartLocation As Variant, _
                                    ByVal ReadPriorRows As Boolean)

    Dim dR            As Integer

    Dim row_num       As Integer

    Dim R             As Integer

    Dim rows_returned As Integer

    If ReadPriorRows Then
        dR = -1
    Else
        dR = 1

    End If

    If IsNull(StartLocation) Then
        If ReadPriorRows Then
            row_num = RowBuf.RowCount - 1
            'row_num = 9
        Else
            row_num = 0

        End If

    Else
        row_num = CLng(StartLocation) + dR

    End If

    rows_returned = 0

    For R = 0 To RowBuf.RowCount - 1

        If row_num < 0 Or row_num > 9 Then Exit For
        RowBuf.Value(R, 0) = campo_precios(row_num).unidad
        RowBuf.Value(R, 1) = campo_precios(row_num).factor
        RowBuf.Value(R, 2) = campo_precios(row_num).precio
        
        RowBuf.Bookmark(R) = row_num
        row_num = row_num + dR
        rows_returned = rows_returned + 1
    Next R

    RowBuf.RowCount = rows_returned

End Sub

Private Sub Form_Load()
    local1 = "01"

End Sub

Private Sub producto_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(producto) = 0 Then
        Exit Sub

    End If

    found = busca_producto("" & producto)

    If found = 0 Then
        descripcio = ""
        precio = ""
        unidad = ""
        producto.SetFocus

    End If

End Sub

Function busca_producto(buf As String)

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    I = 0

    found = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        found = busca_equiva(buf) 'busca en la table codigo barras

        If found = 0 Then
            Exit Function

        End If

        mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Function

        End If

    End If

    If "" & mytablex.Fields("estado") = "N" Then  'si no esta activo
        MsgBox "Producto no activo ", 48, "Aviso"
        mytablex.Close
        Exit Function

    End If
   
    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "SELECT * FROM precios where producto='" & buf & "' and local='" & local1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        mytabley.Close
        mytablex.Close
        Exit Function

    End If
      
    descripcio = "" & mytablex.Fields("descripcio")
    precio = "" & mytabley.Fields("pventa1")
    unidad = "" & mytabley.Fields("unidad1")
      
    busca_producto = 1
    carga_dbgrid4 "" & buf

End Function

Sub carga_dbgrid4(uproducto As String)

    Dim I        As Integer

    Dim xfoto    As String

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sw       As Integer

    Dim xbodega  As String

    Dim xsaldo   As Double

    Dim xbuf     As String

    Dim xcosto   As Double

    Dim xmargen  As Double

    Dim xcostou  As Double

    Dim xfactor  As Double

    Dim xxr      As String

    Dim xxi      As String

    Dim xpreciox As Double

    Dim dmoneda  As String

    On Error GoTo cmd89111_err

    xcostou = 0

    For I = 0 To 9
        campo_precios(I).unidad = ""
        campo_precios(I).factor = ""
        campo_precios(I).precio = ""
    Next I

    'MsgBox uproducto
    xfactor = 1
    xbodega = ""
    xsaldo = 0
    xcosto = 0
    sw = 0

    dmoneda = "S"
    xfoto = ""
    descorto = ""

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM producto where  producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xcostou = 0
        xfactor = Val("" & mytablex.Fields("factor"))
        descorto = "" & mytablex.Fields("presenta")
        dmoneda = "" & mytablex.Fields("monedav")

    End If

    mytablex.Close

    If Val(paridad) <= 0 Then
        paridad = "1"

    End If

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM precios where  producto='" & uproducto & "' and local='" & "" & local1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xcosto = 0
        xpreciox = 0

        If Val("" & mytablex.Fields("factor1")) > 0 Then
            If moneda = "S" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa1"))

                If dmoneda = "D" Then
                    xpreciox = Val("" & mytablex.Fields("pventa1")) * Val(paridad)

                End If

            End If

            If moneda = "D" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa1"))

                If dmoneda = "S" Then
                    xpreciox = Val("" & mytablex.Fields("pventa1")) / Val(paridad)

                End If

            End If

            '------------------------------------------------------------
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor1"))
            campo_precios(0).unidad = "" & mytablex.Fields("unidad1")
            campo_precios(0).factor = Val("" & mytablex.Fields("factor1"))
            campo_precios(0).precio = Format(Val("" & xpreciox), "0.00")
            
        End If

        '---------
        xcosto = 0

        If Val("" & mytablex.Fields("factor2")) > 0 Then
            If moneda = "S" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa2"))

                If dmoneda = "D" Then
                    xpreciox = Val("" & mytablex.Fields("pventa2")) * Val(paridad)

                End If

            End If

            If moneda = "D" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa2"))

                If dmoneda = "S" Then
                    xpreciox = Val("" & mytablex.Fields("pventa2")) / Val(paridad)

                End If

            End If

            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor2"))
            campo_precios(1).unidad = "" & mytablex.Fields("unidad2")
            campo_precios(1).factor = Val("" & mytablex.Fields("factor2"))
            campo_precios(1).precio = Format(Val("" & xpreciox), "0.00")
   
        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor3")) > 0 Then
            If moneda = "S" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa3"))

                If dmoneda = "D" Then
                    xpreciox = Val("" & mytablex.Fields("pventa3")) * Val(paridad)

                End If

            End If

            If moneda = "D" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa3"))

                If dmoneda = "S" Then
                    xpreciox = Val("" & mytablex.Fields("pventa3")) / Val(paridad)

                End If

            End If
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor3"))
            campo_precios(2).unidad = "" & mytablex.Fields("unidad3")
            campo_precios(2).factor = Val("" & mytablex.Fields("factor3"))
            campo_precios(2).precio = Format(Val("" & xpreciox), "0.00")
   
        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor4")) > 0 Then
            If moneda = "S" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa4"))

                If dmoneda = "D" Then
                    xpreciox = Val("" & mytablex.Fields("pventa4")) * Val(paridad)

                End If

            End If

            If moneda = "D" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa4"))

                If dmoneda = "S" Then
                    xpreciox = Val("" & mytablex.Fields("pventa4")) / Val(paridad)

                End If

            End If

            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor4"))
            campo_precios(3).unidad = "" & mytablex.Fields("unidad4")
            campo_precios(3).factor = Val("" & mytablex.Fields("factor4"))
            campo_precios(3).precio = Format(Val("" & xpreciox), "0.00")
   
        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor5")) > 0 Then
            If moneda = "S" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa5"))

                If dmoneda = "D" Then
                    xpreciox = Val("" & mytablex.Fields("pventa5")) * Val(paridad)

                End If

            End If

            If moneda = "D" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa5"))

                If dmoneda = "S" Then
                    xpreciox = Val("" & mytablex.Fields("pventa5")) / Val(paridad)

                End If

            End If

            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor5"))
            campo_precios(4).unidad = "" & mytablex.Fields("unidad5")
            campo_precios(4).factor = Val("" & mytablex.Fields("factor5"))
            campo_precios(4).precio = Format(Val("" & xpreciox), "0.00")
   
        End If

        xcosto = 0
   
        If Val("" & mytablex.Fields("factor6")) > 0 Then
   
            If moneda = "S" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa6"))

                If dmoneda = "D" Then
                    xpreciox = Val("" & mytablex.Fields("pventa6")) * Val(paridad)

                End If

            End If

            If moneda = "D" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa6"))

                If dmoneda = "S" Then
                    xpreciox = Val("" & mytablex.Fields("pventa6")) / Val(paridad)

                End If

            End If
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor6"))
            campo_precios(5).unidad = "" & mytablex.Fields("unidad6")
            campo_precios(5).factor = Val("" & mytablex.Fields("factor6"))
            campo_precios(5).precio = Format(Val("" & xpreciox), "0.00")
   
            'SOLO PARA MAXIMO SE PONE PRECIO=0
            'If caja <> "08" Then
            '   campo_precios(5).precio = 0
            'End If
        End If

        'MsgBox "xx"
        xcosto = 0

        If Val("" & mytablex.Fields("factor7")) > 0 Then
            If moneda = "S" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa7"))

                If dmoneda = "D" Then
                    xpreciox = Val("" & mytablex.Fields("pventa7")) * Val(paridad)

                End If

            End If

            If moneda = "D" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa7"))

                If dmoneda = "S" Then
                    xpreciox = Val("" & mytablex.Fields("pventa7")) / Val(paridad)

                End If

            End If
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor7"))
   
            campo_precios(6).unidad = "" & mytablex.Fields("unidad7")
            campo_precios(6).factor = Val("" & mytablex.Fields("factor7"))
            campo_precios(6).precio = Format(Val("" & xpreciox), "0.00")
   
        End If
   
        xcosto = 0

        If Val("" & mytablex.Fields("factor8")) > 0 Then
   
            If moneda = "S" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa8"))

                If dmoneda = "D" Then
                    xpreciox = Val("" & mytablex.Fields("pventa8")) * Val(paridad)

                End If

            End If

            If moneda = "D" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa8"))

                If dmoneda = "S" Then
                    xpreciox = Val("" & mytablex.Fields("pventa8")) / Val(paridad)

                End If

            End If

            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor8"))
   
            campo_precios(7).unidad = "" & mytablex.Fields("unidad8")
            campo_precios(7).factor = Val("" & mytablex.Fields("factor8"))
            campo_precios(7).precio = Format(Val("" & xpreciox), "0.00")
   
        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor9")) > 0 Then
            If moneda = "S" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa9"))

                If dmoneda = "D" Then
                    xpreciox = Val("" & mytablex.Fields("pventa9")) * Val(paridad)

                End If

            End If

            If moneda = "D" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa9"))

                If dmoneda = "S" Then
                    xpreciox = Val("" & mytablex.Fields("pventa9")) / Val(paridad)

                End If

            End If

            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor9"))
            campo_precios(8).unidad = "" & mytablex.Fields("unidad9")
            campo_precios(8).factor = Val("" & mytablex.Fields("factor9"))
            campo_precios(8).precio = Format(Val("" & xpreciox), "0.00")
   
        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor10")) > 0 Then
            If moneda = "S" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa10"))

                If dmoneda = "D" Then
                    xpreciox = Val("" & mytablex.Fields("pventa10")) * Val(paridad)

                End If

            End If

            If moneda = "D" Then 'si es soles
                xpreciox = Val("" & mytablex.Fields("pventa10"))

                If dmoneda = "S" Then
                    xpreciox = Val("" & mytablex.Fields("pventa10")) / Val(paridad)

                End If

            End If
   
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor10"))
            campo_precios(9).unidad = "" & mytablex.Fields("unidad10")
            campo_precios(9).factor = Val("" & mytablex.Fields("factor10"))
            campo_precios(9).precio = Format(Val("" & xpreciox), "0.00")
   
        End If

        'MsgBox "xx"
   
        'margenes
        sw = 1

    End If

    'MsgBox ""
    'mytablex.Close
    'mytablez.Close
    DBGrid4.refresh
    '----ahora deb cargar tambien la foto del producto...
    Exit Sub
cmd89111_err:
    MsgBox "Error en carga dbgrid4 " + error$, 48, "Aviso"
    Exit Sub

End Sub

