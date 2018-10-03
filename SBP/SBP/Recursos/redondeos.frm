VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim sdx As Double
sdx = Val(Text1.Text) * 0.18
MsgBox (redondear1("" & sdx))
End Sub

Function redondear1(panumero As String) As String
Dim parteentera
Dim partedecimal As String
Dim sdx
Dim num
panumero = Trim(panumero)
If Len(Trim("" & panumero)) = 0 Then
redondear1 = "0"
Exit Function
End If
panumero = Format(Val(panumero), "0.000")
If Val("" & panumero) <= 0 Then
redondear1 = panumero
Exit Function
End If
    parteentera = Int("" & panumero)
    If Not (Len(panumero) - Len(parteentera)) = 0 Then
        partedecimal = Right(panumero, Len(panumero) - Len(parteentera) - 1)
    Else
        partedecimal = "00"
    End If
    'MsgBox partedecimal
    If Val(partedecimal) >= 996 Then
       sdx = Int(panumero) + 1
       panumero = "" & sdx
       redondear1 = panumero
       Exit Function
    End If
   If Mid(partedecimal, 3, 1) > "5" Then
            partedecimal = Left(partedecimal, 2)
            num = Val(partedecimal)
            num = num + 1
            partedecimal = "" & num
      Else
            partedecimal = Left(partedecimal, 2)
      
    End If
    panumero = parteentera & "." & partedecimal
    redondear1 = panumero
    'MsgBox panumero
End Function

