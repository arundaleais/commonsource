VERSION 5.00
Begin VB.Form frmRegistry 
   Caption         =   "Registry Display"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   Icon            =   "frmRegistry.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRegistryDisplay 
      Height          =   6735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function DisplayOutput(Data As String)
    
    With txtRegistryDisplay
        If Len(.Text) + Len(Data) > 65000 Then
            .Text = Right$(.Text, Len(.Text) - Len(Data))
        End If
        .SelStart = Len(.Text)
        .SelText = Data
 '        If Len(.Text) > 4096 Then
 '           .Text = Right$(.Text, 2048)
 '       End If
    End With
End Function


Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        txtRegistryDisplay.Width = ScaleWidth - txtRegistryDisplay.Left * 2
    End If
End Sub
