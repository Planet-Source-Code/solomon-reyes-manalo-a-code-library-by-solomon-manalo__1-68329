VERSION 5.00
Begin VB.Form frmDocument 
   Caption         =   "Document"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4080
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4080
   Begin VB.Frame fraBorder 
      Height          =   3045
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   3975
      Begin CodeLibrary.HTMLControl txtdoc 
         Height          =   2745
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   4842
      End
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pathf As String

Private Sub Form_Load()
mdiMain.Timer1.Enabled = True
txtdoc.bChange = False
  
End Sub

Private Sub Form_Resize()
On Error Resume Next
With Me
  If .Height <= 4155 Or .Width <= 5835 Then
        .Height = 4155
        .Width = 5835
  End If
     '.fraBorder.Left = 45
     '.fraBorder.Top = 510
     .fraBorder.Width = .Width - 200
     .fraBorder.Height = .Height - 500
     
     With .fraBorder
          Me.txtdoc.Left = .Left + 50
          Me.txtdoc.Width = .Width - 180
          Me.txtdoc.Height = .Height - 320
     End With
End With
End Sub

Private Sub mctoolDoc_GotFocus()
mdiMain.Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
txtdoc.bChange = False
End Sub

Private Sub txtDoc_GotFocus()
mdiMain.Timer1.Enabled = True
End Sub
