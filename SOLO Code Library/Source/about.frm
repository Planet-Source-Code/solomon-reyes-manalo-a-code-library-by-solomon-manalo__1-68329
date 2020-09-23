VERSION 5.00
Begin VB.Form about 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About SOLO Code Library"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9525
   ClipControls    =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   300
      ScaleHeight     =   1395
      ScaleWidth      =   2205
      TabIndex        =   1
      Top             =   1500
      Width           =   2205
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visual Basic 6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   930
         Width           =   1350
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Language"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   690
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Solomon R. Manalo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   330
         Width           =   1860
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   90
         Width           =   750
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   """BETTER THAN ANY CODE LIBRARY"""
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1290
      Left            =   510
      TabIndex        =   6
      Top             =   3030
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   1455
      Left            =   270
      Top             =   1470
      Width           =   2265
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"about.frx":038A
      ForeColor       =   &H000000FF&
      Height          =   885
      Left            =   150
      TabIndex        =   0
      Top             =   5700
      Width           =   8580
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   6165
      Left            =   60
      Top             =   60
      Width           =   9405
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   180
      Picture         =   "about.frx":0481
      Top             =   5250
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   90
      Picture         =   "about.frx":0E6B
      Top             =   90
      Width           =   9345
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Label5_Click()
Unload Me
End Sub

Private Sub Label6_Click()
Unload Me
End Sub

Private Sub Label7_Click()
Unload Me
End Sub

Private Sub Label8_Click()
Unload Me
End Sub

Private Sub Picture1_Click()
Unload Me
End Sub
