VERSION 5.00
Begin VB.Form frmComment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4785
   Icon            =   "frmComment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   60
      ScaleHeight     =   435
      ScaleWidth      =   4665
      TabIndex        =   17
      Top             =   4350
      Width           =   4665
      Begin CodeLibrary.lvButtons_H lvOK 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   30
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "SAVE"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockHover       =   1
         cGradient       =   16711680
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmComment.frx":058A
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmComment.frx":0924
      End
      Begin CodeLibrary.lvButtons_H lvCancel 
         Height          =   375
         Left            =   3540
         TabIndex        =   4
         Top             =   30
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Caption         =   "Cancel"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LockHover       =   1
         cGradient       =   16711680
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmComment.frx":0A86
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmComment.frx":0E20
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "NEW Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1125
      Left            =   60
      TabIndex        =   14
      Top             =   3180
      Width           =   4665
      Begin VB.TextBox txtname 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1110
         TabIndex        =   0
         Top             =   330
         Width           =   3165
      End
      Begin VB.TextBox txtpass 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1110
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   3165
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   420
         Width           =   840
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0C0&
         Height          =   345
         Left            =   1080
         Top             =   300
         Width           =   3225
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         Height          =   345
         Left            =   1080
         Top             =   690
         Width           =   3225
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   60
      ScaleHeight     =   1965
      ScaleWidth      =   4665
      TabIndex        =   7
      Top             =   1170
      Width           =   4665
      Begin VB.Image Image2 
         Height          =   240
         Left            =   4350
         Picture         =   "frmComment.frx":0F82
         Top             =   1650
         Width           =   240
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         Height          =   1965
         Left            =   0
         Top             =   0
         Width           =   4665
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "and new password."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   510
         TabIndex        =   13
         Top             =   1590
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "application. So, you will need to provide new username"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   510
         TabIndex        =   12
         Top             =   1350
         Width           =   3885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OR this just the FIRST TIME that you launched this"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   510
         TabIndex        =   11
         Top             =   1110
         Width           =   3630
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   150
         Picture         =   "frmComment.frx":58AC
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "is saved is deleted or misplaced."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   510
         TabIndex        =   10
         Top             =   720
         Width           =   2280
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The Information File wherein the username and password"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   510
         TabIndex        =   9
         Top             =   480
         Width           =   4035
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   150
         Picture         =   "frmComment.frx":5E36
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This Info Message is caused by either of the following:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   150
         Width           =   3825
      End
      Begin VB.Image Image3 
         Height          =   720
         Left            =   3870
         Picture         =   "frmComment.frx":63C0
         Top             =   1170
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -240
      TabIndex        =   5
      Top             =   960
      Width           =   7125
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   4785
      TabIndex        =   2
      Top             =   0
      Width           =   4785
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security Information File NOT FOUND!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   4095
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   60
         Picture         =   "frmComment.frx":728A
         Top             =   60
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lvCancel_Click()
Unload Me
End Sub

Private Sub lvOK_Click()
If txtname.Text = "" Then
   MsgBox "A NEW Username is required.", vbCritical, "Username Error"
   Call HL(txtname)
   Exit Sub
End If
If ValidName(txtname.Text) = False Then
   MsgBox "Invalid Username." & vbNewLine & _
          "---------------------------------------------------" & vbNewLine & _
          "a Username must not consist of any of the following" & vbNewLine & _
          Replace(invFileName, "x", " "), vbCritical, "No Username Found."
   Exit Sub
End If
If txtpass.Text = "" Then
   MsgBox "A NEW Password is required.", vbCritical, "Password Error"
   Call HL(txtpass)
   Exit Sub
End If
If ValidName(txtpass.Text) = False Then
   MsgBox "Invalid Password." & vbNewLine & _
          "---------------------------------------------------" & vbNewLine & _
          "a Password must not consist of any of the following" & vbNewLine & _
          Replace(invFileName, "x", " "), vbCritical, "No Password Found."
   Exit Sub
End If
If SaveSecurityFile(crypt.Encode(txtname.Text), crypt.Encode(txtpass.Text)) = True Then
   MsgBox "A New Username and password is saved.", vbInformation, "Saved"
   Unload Me
Else
   MsgBox "Sorry Unable to Add New Security Information." & _
             vbNewLine & _
             "This error is caused by:" & vbNewLine & _
             "-----------------------------------------------------------------" & _
             vbNewLine & _
             "(x) Unable to Create Security Information due to windows file and folder authorization.", vbCritical, _
             "Security Information Creation Error."
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lvOK_Click
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then lvOK_Click
End Sub
