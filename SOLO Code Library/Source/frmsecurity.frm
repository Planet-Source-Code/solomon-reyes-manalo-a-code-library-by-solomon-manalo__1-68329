VERSION 5.00
Begin VB.Form frmsecurity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Form"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6405
   Icon            =   "frmsecurity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmsecurity.frx":0ECA
   ScaleHeight     =   2850
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
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
      Left            =   870
      TabIndex        =   6
      Top             =   1110
      Width           =   4665
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
         TabIndex        =   8
         Top             =   720
         Width           =   3165
      End
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
         TabIndex        =   7
         Top             =   330
         Width           =   3165
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         Height          =   345
         Left            =   1080
         Top             =   690
         Width           =   3225
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0C0&
         Height          =   345
         Left            =   1080
         Top             =   300
         Width           =   3225
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   420
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   780
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   870
      ScaleHeight     =   435
      ScaleWidth      =   4665
      TabIndex        =   3
      Top             =   2280
      Width           =   4665
      Begin CodeLibrary.lvButtons_H lvOK 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   2400
         TabIndex        =   4
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
         Image           =   "frmsecurity.frx":1D94
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmsecurity.frx":212E
      End
      Begin CodeLibrary.lvButtons_H lvCancel 
         Height          =   375
         Left            =   3540
         TabIndex        =   5
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
         Image           =   "frmsecurity.frx":2290
         cBack           =   -2147483633
         mPointer        =   99
         mIcon           =   "frmsecurity.frx":262A
      End
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -30
      TabIndex        =   2
      Top             =   900
      Width           =   6495
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
      ScaleWidth      =   6405
      TabIndex        =   0
      Top             =   0
      Width           =   6405
      Begin VB.Image Image1 
         Height          =   720
         Left            =   60
         Picture         =   "frmsecurity.frx":278C
         Top             =   60
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Existing Security Account"
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
         Left            =   780
         TabIndex        =   1
         Top             =   330
         Width           =   3660
      End
   End
End
Attribute VB_Name = "frmsecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Folder As New FolderSysObject
Dim File As New FileSysObject

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
Msg = MsgBox("Are you sure you want to change the current Security Information?" & _
             vbNewLine & _
             "----------------------------------------------------------------" & _
             vbNewLine & _
             "The program will be restarted, and you will need to RE-Login again.", _
             vbQuestion + vbYesNo, "Confirmation")

If Msg = vbNo Then
   Exit Sub
End If

If SaveSecurityFile(txtname.Text, txtpass.Text) = True Then
   MsgBox "A New Username and password is saved.", vbInformation, "Saved"
   Unload Me
   Unload mdiMain
   frmLogin.Show
   
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
