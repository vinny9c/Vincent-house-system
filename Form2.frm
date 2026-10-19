VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Algerian"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "Payments"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   5
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000018&
      Caption         =   "Management properties"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      MaskColor       =   &H00FF0000&
      TabIndex        =   1
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Frame Fram 
      BackColor       =   &H80000002&
      Caption         =   "DASHBOARD"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   6135
      Left            =   3120
      TabIndex        =   0
      Top             =   1920
      Width           =   7215
      Begin VB.CommandButton Command7 
         Caption         =   "LOG OUT"
         BeginProperty Font 
            Name            =   "Bernard MT Condensed"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         TabIndex        =   6
         Top             =   3240
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Manage Landlord"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         TabIndex        =   4
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Manage Tenants"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Manage Rentals"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   2
         Top             =   2040
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form3.Show


End Sub

Private Sub Command2_Click()
Form5.Show
End Sub

Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Command4_Click()
Form6.Show
End Sub

Private Sub Command5_Click()
Form7.Show
End Sub

Private Sub Command6_Click()
Form8.Show
End Sub

Private Sub Command7_Click()
 If MsgBox("Are you sure you want to logout?", vbYesNo + vbQuestion, "Logout Confirmation") = vbYes Then
        Form1.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    Fram.Caption = "Welcome, " & loggedinuser

End Sub


