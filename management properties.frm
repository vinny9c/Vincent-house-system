VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   8280
      TabIndex        =   22
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   8280
      TabIndex        =   21
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   8280
      TabIndex        =   20
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text2 O
      Height          =   495
      Left            =   3720
      TabIndex        =   17
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Caption         =   "HOUSE MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   960
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Management properties"
      Height          =   7335
      Left            =   1800
      TabIndex        =   0
      Top             =   1800
      Width           =   10335
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   1920
         TabIndex        =   19
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1920
         TabIndex        =   18
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1920
         TabIndex        =   16
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7440
         TabIndex        =   12
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         TabIndex        =   11
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5040
         TabIndex        =   10
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "EDIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   9
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD NEW"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2400
         TabIndex        =   8
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "@COPYRIGHT VINCENT MUSYOKA 2026"
         BeginProperty Font 
            Name            =   "Vineta BT"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         Top             =   6840
         Width           =   6015
      End
      Begin VB.Label Label8 
         Caption         =   "DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   7
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "DESCRIPTION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5280
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "RENT AMOUNT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "LOCATION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "PROPERTY TYPE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "PROPERTY NAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   495
      Left            =   6960
      TabIndex        =   13
      Top             =   5280
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    con.Execute "INSERT INTO [Management properties] " & _
        "([property name], [property type], [location], [rent amount], [status], [description], [date]) " & _
        "VALUES ('" & Text1.Text & "', '" & Text2.Text & "', '" & Text3.Text & "', '" & _
        Text4.Text & "', '" & Text5.Text & "', '" & Text6.Text & "', '" & Text7.Text & "')"

    MsgBox "Data added successfully"

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""

    Text1.SetFocus

End Sub



Private Sub Command2_Click()

    Dim answer As Integer

    answer = MsgBox("Are you sure you want to delete this property?", _
                    vbYesNo + vbQuestion, "Delete Property")

    If answer = vbYes Then

        con.Execute "DELETE FROM [Management properties] " & _
                    "WHERE [property name] = '" & Text1.Text & "'"

        MsgBox "Property deleted successfully", vbInformation, "Delete"

        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""

        Text1.SetFocus

    End If

End Sub

End Sub
