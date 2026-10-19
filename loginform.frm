VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "RENTAL MANAGEMENT SYSTEM"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   4200
      TabIndex        =   0
      Top             =   2640
      Width           =   4335
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000002&
         Caption         =   "Register"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000002&
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         MaskColor       =   &H80000002&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "RENTAL MANAGEMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   1560
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim rs As ADODB.Recordset
    Dim sql As String

    On Error GoTo LoginError

    'Check username
    If Trim(Text1.Text) = "" Then

        MsgBox "Please enter your username.", _
               vbExclamation, "Login"

        Text1.SetFocus
        Exit Sub

    End If

    'Check password
    If Trim(Text2.Text) = "" Then

        MsgBox "Please enter your password.", _
               vbExclamation, "Login"

        Text2.SetFocus
        Exit Sub

    End If

    'Connect to database if necessary
    If con Is Nothing Then
        connectDB
    ElseIf con.State = adStateClosed Then
        connectDB
    End If

    'SQL query
    sql = "SELECT * FROM Users WHERE UserName='" & _
          Replace(Text1.Text, "'", "''") & _
          "' AND [Password]='" & _
          Replace(Text2.Text, "'", "''") & "'"

    'Create recordset
    Set rs = New ADODB.Recordset

    'Execute query
    rs.Open sql, con, adOpenStatic, adLockReadOnly

    'Check login
    If rs.EOF Then

        MsgBox "Invalid username or password.", _
               vbCritical, "Login Failed"

        Text2.Text = ""
        Text2.SetFocus

    Else

        'Store logged-in username
        loggedinuser = rs!UserName

        MsgBox "Login successful. Welcome " & _
               loggedinuser, _
               vbInformation, "Login Successful"

        'Open dashboard
        Form2.Show

        'Hide login form
        Me.Hide

    End If

    'Close recordset
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            rs.Close
        End If

        Set rs = Nothing
    End If

    Exit Sub

LoginError:

    MsgBox "Login error: " & Err.Description, _
           vbCritical, "Database Error"

    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then
            rs.Close
        End If

        Set rs = Nothing
    End If

End Sub


Private Sub Command2_Click()
    Dim rsCheck As ADODB.Recordset
    Dim sqlCheck As String
    Dim sqlInsert As String
    

    If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Then
        MsgBox "Please enter both a username and a password to register.", vbExclamation, "Registration Failed"
        Exit Sub
    End If
    
    On Error GoTo RegError
    
    sqlCheck = "SELECT * FROM [users] WHERE [username] = '" & Replace(Text1.Text, "'", "''") & "'"
    
    Set rsCheck = New ADODB.Recordset
    rsCheck.CursorLocation = adUseClient
    rsCheck.Open sqlCheck, con, adOpenStatic, adLockReadOnly
    
    If Not (rsCheck.BOF And rsCheck.EOF) Then
        MsgBox "This username is already taken. Please choose another one.", vbCritical, "Registration Failed"
        rsCheck.Close
        Set rsCheck = Nothing
        Exit Sub
    End If
    
    rsCheck.Close
    Set rsCheck = Nothing
    

    sqlInsert = "INSERT INTO [users] ([username], [password]) " & _
                "VALUES ('" & Replace(Text1.Text, "'", "''") & "', '" & _
                Replace(Text2.Text, "'", "''") & "')"
                
    con.Execute sqlInsert
    
    MsgBox "Account registered successfully! You can now log in.", vbInformation, "Registration Success"
    

    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    
    Exit Sub

RegError:
    MsgBox "Error during registration: " & Err.Description, vbCritical, "Database Error"


End Sub

Private Sub Text2_Change()
  Text2.PasswordChar = "*"
End Sub
