VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Caption         =   "manage lardlord"
      Height          =   6975
      Left            =   1800
      TabIndex        =   0
      Top             =   1800
      Width           =   8175
      Begin VB.CommandButton Command4 
         Caption         =   "DASHBOARD"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   13
         Top             =   6480
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Landlords.frx":0000
         Height          =   2775
         Left            =   240
         TabIndex        =   12
         Top             =   3480
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         BackColor       =   16776960
         BorderStyle     =   0
         ForeColor       =   32768
         HeadLines       =   1
         RowHeight       =   18
         RowDividerStyle =   6
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Stencil"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Romantic"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   1800
         Top             =   6360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1085
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\summit\Desktop\House rental\Management.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\summit\Desktop\House rental\Management.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Landlords"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command3 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Bell MT"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
         Caption         =   "Addrress"
         BeginProperty Font 
            Name            =   "Romantic"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         Caption         =   "Email Address"
         BeginProperty Font 
            Name            =   "Romantic"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Romantic"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   "FullNames"
         BeginProperty Font 
            Name            =   "Romantic"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error GoTo ErrHandler

con.Execute "INSERT INTO [Landlords] " & _
"([Fullname], [phone], [Email], [Address]) " & _
"VALUES ('" & Replace(Text1.Text, "'", "''") & "', '" & _
Replace(Text2.Text, "'", "''") & "', '" & _
Replace(Text3.Text, "'", "''") & "', '" & _
Replace(Text4.Text, "'", "''") & "')"

MsgBox "Landlord added successfully."

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

Text1.SetFocus

Adodc1.Refresh

Exit Sub

ErrHandler:
MsgBox "Error adding landlord: " & Err.Description

End Sub


Private Sub Command2_Click()

If Text1.Text = "" Then
    MsgBox "Please select a landlord to update."
    Exit Sub
End If

On Error GoTo ErrHandler

con.Execute "UPDATE [Landlords] SET " & _
"[Fullname] = '" & Replace(Text1.Text, "'", "''") & "', " & _
"[phone] = '" & Replace(Text2.Text, "'", "''") & "', " & _
"[Email] = '" & Replace(Text3.Text, "'", "''") & "', " & _
"[Address] = '" & Replace(Text4.Text, "'", "''") & "' " & _
"WHERE [landlordID] = " & Adodc1.Recordset.Fields("landlordID").Value

MsgBox "Landlord updated successfully."

Adodc1.Refresh

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

Text1.SetFocus

Exit Sub

ErrHandler:
MsgBox "Error updating landlord: " & Err.Description

End Sub


Private Sub Command3_Click()

If Text1.Text = "" Then
    MsgBox "Please select a landlord to delete."
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this landlord?", _
          vbYesNo + vbQuestion, "Delete Landlord") = vbNo Then
    Exit Sub
End If

On Error GoTo ErrHandler

con.Execute "DELETE FROM [Landlords] WHERE [landlordID] = " & _
Adodc1.Recordset.Fields("landlordID").Value

MsgBox "Landlord deleted successfully."

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

Adodc1.Refresh

Exit Sub

ErrHandler:
MsgBox "Error deleting landlord: " & Err.Description

End Sub


Private Sub Command4_Click()
Form2.Show
End Sub

Private Sub DataGrid1_Click()

On Error GoTo ErrHandler

Dim row As Integer

row = DataGrid1.row

If row < 0 Then Exit Sub

Text1.Text = DataGrid1.Columns(1).Text
Text2.Text = DataGrid1.Columns(2).Text
Text3.Text = DataGrid1.Columns(3).Text
Text4.Text = DataGrid1.Columns(4).Text

Exit Sub

ErrHandler:
MsgBox "Error loading landlord: " & Err.Description

End Sub


