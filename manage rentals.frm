VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10935
   LinkTopic       =   "Form5"
   ScaleHeight     =   8520
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "RENTALS"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   2520
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
         Left            =   6000
         TabIndex        =   13
         Top             =   5400
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   12
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "UPDATE "
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD  "
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   1320
         Top             =   5400
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
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
         BackColor       =   -2147483640
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
         RecordSource    =   "Rentals"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Rockwell Condensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "manage rentals.frx":0000
         Height          =   2295
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16776960
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   19
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
            Size            =   9.75
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
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Romantic"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Romantic"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Romantic"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Romantic"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000002&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000002&
         Caption         =   "Monthly Rent"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000002&
         Caption         =   "House No."
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000002&
         Caption         =   "Property"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

con.Execute "INSERT INTO [Rentals] " & _
"([Property], [House No], [Rent], [Status]) " & _
"VALUES ('" & Replace(Text1.Text, "'", "''") & "', '" & Replace(Text2.Text, "'", "''") & "', " & _
Val(Text3.Text) & ", '" & Replace(Text4.Text, "'", "''") & "')"

MsgBox "Data added successfully", vbInformation, "Success"


Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

Text1.SetFocus

End Sub
Private Sub Command2_Click()

    If Text1.Text = "" Then
        MsgBox "Please select a rental to update.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrHandler

    con.Execute "UPDATE [Rentals] SET " & _
        "[Property] = '" & Replace(Text1.Text, "'", "''") & "', " & _
        "[House No] = '" & Replace(Text2.Text, "'", "''") & "', " & _
        "[Rent] = " & Val(Text3.Text) & ", " & _
        "[Status] = '" & Replace(Text4.Text, "'", "''") & "' " & _
        "WHERE [id] = " & Val(DataGrid1.Columns(0).Text)

    MsgBox "Data updated successfully.", vbInformation


    Adodc1.Refresh

    
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""

    Text1.SetFocus

    Exit Sub

ErrHandler:
    MsgBox "Error updating data: " & Err.Description, vbExclamation

End Sub


Private Sub Command3_Click()

If Text1.Text = "" Then
    MsgBox "Please select a rental to delete."
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this rental?", _
          vbYesNo + vbQuestion, "Delete Rental") = vbNo Then
    Exit Sub
End If

On Error GoTo ErrHandler

con.Execute "DELETE FROM [Rentals] WHERE [id] = " & _
Val(DataGrid1.Columns(0).Text)

MsgBox "Data deleted successfully."

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

Adodc1.Refresh

Exit Sub

ErrHandler:
MsgBox "Error deleting data: " & Err.Description

End Sub

Private Sub Command4_Click()
Form2.Show
End Sub

Private Sub DataGrid1_Click()

    On Error GoTo ErrHandler

    Text1.Text = DataGrid1.Columns(1).Text & ""
    Text2.Text = DataGrid1.Columns(2).Text & ""
    Text3.Text = DataGrid1.Columns(3).Text & ""
    Text4.Text = DataGrid1.Columns(4).Text & ""

    Exit Sub

ErrHandler:
    MsgBox "Error loading rental data: " & Err.Description

End Sub


