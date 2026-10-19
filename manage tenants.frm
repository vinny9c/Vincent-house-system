VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Manage tenants"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   840
      TabIndex        =   0
      Top             =   -120
      Width           =   8895
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
         Left            =   6840
         TabIndex        =   15
         Top             =   7800
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   1440
         Top             =   7800
         Width           =   4935
         _ExtentX        =   8705
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
         RecordSource    =   "Manage tenants"
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
         Bindings        =   "manage tenants.frx":0000
         Height          =   2415
         Left            =   240
         TabIndex        =   14
         Top             =   5160
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4260
         _Version        =   393216
         BackColor       =   16776960
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   18
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
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
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
         TabIndex        =   13
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000008&
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
         Height          =   495
         Left            =   7440
         TabIndex        =   12
         Top             =   3480
         Width           =   975
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
         Height          =   495
         Left            =   7440
         TabIndex        =   11
         Top             =   2520
         Width           =   975
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
         Height          =   495
         Left            =   7440
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
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
         Left            =   3480
         TabIndex        =   9
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
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
         Left            =   3480
         TabIndex        =   8
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
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
         TabIndex        =   7
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
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
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000002&
         Caption         =   " House No."
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
         Left            =   3840
         TabIndex        =   5
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000002&
         Caption         =   "  Property"
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
         Left            =   3840
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000002&
         Caption         =   "  Email"
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
         Left            =   840
         TabIndex        =   3
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000002&
         Caption         =   "  Phone"
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
         Left            =   840
         TabIndex        =   2
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000002&
         Caption         =   " Tenants     name"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   1080
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

con.Execute "INSERT INTO [Manage] " & _
"([Tenants name], [phone no], [Email], [Property], [House number]) " & _
"VALUES ('" & Text1.Text & "', '" & Text2.Text & "', '" & Text3.Text & "', '" & _
Text4.Text & "', '" & Text5.Text & "')"

MsgBox "Data added successfully"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

Text1.SetFocus

End Sub
Private Sub Command2_Click()

If Text1.Text = "" Then
    MsgBox "Please select a tenant to update."
    Exit Sub
End If

On Error GoTo ErrHandler

con.Execute "UPDATE [Manage] SET " & _
"[Tenants name] = '" & Replace(Text1.Text, "'", "''") & "', " & _
"[phone no] = '" & Replace(Text2.Text, "'", "''") & "', " & _
"[Email] = '" & Replace(Text3.Text, "'", "''") & "', " & _
"[Property] = '" & Replace(Text4.Text, "'", "''") & "', " & _
"[House number] = '" & Replace(Text5.Text, "'", "''") & "' " & _
"WHERE [id] = " & Adodc1.Recordset.Fields("id").Value

MsgBox "Data updated successfully."

Adodc1.Refresh

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

Text1.SetFocus

Exit Sub

ErrHandler:
MsgBox "Error updating data: " & Err.Description

End Sub

Private Sub Command3_Click()

If Text1.Text = "" Then
    MsgBox "Please select a tenant to delete."
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this tenant?", _
          vbYesNo + vbQuestion, "Delete Tenant") = vbNo Then
    Exit Sub
End If

On Error GoTo ErrHandler

con.Execute "DELETE FROM [Manage] WHERE [id] = " & _
Adodc1.Recordset.Fields("id").Value

MsgBox "Data deleted successfully."

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

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

    Dim row As Integer
    row = DataGrid1.row

    If row < 0 Then Exit Sub

    Text1.Text = DataGrid1.Columns(1).Text
    Text2.Text = DataGrid1.Columns(2).Text
    Text3.Text = DataGrid1.Columns(3).Text
    Text4.Text = DataGrid1.Columns(4).Text
    Text5.Text = DataGrid1.Columns(5).Text

    Exit Sub

ErrHandler:
    MsgBox "Error loading tenant: " & Err.Description

End Sub

