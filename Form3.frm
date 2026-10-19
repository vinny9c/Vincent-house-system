VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   20
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox Text6 
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
      Left            =   5640
      TabIndex        =   19
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text5 
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
      Left            =   5640
      TabIndex        =   18
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox Text2 
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
      Left            =   2400
      TabIndex        =   15
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "Management properties"
      ForeColor       =   &H80000001&
      Height          =   10695
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   12015
      Begin VB.CommandButton Command5 
         Caption         =   "BACK TO DASHBOARD"
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
         Left            =   7680
         TabIndex        =   22
         Top             =   2880
         Width           =   3495
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   1095
         Left            =   4440
         Top             =   9120
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1931
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
         RecordSource    =   "Management"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form3.frx":0000
         Height          =   4815
         Left            =   960
         TabIndex        =   21
         Top             =   4200
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8493
         _Version        =   393216
         Appearance      =   0
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
            Size            =   0
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         TabIndex        =   17
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox Text3 
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
         Left            =   840
         TabIndex        =   16
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   14
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
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
         Left            =   9840
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
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
         Height          =   615
         Left            =   7440
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
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
         Left            =   9840
         TabIndex        =   9
         Top             =   600
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
         Left            =   7440
         TabIndex        =   8
         Top             =   720
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
         Left            =   3720
         TabIndex        =   13
         Top             =   10320
         Width           =   6015
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000002&
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
         Left            =   4560
         TabIndex        =   7
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000002&
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
         Left            =   4560
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000002&
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
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000002&
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
         Left            =   840
         TabIndex        =   4
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000002&
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
         Left            =   840
         TabIndex        =   3
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000002&
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
         Left            =   840
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000002&
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
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   495
      Left            =   6960
      TabIndex        =   12
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

con.Execute "INSERT INTO [Management] " & _
"([property name], [property type], [location], [rent amount], [status], [description], [date]) " & _
"VALUES ('" & Text1.Text & "', '" & Text2.Text & "', '" & Text3.Text & "', " & _
Val(Text4.Text) & ", '" & Text5.Text & "', '" & Text6.Text & "', #" & Text7.Text & "#)"

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

If Text1.Text = "" Then
    MsgBox "Please select a property to delete."
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this property?", _
          vbYesNo + vbQuestion, "Delete Property") = vbNo Then
    Exit Sub
End If

On Error GoTo ErrHandler

con.Execute "DELETE FROM [Management] WHERE [property id] = " & _
Val(DataGrid1.Columns(0).Text)

MsgBox "Data deleted successfully."

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""

Adodc1.Refresh

Exit Sub

ErrHandler:
MsgBox "Error deleting data: " & Err.Description

End Sub





Private Sub Command3_Click()
       Dim searchCrit As String
    Dim sql As String
    
    searchCrit = InputBox("Enter property name, type, or location to search:", "Search Property")
    
    If Trim(searchCrit) = "" Then
        sql = "SELECT * FROM [Management]"
    Else
        searchCrit = Replace(searchCrit, "'", "''")
        sql = "SELECT * FROM [Management] WHERE " & _
              "[property name] LIKE '%" & searchCrit & "%' OR " & _
              "[property type] LIKE '%" & searchCrit & "%' OR " & _
              "[location] LIKE '%" & searchCrit & "%'"
    End If
    
    On Error GoTo SearchError
    
    If Adodc1.Recordset.State = 1 Then Adodc1.Recordset.Close
    
    Adodc1.Recordset.Open sql, con, 3, 3
    
    Set DataGrid1.DataSource = Nothing
    Set DataGrid1.DataSource = Adodc1
    
    Exit Sub

SearchError:
    MsgBox "Error during search: " & Err.Description, vbCritical
End Sub

Private Sub Command4_Click()

If Text1.Text = "" Then
    MsgBox "Please select a property to update."
    Exit Sub
End If

On Error GoTo ErrHandler

con.Execute "UPDATE [Management] SET " & _
"[property name] = '" & Replace(Text1.Text, "'", "''") & "', " & _
"[property type] = '" & Replace(Text2.Text, "'", "''") & "', " & _
"[location] = '" & Replace(Text3.Text, "'", "''") & "', " & _
"[rent amount] = " & Val(Text4.Text) & ", " & _
"[status] = '" & Replace(Text5.Text, "'", "''") & "', " & _
"[description] = '" & Replace(Text6.Text, "'", "''") & "', " & _
"[date] = #" & Replace(Text7.Text, "'", "''") & "# " & _
"WHERE [property id] = " & Val(DataGrid1.Columns(0).Text)

MsgBox "Data updated successfully."

Adodc1.Refresh

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""

Text1.SetFocus

Exit Sub

ErrHandler:
MsgBox "Error updating data: " & Err.Description

End Sub




Private Sub Command5_Click()
Form2.Show
End Sub

Private Sub DataGrid1_Click()

On Error GoTo ErrHandler

If Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF Then Exit Sub

If DataGrid1.Columns.Count >= 8 Then
    Text1.Text = DataGrid1.Columns(1).Text
    Text2.Text = DataGrid1.Columns(2).Text
    Text3.Text = DataGrid1.Columns(3).Text
    Text4.Text = DataGrid1.Columns(4).Text
    Text5.Text = DataGrid1.Columns(5).Text
    Text6.Text = DataGrid1.Columns(6).Text
    Text7.Text = DataGrid1.Columns(7).Text
End If

Exit Sub

ErrHandler:
MsgBox "Error loading property: " & Err.Description


End Sub


