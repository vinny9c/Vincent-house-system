VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000002&
      Caption         =   "PAYMENTS"
      Height          =   8175
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   10335
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   1200
         Top             =   7440
         Width           =   5055
         _ExtentX        =   8916
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
         BackColor       =   16776960
         ForeColor       =   16711680
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\summit\Desktop\House rental\Management.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\summit\Desktop\House rental\Management.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "PAYMENTS"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Romantic"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "PAYMENTS.frx":0000
         Height          =   3135
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5530
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
         Left            =   8040
         TabIndex        =   14
         Top             =   7560
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9120
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
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
         Height          =   615
         Left            =   7680
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3480
         TabIndex        =   10
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   8
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   3240
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000002&
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000002&
         Caption         =   "DATE"
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
         Left            =   3480
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000002&
         Caption         =   "AMOUNT"
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
         TabIndex        =   5
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000002&
         Caption         =   "PROPERTY NAME"
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
         TabIndex        =   3
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000002&
         Caption         =   "TENANT NAME"
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
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error GoTo ErrHandler

    con.Execute "INSERT INTO [PAYMENTS] " & _
    "([tenantname], [propertyname], [amount], [date], [status]) " & _
    "VALUES ('" & Replace(Text1.Text, "'", "''") & "', '" & _
    Replace(Text2.Text, "'", "''") & "', " & _
    Val(Text3.Text) & ", #" & Text4.Text & "#, '" & _
    Replace(Text5.Text, "'", "''") & "')"

    MsgBox "Payment recorded successfully.", vbInformation

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""

    Text1.SetFocus
    Adodc1.Refresh
    Exit Sub

ErrHandler:
    MsgBox "Error adding payment: " & Err.Description, vbCritical
End Sub

Private Sub Command2_Click()
    If Text1.Text = "" Then
        MsgBox "Please select a payment record to delete.", vbExclamation
        Exit Sub
    End If

    If MsgBox("Are you sure you want to delete this payment record?", _
              vbYesNo + vbQuestion, "Delete Record") = vbNo Then
        Exit Sub
    End If

    On Error GoTo ErrHandler

    con.Execute "DELETE FROM [PAYMENTS] WHERE [PAYMENTID] = " & _
    Val(DataGrid1.Columns(0).Text)

    MsgBox "Payment record deleted successfully.", vbInformation

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""

    Adodc1.Refresh
    Exit Sub

ErrHandler:
    MsgBox "Error deleting record: " & Err.Description, vbCritical
End Sub

Private Sub Command3_Click()
    If Text1.Text = "" Then
        MsgBox "Please select a payment record to update.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrHandler

    con.Execute "UPDATE [PAYMENTS] SET " & _
    "[tenantname] = '" & Replace(Text1.Text, "'", "''") & "', " & _
    "[propertyname] = '" & Replace(Text2.Text, "'", "''") & "', " & _
    "[amount] = " & Val(Text3.Text) & ", " & _
    "[date] = #" & Text4.Text & "#, " & _
    "[status] = '" & Replace(Text5.Text, "'", "''") & "' " & _
    "WHERE [PAYMENTID] = " & Val(DataGrid1.Columns(0).Text)

    MsgBox "Payment record updated successfully.", vbInformation

    Adodc1.Refresh

    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""

    Text1.SetFocus
    Exit Sub

ErrHandler:
    MsgBox "Error updating record: " & Err.Description, vbCritical
End Sub

Private Sub Command4_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub DataGrid1_Click()
    On Error GoTo ErrHandler

    If Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF Then Exit Sub

    If DataGrid1.Columns.Count >= 6 Then
        Text1.Text = DataGrid1.Columns(1).Text
        Text2.Text = DataGrid1.Columns(2).Text
        Text3.Text = DataGrid1.Columns(3).Text
        Text4.Text = DataGrid1.Columns(4).Text
        Text5.Text = DataGrid1.Columns(5).Text
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error loading payment data: " & Err.Description, vbCritical
End Sub


