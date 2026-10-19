VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form8 
   BackColor       =   &H80000002&
   Caption         =   "Form8"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000B&
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
      Height          =   615
      Left            =   8160
      TabIndex        =   2
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Caption         =   "FILTER BY STATUS"
      BeginProperty Font 
         Name            =   "Bernard MT Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8160
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3000
      Top             =   4800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   32768
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
      RecordSource    =   $"REPORTS.frx":0000
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
      Bindings        =   "REPORTS.frx":001D
      Height          =   3855
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6800
      _Version        =   393216
      BackColor       =   65280
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   22
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
         Size            =   11.25
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
   Begin VB.Label Label2 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSACTIONS :"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   7080
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL REVENUE:"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   6120
      Width           =   3735
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error GoTo LoadErr
    
    con.CursorLocation = adUseClient
    
    Dim sql As String
    sql = "SELECT * FROM [PAYMENTS]"
    
    If Adodc1.Recordset.State = 1 Then Adodc1.Recordset.Close
    Adodc1.Recordset.Open sql, con, adOpenStatic, adLockOptimistic
    
    Set DataGrid1.DataSource = Adodc1
    
    CalculateTotals
    Exit Sub
    
LoadErr:
    MsgBox "Error initializing report data: " & Err.Description, vbCritical
End Sub

Private Sub CalculateTotals()
    Dim totalRevenue As Double
    Dim recordCount As Long
    Dim rsTemp As ADODB.Recordset
    
    totalRevenue = 0
    recordCount = 0
    
    Set rsTemp = Adodc1.Recordset.Clone
    
    If Not (rsTemp.BOF And rsTemp.EOF) Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If UCase(rsTemp.Fields("status").Value) = "PAID" Then
                totalRevenue = totalRevenue + Val(rsTemp.Fields("amount").Value)
            End If
            recordCount = recordCount + 1
            rsTemp.MoveNext
        Loop
    End If
    
    Label1.Caption = "Total Revenue: KSh " & Format(totalRevenue, "#,##0.00")
    Label2.Caption = "Total Transactions: " & recordCount
    
    Set rsTemp = Nothing
End Sub

Private Sub Command1_Click()
    Dim filterStatus As String
    Dim sql As String
    
    filterStatus = InputBox("Enter status to filter by (Paid, Partial, Pending) or leave blank to show all:", "Filter Report")
    
    If Trim(filterStatus) = "" Then
        sql = "SELECT * FROM [PAYMENTS]"
    Else
        filterStatus = Replace(filterStatus, "'", "''")
        sql = "SELECT * FROM [PAYMENTS] WHERE [status] = '" & filterStatus & "'"
    End If
    
    On Error GoTo FilterErr
    If Adodc1.Recordset.State = 1 Then Adodc1.Recordset.Close
    Adodc1.Recordset.Open sql, con, adOpenStatic, adLockOptimistic
    
    Set DataGrid1.DataSource = Nothing
    Set DataGrid1.DataSource = Adodc1
    
    CalculateTotals
    Exit Sub
    
FilterErr:
    MsgBox "Error filtering reports: " & Err.Description, vbCritical
End Sub

Private Sub Command2_Click()
    Form2.Show
    Unload Me
End Sub

