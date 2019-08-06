VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form4"
   ScaleHeight     =   6225
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows ‚ÌŠù’è’l
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3600
      TabIndex        =   1
      Top             =   3960
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   975
      Left            =   2640
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1720
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "‚l‚r ‚oƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid60.TDBGrid TDBGrid1 
      Height          =   2535
      Left            =   1440
      OleObjectBlob   =   "Formx.frx":0000
      TabIndex        =   0
      Top             =   720
      Width           =   6255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call CConvert
End Sub

Private Sub Form_Load()
    cnstr = "Driver={SQL Server};SERVER=UGCH-003-4;" & "DATABASE=SPSDATA;UID=sa;PWD=5473nsk0036;"
    ssql = "select * from uriage"
    
    Adodc1.ConnectionString = cnstr
    Adodc1.RecordSource = ssql
    Adodc1.Refresh

    TDBGrid1.DataSource = Adodc1
    TDBGrid1.Refresh
End Sub


