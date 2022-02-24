VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.MDIForm MDIHome 
   BackColor       =   &H8000000C&
   Caption         =   "TicketTrackingSystem"
   ClientHeight    =   7185
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10890
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   1296
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
   Begin VB.Menu Create 
      Caption         =   "Create a Ticket"
   End
   Begin VB.Menu log 
      Caption         =   "View a Ticket"
   End
   Begin VB.Menu close 
      Caption         =   "Close a Ticket"
   End
End
Attribute VB_Name = "MDIHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
closeticket.Show
End Sub

Private Sub Create_Click()
createticket.Show
End Sub

Private Sub log_Click()
logticket.Show
End Sub


Private Sub MDIForm_Load()
If ChkDevops = True Then
    mnuCreate.Visible = False
End If
End Sub

