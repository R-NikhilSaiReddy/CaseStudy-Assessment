VERSION 5.00
Begin VB.Form createticket 
   Caption         =   "Creating a ticket"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   11520
   Begin VB.CommandButton cmdCancel 
      Caption         =   "cancel"
      Height          =   735
      Left            =   3360
      TabIndex        =   9
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "submit"
      Height          =   735
      Left            =   1320
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Desc 
      Height          =   1095
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3240
      Width           =   2775
   End
   Begin VB.ComboBox CmbSeverity 
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Datepick 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
   End
   Begin VB.ComboBox ComEmployee 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Ticket Description"
      Height          =   735
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Severity"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Ticket Date-Time"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Employee"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "createticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myconnection As New dbconnection

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSubmit_Click()
Dim msg As String
Dim myTicket As New SubmitTicket



msg = myTicket.Save(ComEmployee.Text, Datepick.Text, CmbSeverity.Text, Desc.Text)

MsgBox msg
End Sub

Private Sub Form_Load()
Dim mydept As New dept
myconnection.SetUpConnection
Dim item() As String
item = mydept.getEmp()

For i = 0 To UBound(item)
ComEmployee.AddItem item(i)
Next

CmbSeverity.AddItem "Critical"
CmbSeverity.AddItem "Major"
CmbSeverity.AddItem "Normal"
End Sub


