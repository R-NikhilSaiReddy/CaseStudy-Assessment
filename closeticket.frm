VERSION 5.00
Begin VB.Form closeticket 
   Caption         =   "Closing a Ticket"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   10905
   Begin VB.CommandButton cmdCancel 
      Caption         =   "cancel"
      Height          =   615
      Left            =   3120
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "submit"
      Height          =   615
      Left            =   1320
      TabIndex        =   6
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox Resolution 
      Height          =   1095
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ComboBox ComResolvedBy 
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ComboBox ComTicketId 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Resolution"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Resolved By"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Ticket ID"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "closeticket"
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

msg = myTicket.Update(CInt(ComTicketId.Text), ComResolvedBy.Text, Resolution.Text)

MsgBox msg

End Sub

Private Sub Form_Load()
 Dim mydept As New dept
  Dim mydept2 As New dept
  
myconnection.SetUpConnection

Dim item() As String
Dim item2() As String

item = mydept.getEmpD()
item2 = mydept2.getTkt()

For i = 0 To UBound(item)
ComResolvedBy.AddItem item(i)
Next

For i = 0 To UBound(item2)
ComTicketId.AddItem item2(i)
Next
End Sub

