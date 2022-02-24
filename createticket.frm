VERSION 5.00
Begin VB.Form logticket 
   Caption         =   "Log a ticket"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   10905
   Begin VB.CommandButton cmdCancel 
      Caption         =   "cancel"
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "submit"
      Height          =   615
      Left            =   840
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
   Begin VB.ComboBox ComDept 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox Password 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Eid 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Department"
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "EmployeeId"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "logticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myconnection As New dbconnection
Dim myAuth As New EmployeeAuth
Dim check As Boolean

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdSubmit_Click()
Call myAuth.Authenticate(Eid.Text, Password.Text, ComDept.Text)
check = myAuth.IsAuthenticated
If ComDept.Text = "DEVOPS" And check Then
    frmgeneratereport.Show
    Unload Me
ElseIf ComDept.Text <> "DEVOPS" And check Then
    MDIHome.Show
    Unload Me
Else
    MsgBox "Wrong"
End If
    frmgeneratereport.Show
End Sub

Private Sub Form_Load()
 Dim mydept As New dept
myconnection.SetUpConnection
Dim item() As String
item = mydept.getDep()

For i = 0 To UBound(item)
ComDept.AddItem item(i)
Next

End Sub







