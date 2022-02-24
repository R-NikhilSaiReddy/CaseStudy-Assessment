VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubmitTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim count As Integer
Private arr As String

Public Function save(ByVal loggedBy As String, ByVal rDate As String, ByVal severity As String, ByVal Desc As String) As String
Dim sql As String
Dim message As String
Dim myRecSet As ADODB.Recordset

On Error GoTo errorhandler

sql = "SET DATEFORMAT DMY" & vbCrLf & _
        "insert into TICKET (LOGGED_BY, RAISED_Date, Severity, TICKET_Desc, ticket_Status) values('" & _
           loggedBy & "','" & rDate & "','" & severity & "','" & Desc & "', 'OPEN')"
           
Call myConnection.Execute(sql)

Set myRecSet = myConnection.Execute("SELECT t.Ticket_Id FROM Ticket as t")
    While Not myRecSet.EOF
        count = count + 1
        myRecSet.MoveNext
    Wend

message = "Ticket " & CStr(count) & " is submitted Successfully "
save = message
Exit Function
errorhandler:
    message = "Error! " & Err.Number & ": " & Err.Description
    save = message
End Function

Public Function update(ByVal Ticket_Id As Integer, ByVal resolvedBy As String, ByVal Resolution As String) As String
Dim sql As String
Dim message As String
Dim myRecSet As ADODB.Recordset

On Error GoTo errorhandler

sql = "SET DATEFORMAT DMY" & vbCrLf & _
 "update TICKET set Resolved_By='" & _
           resolvedBy & "',Resolution='" & Resolution & "',Resolved_Date=Format(GETDATE(), 'dd-MM-yyyy hh:mm'),ticket_status='CLOSED'" & _
           "Where Ticket_ID =" & Ticket_Id
           
Call myConnection.Execute(sql)

Set myRecSet = myConnection.Execute("SELECT t.Ticket_Id FROM Ticket as t")
    While Not myRecSet.EOF
        count = count + 1
        myRecSet.MoveNext
    Wend

message = "Ticket " & CStr(count) & " is closed"
update = message
Exit Function
errorhandler:
    message = "Error! " & Err.Number & ": " & Err.Description
    update = message
End Function


Public Function getTkt(ByVal Ticket_Id As Integer) As String
Dim sql As String
Dim message As String
Dim myRecSet As ADODB.Recordset


sql = "select * from Ticket " & _
           "Where Ticket_ID =" & Ticket_Id
           
Set myRecSet = myConnection.Execute(sql)

While Not myRecSet.EOF
        arr = arr & myRecSet("Logged_By").Value & vbCrLf & vbCrLf
        arr = arr & myRecSet("Raised_Date").Value & vbCrLf & vbCrLf
        arr = arr & myRecSet("Severity").Value & vbCrLf & vbCrLf
        arr = arr & myRecSet("Ticket_Desc").Value & vbCrLf & vbCrLf
        arr = arr & myRecSet("Resolved_By").Value & vbCrLf & vbCrLf
        arr = arr & myRecSet("Resolution").Value & vbCrLf & vbCrLf
        arr = arr & myRecSet("Resolved_Date").Value & vbCrLf & vbCrLf
        arr = arr & myRecSet("Status").Value & vbCrLf & vbCrLf

        myRecSet.MoveNext
    Wend


getTkt = arr
End Function





