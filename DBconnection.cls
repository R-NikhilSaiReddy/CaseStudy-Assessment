VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBconnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function SetUpConnection()
    Set myConnection = New ADODB.Connection
    Set myRec = New ADODB.Recordset
    myConnection.ConnectionString = "DRIVER=SQL Server;SERVER=DESKTOP-KDF0NTM;DATABASE=Tickettrackingsystem;"
    Call myConnection.Open
    
    SetUpConnection = myConnection
End Function

