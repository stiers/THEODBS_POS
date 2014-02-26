Attribute VB_Name = "Module1"
Option Explicit

Public rs As New ADODB.Recordset
Public con As New ADODB.Connection
Public SQL As String

Public rs1 As New ADODB.Recordset

Public ORNum As Integer

Public trans As New clsTransaction

Public Sub Main()
    Set rs = New ADODB.Recordset
    Set con = New ADODB.Connection
    
    Set trans = New clsTransaction
    
    With rs
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With
    
    Dim dbPath As String
    Dim ConString As String
    
    dbPath = App.Path & "\pos.mdb"
    
    ConString = "Provider=Microsoft.Jet.OLEDB.4.0; " & _
                "Data Source=" & dbPath & ";Persist Security Info=False"
    
    con.Open ConString
    
    Form1.Show
End Sub

