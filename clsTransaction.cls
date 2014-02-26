VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Class_Terminate()
    Set rs = Nothing
    Set con = Nothing
End Sub

Public Function GetItems(bcode As String, qty As Integer, grd As MSFlexGrid)
    If rs.State = 1 Then rs.Close
    
    SQL = "SELECT * FROM products WHERE barcode = '" & bcode & "'"
    rs.Open SQL, con
    
    If Not rs.EOF Then
        With grd
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .Col = 0: .Text = rs!barcode
            .Col = 1: .Text = qty: .CellAlignment = 4
            .Col = 2: .Text = rs!Description
            .Col = 3: .Text = Format(rs!sellingprice, "#,###.00")
            
            Dim total As Double
            
            total = CDbl(qty) * rs!sellingprice
            
            .Col = 4: .Text = Format(total, "#,###.00")
        End With
    End If
End Function

Public Function AutoGenerateORNumber() As Integer
    If rs.State = 1 Then rs.Close
    
    SQL = "SELECT MAX(ORNumber) AS MaxOfORNumber FROM MainTrans"
    rs.Open SQL, con
    
    AutoGenerateORNumber = IIf(IsNull(rs!MaxOfORNumber), 1000, rs!MaxOfORNumber + 1)
End Function

Public Function SaveTransaction(grd As MSFlexGrid, cash As Double, _
            totalamount As Double, change As Double, cashier As String)
            
    ORNum = AutoGenerateORNumber
    
    'con.BeginTrans
    SQL = "INSERT INTO MainTrans(ORNumber,TransDate,TransTime,CashTendered, " & _
            "TotalAmount,Change,Cashier) VALUES(" & ORNum & ",#" & Date & "#, " & _
            "#" & Time & "#," & cash & "," & totalamount & "," & change & ", " & _
            "'" & cashier & "')"
    con.Execute (SQL)
    
    Dim ctr As Integer
    
    If rs.State = 1 Then rs.Close
    
    SQL = "SELECT * FROM SubTrans WHERE ORNumber = " & ORNum & ""
    rs.Open SQL, con
    
    Dim curStockOnHand      As Integer 'ang ara sa inventory
    Dim newStockOnHand      As Integer 'curStockOnHand buhinan sang ginbakal
    Dim ginbakal            As Integer 'quantity
    
    If rs.EOF Then 'empty
        With grd
            For ctr = 1 To grd.Rows - 1
                .Row = ctr
                
                rs.AddNew
                    rs!ORNumber = ORNum
                    .Col = 1: rs!qty = .Text: ginbakal = .Text
                    
                    .Col = 0: rs!bcode = .Text
                    
                    If rs1.State = 1 Then rs1.Close
                    SQL = "SELECT * FROM products WHERE barcode ='" & .Text & "'"
                    rs1.Open SQL, con, adOpenStatic
                    
                    If Not rs1.EOF Then
                        curStockOnHand = rs1!soh
                        newStockOnHand = curStockOnHand - ginbakal
                        
                        SQL = "UPDATE products SET soh=" & newStockOnHand & " " & _
                                "WHERE barcode = '" & .Text & "'"
                        con.Execute (SQL)
                    End If
                    
                    .Col = 4: rs!subtotal = .Text
                rs.Update
                
            Next
        End With
    End If
    'con.CommitTrans 'save all transactions at once
    
    MsgBox "Transaction was successfully saved!    ", vbInformation, "POS v.1000"
End Function

Public Function PreviewResibo(myDE As de)
    If myDE.rsresibo_Grouping.State = 1 Then myDE.rsresibo_Grouping.Close
    
    myDE.resibo_Grouping ORNum
    
    drResibo.Show
End Function
















