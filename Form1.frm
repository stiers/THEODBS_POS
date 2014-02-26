VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   1830
   ClientTop       =   1185
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11775
   Begin VB.Timer tmrDulaFrame 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3840
      Top             =   840
   End
   Begin VB.Timer tmrBlinker 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4920
      Top             =   720
   End
   Begin VB.Frame fraCash 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   3360
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtCT 
         Alignment       =   1  'Right Justify
         Height          =   495
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblCCD 
         Alignment       =   2  'Center
         Caption         =   "Close the Cash Drawer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblChange 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblTA 
         Alignment       =   1  'Right Justify
         Caption         =   "0.00"
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Change:"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Total Amount:"
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Cash Tendered:"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.TextBox txtBcode 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   0
      Top             =   1440
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid grdItems 
      Height          =   4095
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7223
      _Version        =   393216
   End
   Begin VB.Label lblTotalAmount 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   3
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Total Amount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   6360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call CustomizeGridHeader
End Sub

Private Sub CustomizeGridHeader()
    With grdItems
        .Rows = 1: .Cols = 5
        .Row = 0
        .Col = 1: .Text = "QTY": .ColWidth(1) = 1000: .CellAlignment = 4
        .Col = 2: .Text = "DESCRIPTION": .ColWidth(2) = 5500: .CellAlignment = 4
        .Col = 3: .Text = "PRICE": .ColWidth(3) = 1500: .CellAlignment = 4
        .Col = 4: .Text = "TOTAL": .ColWidth(4) = 1700: .CellAlignment = 4
    End With
End Sub

Private Sub tmrBlinker_Timer()
    lblCCD.Visible = Not lblCCD.Visible = True
End Sub

Private Sub tmrDulaFrame_Timer()
    fraCash.Visible = False
    Call CustomizeGridHeader
    lblTotalAmount.Caption = "0.00"
    txtCT.Text = ""
    lblTA.Caption = "0.00"
    lblChange.Caption = "0.00"
    
    tmrBlinker.Enabled = False
    tmrDulaFrame.Enabled = False
End Sub

Private Sub txtBcode_KeyDown(KeyCode As Integer, Shift As Integer)
    'MsgBox KeyCode
    Dim x() As String
    Dim ctr As Integer
    Dim curQTY As Integer
    Dim sprice As Double
    Dim newQTY As Integer
    Dim newSubTotal As Double
    
    If KeyCode = 13 Then 'enter
        
        On Error GoTo ambot
        
        x = Split(txtBcode.Text, "*")
        
            With grdItems
                If .Rows > 1 Then
                    For ctr = 1 To .Rows - 1
                        .Row = ctr
                        
                        .Col = 0 ' barcode
                        If .Text = x(1) Then 'x(1) unod sini is barcode
                            .Col = 1: curQTY = CInt(.Text)
                            .Col = 3: sprice = CDbl(.Text)
                            
                            newQTY = curQTY + x(0)
                            newSubTotal = newQTY * sprice
                            
                            'paslak ta sa grid
                            .Col = 1: .Text = newQTY
                            .Col = 4: .Text = Format(newSubTotal, "#,###.00")
                            
                            GoTo ContHere
                        End If
                    Next
                End If
            End With
            
        Call trans.GetItems(x(1), CInt(x(0)), grdItems)
ContHere:
        lblTotalAmount.Caption = Format(ComputeTotalAmount, "#,###.00")
        txtBcode.Text = ""
    
    ElseIf KeyCode = 112 Then 'F1
        lblTA.Caption = Format(lblTotalAmount.Caption, "#,###.00")
        fraCash.Visible = True
        txtCT.SetFocus
    End If
    
    Exit Sub

ambot:
        With grdItems
            If .Rows > 1 Then
                For ctr = 1 To .Rows - 1
                        .Row = ctr
                        
                        .Col = 0 ' barcode
                        If .Text = txtBcode.Text Then 'x(1) unod sini is barcode
                            .Col = 1: curQTY = CInt(.Text)
                            .Col = 3: sprice = CDbl(.Text)
                            
                            newQTY = curQTY + 1
                            newSubTotal = newQTY * sprice
                            
                            'paslak ta sa grid
                            .Col = 1: .Text = newQTY
                            .Col = 4: .Text = Format(newSubTotal, "#,###.00")
                            
                            GoTo ContHere1
                        End If
                    Next
                End If
            End With

    Call trans.GetItems(txtBcode.Text, 1, grdItems)
ContHere1:
    lblTotalAmount.Caption = Format(ComputeTotalAmount, "#,###.00")
    txtBcode.Text = ""
End Sub

Private Function ComputeTotalAmount() As Double
    Dim total As Double
    Dim ctr As Integer
    
    With grdItems
        For ctr = 1 To .Rows - 1
            .Row = ctr
            .Col = 4: total = .Text
            
            ComputeTotalAmount = ComputeTotalAmount + CDbl(total)
        Next
    End With
End Function

Private Sub txtCT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then 'esc key
        fraCash.Visible = False
        txtBcode.SetFocus
    ElseIf KeyCode = 13 Then 'enter
        Dim cash As Double
        Dim change As Double
        Dim ta As Double 'total amount
        
        cash = CDbl(txtCT.Text)
        ta = CDbl(lblTA.Caption)
        
        If cash < ta Then
            MsgBox "Abno ka ba?  Kulang kwarta mo!   ", vbExclamation, "Heller"
            txtCT.SelStart = 0
            txtCT.SelLength = Len(txtCT.Text)
            Exit Sub
        End If
        
        change = cash - ta
        
        lblChange.Caption = Format(change, "#,###.00")
        
        Call trans.SaveTransaction(grdItems, txtCT.Text, lblTA.Caption, _
                      lblChange.Caption, "Julius")
        
        Call trans.PreviewResibo(de)
        
        tmrBlinker.Enabled = True
        tmrDulaFrame.Enabled = True
    End If
End Sub












