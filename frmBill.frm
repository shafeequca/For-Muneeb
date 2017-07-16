VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBill 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bill Manager"
   ClientHeight    =   8685
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   12525
   Icon            =   "frmBill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   12525
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmBill.frx":0442
      Left            =   5400
      List            =   "frmBill.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   8160
      Width           =   1815
   End
   Begin VB.TextBox txtStateCode 
      Height          =   372
      Left            =   10560
      TabIndex        =   47
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtState 
      Height          =   372
      Left            =   8160
      TabIndex        =   46
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtIGST 
      Enabled         =   0   'False
      Height          =   372
      Left            =   2160
      TabIndex        =   44
      Top             =   7560
      Width           =   2175
   End
   Begin VB.TextBox txtIGSTPer 
      Height          =   372
      Left            =   1200
      TabIndex        =   43
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox txtGST 
      Enabled         =   0   'False
      Height          =   372
      Left            =   9360
      TabIndex        =   42
      Top             =   7080
      Width           =   2172
   End
   Begin VB.TextBox txtSGST 
      Enabled         =   0   'False
      Height          =   372
      Left            =   2160
      TabIndex        =   39
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox txtSGSTPer 
      Height          =   372
      Left            =   1200
      TabIndex        =   38
      Top             =   7080
      Width           =   735
   End
   Begin VB.ComboBox cboFormat 
      Height          =   315
      ItemData        =   "frmBill.frx":0474
      Left            =   9360
      List            =   "frmBill.frx":0481
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   8160
      Width           =   2175
   End
   Begin VB.ComboBox txtParty1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   3612
   End
   Begin VB.TextBox txtaddress 
      Height          =   735
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   3612
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   492
      Left            =   2160
      TabIndex        =   33
      Top             =   8040
      Width           =   2172
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   492
      Left            =   600
      TabIndex        =   32
      Top             =   8040
      Width           =   1332
   End
   Begin VB.TextBox txtGrandTotal 
      Enabled         =   0   'False
      Height          =   372
      Left            =   9360
      TabIndex        =   31
      Top             =   7560
      Width           =   2172
   End
   Begin VB.TextBox txtBTotal 
      Enabled         =   0   'False
      Height          =   372
      Left            =   9360
      TabIndex        =   30
      Top             =   6600
      Width           =   2172
   End
   Begin VB.TextBox txtCGST 
      Enabled         =   0   'False
      Height          =   372
      Left            =   2160
      TabIndex        =   29
      Top             =   6600
      Width           =   2175
   End
   Begin VB.TextBox txtCGSTPer 
      Height          =   372
      Left            =   1200
      TabIndex        =   28
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      Height          =   732
      Left            =   10320
      TabIndex        =   12
      Top             =   2880
      Width           =   1212
   End
   Begin VB.ComboBox cboItems 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2160
      TabIndex        =   8
      Top             =   3240
      Width           =   2892
   End
   Begin VB.TextBox txtTotal 
      Enabled         =   0   'False
      Height          =   372
      Left            =   8160
      TabIndex        =   11
      Top             =   3240
      Width           =   1932
   End
   Begin VB.TextBox txtRate 
      Height          =   372
      Left            =   6600
      TabIndex        =   10
      Top             =   3240
      Width           =   1572
   End
   Begin VB.TextBox txtQTY 
      Height          =   372
      Left            =   5040
      TabIndex        =   9
      Top             =   3240
      Width           =   1572
   End
   Begin VB.TextBox txtCode 
      Height          =   372
      Left            =   600
      TabIndex        =   7
      Top             =   3240
      Width           =   1572
   End
   Begin VB.TextBox txtVehicle 
      Height          =   372
      Left            =   8160
      TabIndex        =   6
      Top             =   2280
      Width           =   3372
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   372
      Left            =   8160
      TabIndex        =   5
      Top             =   1200
      Width           =   3372
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Format          =   94371841
      CurrentDate     =   42490
   End
   Begin VB.TextBox txtInvoice 
      Height          =   372
      Left            =   8160
      TabIndex        =   2
      Top             =   600
      Width           =   3372
   End
   Begin VB.TextBox txtPhone 
      Height          =   372
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   3612
   End
   Begin VB.TextBox txtGSTIN 
      Height          =   372
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
      Width           =   3612
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2652
      Left            =   600
      TabIndex        =   24
      Top             =   3840
      Width           =   10932
      _ExtentX        =   19288
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Format"
      Height          =   255
      Left            =   7440
      TabIndex        =   48
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "State and Code"
      Height          =   255
      Left            =   6600
      TabIndex        =   45
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "GST"
      Height          =   255
      Left            =   7440
      TabIndex        =   41
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "IGST"
      Height          =   255
      Left            =   600
      TabIndex        =   40
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "SGST"
      Height          =   255
      Left            =   600
      TabIndex        =   37
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Type"
      Height          =   255
      Left            =   4560
      TabIndex        =   35
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label slno 
      Height          =   372
      Left            =   11760
      TabIndex        =   34
      Top             =   2880
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "CGST"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      Height          =   255
      Left            =   7440
      TabIndex        =   26
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   252
      Left            =   7440
      TabIndex        =   25
      Top             =   6720
      Width           =   1452
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TOTAL"
      Height          =   372
      Left            =   8160
      TabIndex        =   23
      Top             =   2880
      Width           =   1932
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RATE"
      Height          =   372
      Left            =   6600
      TabIndex        =   22
      Top             =   2880
      Width           =   1572
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QTY"
      Height          =   372
      Left            =   5040
      TabIndex        =   21
      Top             =   2880
      Width           =   1572
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ITEM"
      Height          =   372
      Left            =   2160
      TabIndex        =   20
      Top             =   2880
      Width           =   2892
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ITEM CODE"
      Height          =   372
      Left            =   600
      TabIndex        =   19
      Top             =   2880
      Width           =   1572
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Number"
      Height          =   255
      Left            =   6600
      TabIndex        =   18
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   252
      Left            =   6600
      TabIndex        =   17
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number"
      Height          =   252
      Left            =   6600
      TabIndex        =   16
      Top             =   720
      Width           =   1452
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      Height          =   252
      Left            =   600
      TabIndex        =   15
      Top             =   2280
      Width           =   1452
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "GSTIN"
      Height          =   252
      Left            =   600
      TabIndex        =   14
      Top             =   1800
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Party"
      Height          =   252
      Left            =   600
      TabIndex        =   13
      Top             =   600
      Width           =   1452
   End
End
Attribute VB_Name = "frmBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboItems_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtQTY.SetFocus
End If
End Sub

Private Sub cmdAdd_Click()
If cboItems.Text = "" Then
    MsgBox "Please select an item"
    cboItems.SetFocus
ElseIf Val(txtQTY.Text) = 0 Or txtQTY.Text = "" Then
    MsgBox "Please enter quantity"
    txtQTY.SetFocus
ElseIf Val(txtRate.Text) = 0 Or txtRate.Text = "" Then
    MsgBox "Please enter rate"
    txtRate.SetFocus
Else
If slno.Caption = "" Then
    If rs.State = 1 Then rs.Close
        rs.Open "select count(*) from Bill", con, adOpenDynamic, adLockOptimistic
        Dim sl As Integer
        sl = Val(rs(0)) + 1
        con.Execute "insert into Bill values(2," & sl & ",'" & txtCode.Text & "','" & cboItems.Text & "'," & Val(txtQTY.Text) & "," & Val(txtRate.Text) & "," & Val(txtTotal.Text) & ",0,0,0,0,0,0,0,0)"
Else
    con.Execute "update Bill set ITEM_CODE='" & txtCode.Text & "',ITEM='" & cboItems.Text & "',QTY=" & Val(txtQTY.Text) & ",RATE=" & Val(txtRate.Text) & ",TOTAL=" & Val(txtTotal.Text) & " where SL_NO=" & Val(slno.Caption)
End If
        If rs.State = 1 Then rs.Close
        rs.Open "select * from BILL", con, adOpenDynamic, adLockOptimistic
        Set MSHFlexGrid1.DataSource = rs
        If rs.State = 1 Then rs.Close
        rs.Open "select sum(TOTAL) from BILL", con, adOpenDynamic, adLockOptimistic
        txtBTotal.Text = rs(0)
        clearItems
        cboItems.SetFocus
End If
End Sub
Private Sub clearItems()
txtCode.Text = ""
cboItems.Text = ""
txtQTY.Text = ""
txtRate.Text = ""
slno.Caption = ""
End Sub

Private Sub cmdNew_Click()
txtBTotal.Text = ""
clearItems
txtaddress.Text = ""
txtCGST.Text = ""
txtCGSTPer.Text = ""
txtGrandTotal.Text = ""
txtInvoice.Text = ""
txtPhone.Text = ""
txtGSTIN.Text = ""
txtVehicle.Text = ""
txtSGST.Text = ""
txtSGSTPer.Text = ""
txtIGST.Text = ""
txtIGSTPer.Text = ""
txtGST.Text = ""
txtState.Text = ""
txtStateCode = ""
Form_Load
txtParty1.SetFocus
End Sub

Private Sub cmdPrint_Click()

con.Execute "Update Bill SET CGSTPER='" & Val(txtCGSTPer.Text) & "',SGSTPER='" & Val(txtSGSTPer.Text) & "',IGSTPER='" & Val(txtIGSTPer.Text) & "'"
con.Execute "Update Bill SET CGST=TOTAL*(CGSTPER/100),SGST=TOTAL*(SGSTPER/100),IGST=TOTAL*(IGSTPER/100)"
con.Execute "Update Bill SET GSTTOTAL=CGST+SGST+IGST"



    Data
    Dim gval As String
    
    Unload rptGSTBillNew
    rptGSTBillNew.DiscardSavedData
    rptGSTBillNew.VerifyOnEveryPrint = True
    rptGSTBillNew.FormulaFields(2).Text = "'" & txtInvoice.Text & "'"
    rptGSTBillNew.FormulaFields(7).Text = "'" & txtParty1.Text & "'"
    rptGSTBillNew.FormulaFields(3).Text = "'" & Format(DTPicker1.Value, "dd-MM-yyyy") & "'"
    rptGSTBillNew.FormulaFields(4).Text = "'" & txtState.Text & "'"
    rptGSTBillNew.FormulaFields(5).Text = "'" & txtStateCode.Text & "'"
    rptGSTBillNew.FormulaFields(6).Text = "'" & txtVehicle.Text & "'"
    rptGSTBillNew.FormulaFields(8).Text = "'" & txtaddress.Text & "'"
    rptGSTBillNew.FormulaFields(9).Text = "'" & txtGSTIN.Text & "'"
    rptGSTBillNew.FormulaFields(11).Text = "'" & Replace(FormatCurrency(txtBTotal.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "Rs.", "") & "'"
    rptGSTBillNew.FormulaFields(10).Text = "'" & txtVehicle.Text & "'"
    rptGSTBillNew.FormulaFields(12).Text = "'" & Replace(FormatCurrency(txtSGST.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "Rs.", "") & "'"
    rptGSTBillNew.FormulaFields(13).Text = "'" & Replace(FormatCurrency(txtIGST.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "Rs.", "") & "'"
    rptGSTBillNew.FormulaFields(14).Text = "'" & Replace(FormatCurrency(txtGST.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "Rs.", "") & "'"
   ' rptGSTPortrait.FormulaFields(15).Text = "'" & Replace(FormatCurrency(txtGrandTotal.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "$", "") & "'"
    rptGSTBillNew.FormulaFields(16).Text = "'" & Replace(FormatCurrency(txtCGST.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "Rs.", "") & "'"
    
    gval = Replace(FormatCurrency(Round(Val(txtGrandTotal.Text)), 2, vbUseDefault, vbFalse, vbUseDefault), "Rs.", "")
    rptGSTBillNew.FormulaFields(15).Text = "'" & gval & "'"
    rptGSTBillNew.PrintOut
    
'
'If cboFormat.Text = "Format 1" Then
'    Unload rptBill3
'    rptBill3.DiscardSavedData
'    rptBill3.VerifyOnEveryPrint = True
'    rptBill3.FormulaFields(1).Text = "'" & txtParty1.Text & "'"
'    rptBill3.FormulaFields(2).Text = "'" & txtInvoice.Text & "'"
'    rptBill3.FormulaFields(3).Text = "'" & Format(DTPicker1.Value, "dd-MM-yyyy") & "'"
'    rptBill3.FormulaFields(4).Text = "'" & txtTIN.Text & "'"
'    rptBill3.FormulaFields(5).Text = "'" & txtPhone.Text & "'"
'    rptBill3.FormulaFields(6).Text = "'" & txtVehicle.Text & "'"
'    rptBill3.FormulaFields(7).Text = "'" & txtParty2.Text & "'"
'    rptBill3.FormulaFields(8).Text = "'" & txtParty3.Text & "'"
'    rptBill3.FormulaFields(9).Text = "'" & Replace(FormatCurrency(txtBTotal.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "$", "") & "'"
'    rptBill3.FormulaFields(10).Text = "'" & txtVat.Text & " %'"
'    rptBill3.FormulaFields(11).Text = "'" & Replace(FormatCurrency(txtVatAmount.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "$", "") & "'"
'    gval = Replace(FormatCurrency(Round(Val(txtGrandTotal.Text)), 2, vbUseDefault, vbFalse, vbUseDefault), "$", "")
'    rptBill3.FormulaFields(12).Text = "'" & gval & "'"
'    rptBill3.PrintOut
'ElseIf cboFormat.Text = "Format 2" Then
'    Unload rptBill2
'    rptBill2.DiscardSavedData
'    rptBill2.VerifyOnEveryPrint = True
'    rptBill2.FormulaFields(1).Text = "'" & txtParty1.Text & "'"
'    rptBill2.FormulaFields(2).Text = "'" & txtInvoice.Text & "'"
'    rptBill2.FormulaFields(3).Text = "'" & Format(DTPicker1.Value, "dd-MM-yyyy") & "'"
'    rptBill2.FormulaFields(4).Text = "'" & txtTIN.Text & "'"
'    rptBill2.FormulaFields(5).Text = "'" & txtPhone.Text & "'"
'    rptBill2.FormulaFields(6).Text = "'" & txtVehicle.Text & "'"
'    rptBill2.FormulaFields(7).Text = "'" & txtParty2.Text & "'"
'    rptBill2.FormulaFields(8).Text = "'" & txtParty3.Text & "'"
'    rptBill2.FormulaFields(9).Text = "'" & Replace(FormatCurrency(txtBTotal.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "$", "") & "'"
'    rptBill2.FormulaFields(10).Text = "'" & txtVat.Text & " %'"
'    rptBill2.FormulaFields(11).Text = "'" & Replace(FormatCurrency(txtVatAmount.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "$", "") & "'"
'    gval = Replace(FormatCurrency(Round(Val(txtGrandTotal.Text)), 2, vbUseDefault, vbFalse, vbUseDefault), "$", "")
'    rptBill2.FormulaFields(12).Text = "'" & gval & "'"
'    rptBill2.PrintOut
'ElseIf cboFormat.Text = "Format 3" Then
'    Unload rptBill
'    rptBill.DiscardSavedData
'    rptBill.VerifyOnEveryPrint = True
'    rptBill.FormulaFields(1).Text = "'" & txtParty1.Text & "'"
'    rptBill.FormulaFields(2).Text = "'" & txtInvoice.Text & "'"
'    rptBill.FormulaFields(3).Text = "'" & Format(DTPicker1.Value, "dd-MM-yyyy") & "'"
'    rptBill.FormulaFields(4).Text = "'" & txtTIN.Text & "'"
'    rptBill.FormulaFields(5).Text = "'" & txtPhone.Text & "'"
'    rptBill.FormulaFields(6).Text = "'" & txtVehicle.Text & "'"
'    rptBill.FormulaFields(7).Text = "'" & txtParty2.Text & "'"
'    rptBill.FormulaFields(8).Text = "'" & txtParty3.Text & "'"
'    rptBill.FormulaFields(9).Text = "'" & Replace(FormatCurrency(txtBTotal.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "$", "") & "'"
'    rptBill.FormulaFields(10).Text = "'" & txtVat.Text & " %'"
'    rptBill.FormulaFields(11).Text = "'" & Replace(FormatCurrency(txtVatAmount.Text, 2, vbUseDefault, vbFalse, vbUseDefault), "$", "") & "'"
'    gval = Replace(FormatCurrency(Round(Val(txtGrandTotal.Text)), 2, vbUseDefault, vbFalse, vbUseDefault), "$", "")
'    rptBill.FormulaFields(12).Text = "'" & gval & "'"
'    rptBill.PrintOut
'End If
End Sub
Private Sub Data()
con.Execute "UPDATE BILL set TempValue=" & Round(Val(txtGrandTotal.Text)) & ""
End Sub
Private Sub Form_Load()
cboFormat.Text = "Format 1"
DTPicker1.Value = Date
con.Execute "delete from Bill"
con.Execute "delete from Header"
If rs.State = 1 Then rs.Close
rs.Open "select * from items", con, adOpenDynamic, adLockOptimistic
cboItems.Clear
While rs.EOF = False
    cboItems.AddItem (rs(1))
    rs.MoveNext
Wend

If rs.State = 1 Then rs.Close
rs.Open "select * from Party", con, adOpenDynamic, adLockOptimistic
txtParty1.Clear
While rs.EOF = False
    txtParty1.AddItem (rs(1))
    rs.MoveNext
Wend
If rs.State = 1 Then rs.Close
rs.Open "select * from BILL", con, adOpenDynamic, adLockOptimistic
Set MSHFlexGrid1.DataSource = rs

MSHFlexGrid1.ColWidth(0) = 0
MSHFlexGrid1.ColWidth(1) = 700
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 3500
MSHFlexGrid1.ColWidth(4) = 1500
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 1800
MSHFlexGrid1.ColWidth(7) = 0
MSHFlexGrid1.ColWidth(8) = 0
MSHFlexGrid1.ColWidth(9) = 0
MSHFlexGrid1.ColWidth(10) = 0
MSHFlexGrid1.ColWidth(11) = 0
MSHFlexGrid1.ColWidth(12) = 0
MSHFlexGrid1.ColWidth(13) = 0
MSHFlexGrid1.ColWidth(14) = 0

End Sub

Private Sub Calculation()

txtCGST = Val(txtBTotal.Text) * (Val(txtCGSTPer.Text) / 100)
txtSGST = Val(txtBTotal.Text) * (Val(txtSGSTPer.Text) / 100)
txtIGST = Val(txtBTotal.Text) * (Val(txtIGSTPer.Text) / 100)

txtGST.Text = Val(txtCGST.Text) + Val(txtSGST.Text) + Val(txtIGST.Text)
txtGrandTotal.Text = Val(txtGST.Text) + Val(txtBTotal.Text)

End Sub



Private Sub MSHFlexGrid1_DblClick()
Dim r As Integer
r = MSHFlexGrid1.Row
If r > 0 Then
    cboItems.Text = MSHFlexGrid1.TextMatrix(r, 3)
    slno.Caption = MSHFlexGrid1.TextMatrix(r, 1)
    txtQTY.Text = MSHFlexGrid1.TextMatrix(r, 4)
    txtRate.Text = MSHFlexGrid1.TextMatrix(r, 5)
    txtTotal.Text = MSHFlexGrid1.TextMatrix(r, 6)
End If
End Sub

Private Sub txtBTotal_Change()
txtVat_Change
End Sub

Private Sub txtCGSTPer_Change()
Calculation
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboItems.SetFocus
End If


End Sub

Private Sub txtGSTIN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPhone.SetFocus
End If
End Sub

Private Sub txtIGSTPer_Change()
Calculation
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtGSTIN.SetFocus
End If
End Sub

Private Sub txtParty1_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from Party where Party='" & txtParty1.Text & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
txtaddress.Text = IIf(IsNull(rs(2)), "", rs(2))
'txtGSTIN.Text = IIf(IsNull(rs(3)), "", rs(3))
txtGSTIN.Text = IIf(IsNull(rs(4)), "", rs(4))
txtPhone.Text = IIf(IsNull(rs(5)), "", rs(5))
txtState.Text = IIf(IsNull(rs(6)), "", rs(6))
txtStateCode = IIf(IsNull(rs(7)), "", rs(7))
End If


End Sub

Private Sub txtParty1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtaddress.SetFocus
End If
End Sub




Private Sub txtPhone_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtVehicle.SetFocus
End If
End Sub

Private Sub txtQTY_Change()
txtTotal.Text = Val(txtQTY) * Val(txtRate)

End Sub

Private Sub txtQTY_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtRate.SetFocus
End If
End Sub

Private Sub txtRate_Change()
txtTotal.Text = Val(txtQTY) * Val(txtRate)
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAdd_Click
End If
End Sub


Private Sub txtVat_Change()
Calculation
End Sub

Private Sub txtSGSTPer_Change()
Calculation
End Sub

Private Sub txtVehicle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboItems.SetFocus
End If
End Sub
