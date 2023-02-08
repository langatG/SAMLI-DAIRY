VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmBankSetup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Bank Setup"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   Icon            =   "frmBankSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwSummary 
      Height          =   2775
      Left            =   120
      TabIndex        =   21
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Bank Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bank Accno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Gl AccNO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Bank Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Gl Account"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Frame fraBank 
      Height          =   2895
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   7680
      Begin VB.ComboBox cboAccType 
         Height          =   315
         ItemData        =   "frmBankSetup.frx":030A
         Left            =   2280
         List            =   "frmBankSetup.frx":0317
         TabIndex        =   28
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtBankAccNO 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   975
         Width           =   2055
      End
      Begin VB.CommandButton cmdAcctsSearch 
         Height          =   300
         Left            =   3090
         Picture         =   "frmBankSetup.frx":0334
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2167
         Width           =   330
      End
      Begin VB.TextBox txtAccNames 
         Height          =   315
         Left            =   3420
         TabIndex        =   24
         Top             =   2160
         Width           =   3225
      End
      Begin VB.ComboBox cboAccno 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2160
         Width           =   1200
      End
      Begin VB.TextBox txtBankCode 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   375
         Width           =   1575
      End
      Begin VB.TextBox txtBankName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   375
         Width           =   5775
      End
      Begin VB.TextBox txtTelephone 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1695
         Width           =   2055
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtBranchName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4440
         TabIndex        =   11
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Associated Gl Account"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2190
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Bank Acc. Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Bank Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   105
         TabIndex        =   20
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Account Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   18
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Phone No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2355
         TabIndex        =   16
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Branch Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   495
      Left            =   4680
      Picture         =   "frmBankSetup.frx":0436
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cancel Process"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      Height          =   495
      Left            =   4200
      Picture         =   "frmBankSetup.frx":0538
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Save Record"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdDelete 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Picture         =   "frmBankSetup.frx":063A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete Record"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdEdit 
      Height          =   495
      Left            =   3240
      Picture         =   "frmBankSetup.frx":073C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Edit Record"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1095
      TabIndex        =   2
      ToolTipText     =   "Move to the Next"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1575
      Picture         =   "frmBankSetup.frx":083E
      TabIndex        =   3
      ToolTipText     =   "Move to Last record"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "Move to the Previous record"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Move to the Last record"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   495
      Left            =   2760
      Picture         =   "frmBankSetup.frx":0940
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Add New record"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   22
      Top             =   3240
      Width           =   1230
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankSetup.frx":0E72
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankSetup.frx":0F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankSetup.frx":1096
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBankSetup.frx":11A8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBankSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim disablemodifying As Boolean
Private Sub cboAccno_Change()
    Dim ACCNO As String
    ACCNO = cboAccno.Text
    sql = "select GLACCNAME from glsetup where accno='" & ACCNO & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
        txtAccNames.Text = rs(0)
    End If
End Sub
Private Sub cboAccno_Click()
    cboAccno_Change
End Sub

Private Sub cmdAcctsSearch_Click()
    frmAcctsSearch.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            cboAccno.List(0) = SearchValue
            cboAccno.Text = cboAccno.List(0)
            SearchValue = ""
            Continue = False
        End If
    End If
End Sub
Private Sub cmdAdd_Click()
    On Error GoTo errFix
        NewRecord = True
        lvwSummary.Visible = False
        UnlockControls Me
        cmdupdate.Enabled = True
        Load
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdAdd_Click
    txtBankCode.SetFocus
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdcancel_Click()
    On Error GoTo errFix
    cmdAdd.Enabled = True
    lvwSummary.Visible = True
    cmdupdate.Enabled = False
    cmdEdit.Enabled = True
    Exit Sub
errFix:
        MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdcancel_Click
End If
Exit Sub
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdclose_Click()
On Error GoTo errFix
Unload Me
action = ""
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub



Private Sub cmdClose_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdclose_Click
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmddelete_Click()
On Error GoTo errFix
If lvwSummary.ListItems.Count > 0 Then
    sel = lvwSummary.SelectedItem
End If
If lvwSummary.Visible Then

    Set rst = oSaccoMaster.GetRecordset("select * from banks where bankcode= '" & sel & "'")
    If rst.RecordCount > 0 Then
        If MsgBox("Are you sure you want to delete " & rst!BankName & "" & rst!branchname & " ? ", vbYesNo, "Bank deletion") = vbYes Then
            Set Rst5 = oSaccoMaster.GetRecordset("select * from banks where bankcode= '" & sel & "'")
            If Not Rst5.EOF Then
                Rst5.Delete
                Rst5.Update
            End If
             
             
'            load_records
         End If
     End If
 
Else
    Set rst = oSaccoMaster.GetRecordset("select * from banks where bankcode= '" & txtBankCode.Text & "'")
    If rst.RecordCount > 0 Then
    If MsgBox("Are you sure you want to delete " & rst!BankName & "" & rst!branchname & "" & " ? ", vbYesNo, "bank deletion") = vbYes Then
        Set rst = oSaccoMaster.GetRecordset("select * from banks where bankcode= '" & txtBankCode.Text & "'")
        rst.Delete
        rst.Update
        Set Rst5 = oSaccoMaster.GetRecordset("select * from banks")
        Rst5.MoveFirst
        txtBankCode.Text = Rst5!BankCode & ""
'        load_records
     End If
 End If
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdDelete_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmddelete_Click
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdedit_Click()
On Error GoTo errFix
    NewRecord = False
    lvwSummary.Visible = False
    cmdupdate.Enabled = True
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdEdit_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdedit_Click
    txtBankName.SetFocus
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdFirst_Click()
On Error GoTo errFix
'Toolbar1.Buttons("bSearch").Enabled = True
'Toolbar1.Buttons("bView").Enabled = True
'Toolbar1.Buttons("bPrint").Enabled = True
'Toolbar1.Buttons("bImport").Enabled = True
Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
'        If TypeOf ctrl Is CheckBox Then
'            If Not ctrl = chkPreviewReport Then
'                ctrl.Enabled = False
'            End If
'        End If
    Next ctrl
Set rst1 = oSaccoMaster.GetRecordset("select bankcode from banks order by bankcode")
With rst1
    If .RecordCount > 0 Then
        .MoveFirst
        txtBankCode.Text = rst1!BankCode & ""
'        load_records
        cmdFirst.Enabled = False
        cmdPrevious.Enabled = False
        cmdNext.Enabled = True
        cmdLast.Enabled = True
    End If
End With

rst1.Close
action = ""
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdFirst_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdFirst_Click
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdHelp_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    'cmdHelp_Click
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdLast_Click()
On Error GoTo errFix
'Toolbar1.Buttons("bSearch").Enabled = True
'Toolbar1.Buttons("bView").Enabled = True
'Toolbar1.Buttons("bPrint").Enabled = True
'Toolbar1.Buttons("bImport").Enabled = True
Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
'        If TypeOf ctrl Is CheckBox Then
'            If Not ctrl = chkPreviewReport Then
'                ctrl.Enabled = False
'            End If
'        End If
    Next ctrl
Set rst1 = oSaccoMaster.GetRecordset("select bankcode from banks order by bankcode")
With rst1
    If .RecordCount > 0 Then
        .MoveLast
        txtBankCode.Text = !BankCode & ""
        cmdFirst.Enabled = True
        cmdPrevious.Enabled = True
        cmdNext.Enabled = False
        cmdLast.Enabled = False
'        load_records
    End If
End With

If action = "editingRecords" Or action = "addingRecords" Then
    If disablemodifying = False Then
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    End If
    cmdCancel.Enabled = False
    cmdDelete.Enabled = True
End If
rst1.Close
action = ""
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdLast_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdLast_Click
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdNext_Click()
On Error GoTo errFix
'Toolbar1.Buttons("bSearch").Enabled = True
'Toolbar1.Buttons("bView").Enabled = True
'Toolbar1.Buttons("bPrint").Enabled = True
'Toolbar1.Buttons("bImport").Enabled = True
Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
'        If TypeOf ctrl Is CheckBox Then
'            If Not ctrl = chkPreviewReport Then
'                ctrl.Enabled = False
'            End If
'        End If
    Next ctrl
Set rst1 = oSaccoMaster.GetRecordset("select * from banks order by bankcode")
If cmdupdate.Enabled = True Then
    If Not rst1.EOF Then
        rst1.Bookmark = MyBookMark
        txtBankCode.Text = rst1!BankCode & ""
    End If
End If
With rst1
    If .RecordCount > 0 Then
        .Find "bankcode= '" & txtBankCode.Text & "'"
        If Not .EOF Then
            .MoveNext
            If .EOF Then
                .MoveLast
                cmdFirst.Enabled = True
                cmdPrevious.Enabled = True
                cmdNext.Enabled = False
                cmdLast.Enabled = False
            Else
                cmdFirst.Enabled = True
                cmdPrevious.Enabled = True
                cmdNext.Enabled = True
                cmdLast.Enabled = True
            End If
            txtBankCode.Text = !BankCode & ""
'            load_records
        End If
    End If
End With
rst1.Close
action = ""
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub
Private Sub cmdNext_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdNext_Click
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub
Private Sub cmdPrevious_Click()
On Error GoTo errFix
Dim ctrl As Control
For Each ctrl In Me
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        End If
        
        If TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        End If
        
    Next ctrl
Set rst1 = oSaccoMaster.GetRecordset("select bankcode from banks order by bankcode")

If cmdupdate.Enabled = True Then
    rst1.Bookmark = MyBookMark
    txtBankCode.Text = rst1!BankCode & ""
End If
With rst1
    If Not rst1.EOF Then
        .MovePrevious
        .Find ("bankcode= '" & txtBankCode.Text & "'")
        If Not .EOF Then
            .MovePrevious
            If .BOF Then
                .MoveFirst
                cmdFirst.Enabled = False
                cmdPrevious.Enabled = False
                cmdNext.Enabled = True
                cmdLast.Enabled = True
            Else
                cmdFirst.Enabled = True
                cmdPrevious.Enabled = True
                cmdNext.Enabled = True
                cmdLast.Enabled = True
            End If
          txtBankCode.Text = !BankCode & ""
'          load_records
        End If
    End If
End With
rst1.Close
action = ""
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub cmdPrevious_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdPrevious_Click
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub
Private Sub cmdupdate_Click()
    On Error GoTo errFix
    Set rst = oSaccoMaster.GetRecordset("Select * from d_Banks order by bankcode")
        With rst
            If NewRecord = True Then
                .AddNew
                !BankCode = txtBankCode.Text & ""
            End If
            !BankName = txtBankName.Text & ""
            !BankAccNo = txtBankAccNO.Text
            !ACCNO = cboAccno.Text
            !branchname = txtBranchName.Text & ""
            !Telephone = txtTelephone.Text & ""
            !Address = txtAddress.Text & ""
            !auditid = User
            !accType = cboAccType
            .Update
            LoadBanks
        End With
    MsgBox "Record Saved Successfully"
    cmdAdd.Enabled = True
    cmdupdate.Enabled = False
    Form_Load
    'Load
    fraBank.Enabled = False
    Exit Sub
errFix:
        MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub
Private Sub Form_Load()
    On Error GoTo errFix
    PositionForm Me
    LoadBanks
    'Load Gl's
    sql = "Select bankcode from d_banks order by bankcode asc"
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
        cboAccno.AddItem (rst(0))
        rst.MoveNext
    Wend
    
            Exit Sub
errFix:
        MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub
Function EncryptPassword()
On Error GoTo errFix
    Dim Pwd As Variant
    Dim Temp As String, PwdChr As Long
    Dim EncryptKey As Long
    Pwd = valToEncrOrDecr
    EncryptKey = Int(Sqr(Len(Pwd) * 81)) + 23
    
    For PwdChr = 1 To Len(Pwd)
        Temp = Temp + Chr(Asc(mid(Pwd, PwdChr, 1)) Xor EncryptKey)
    Next PwdChr
    
    EncryptPass = Temp
    Exit Function
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Function
Private Function fieldFound(FieldName As String) As Boolean
On Error GoTo errFix
fieldFound = False
For I = 0 To rstRecordsImported.Fields.Count - 1
     If UCase(rstRecordsImported.Fields(I).Name) = UCase(FieldName) Then
        fieldFound = True
        Exit Function
    End If
Next I
Exit Function
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errFix
action = "" 'TO cancel edit or add mode
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub
Private Sub lvwSummary_DblClick()
    With lvwSummary
        If .ListItems.Count = 0 Then
            Exit Sub
        End If
        txtBankCode.Text = .SelectedItem.Text
        .Visible = False
    End With
End Sub

Private Sub txtAddress_Change()
On Error GoTo errFix
txtAddress.Text = UCase(txtAddress.Text)
txtAddress.SelStart = Len(txtAddress.Text)
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtAddress.Text) > 29 Then
      Beep
      MsgBox "Can't enter more than 30 characters", vbExclamation
      KeyAscii = 8
End If
If KeyAscii = 13 Then
    txtTelephone.SetFocus
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtBankCode_Change()
    On Error GoTo errFix
    If txtBankCode.Text = "" Then
        Exit Sub
    End If
    txtBankCode.Text = UCase(txtBankCode.Text)
    txtBankCode.SelStart = Len(txtBankCode.Text)
    
    Set Rst5 = oSaccoMaster.GetRecordset("select * from banks where bankcode= '" & txtBankCode.Text & "'")
    With Rst5
        If Not .EOF Then
            txtBankCode.Text = !BankCode
            txtBankName.Text = !BankName
            txtBankAccNO.Text = !BankAccNo
            cboAccno.Text = !ACCNO
            txtBranchName.Text = !branchname
            txtTelephone.Text = !Telephone
            txtAddress.Text = !Address
            User = !auditid
            cboAccType = !accType
        Else
            txtBankName.Text = ""
            txtBankAccNO.Text = ""
            txtBranchName.Text = ""
            txtTelephone.Text = ""
            txtAddress.Text = ""
            cboAccType = ""
            'txtBankName.SetFocus
        End If
    End With
    
    Exit Sub

errFix:
        MsgBox err.number, vbOKOnly, "Bank Setup"
End Sub



Private Sub txtBankCode_Click()
    txtBankCode_Change
End Sub



Private Sub txtBankCode_LostFocus()
On Error GoTo errFix
    If action = "addingRecords" Then
        Set Rst5 = oSaccoMaster.GetRecordset("select bankcode from banks where bankcode= '" & txtBankCode.Text & "'")
        If Not Rst5.EOF Then
            MsgBox "Bank with same code already exists", vbInformation, "Bank Code"
            txtBankCode.Text = ""
            txtBankCode.SetFocus
        
        End If
    End If
    Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtBankName_Change()
On Error GoTo errFix
txtBankName.Text = UCase(txtBankName.Text)
txtBankName.SelStart = Len(txtBankName.Text)
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtBankName_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtBankName.Text) > 49 Then
      Beep
      MsgBox "Can't enter more than 50 characters", vbExclamation
      KeyAscii = 8
End If
If KeyAscii = 13 Then
    txtBranchName.SetFocus
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtBranchName_Change()
On Error GoTo errFix
txtBranchName.Text = UCase(txtBranchName.Text)
txtBranchName.SelStart = Len(txtBranchName.Text)
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtBranchName_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If Len(txtBranchName.Text) > 49 Then
      Beep
      MsgBox "Can't enter more than 50 characters", vbExclamation
      KeyAscii = 8
End If
If KeyAscii = 13 Then
    txtAddress.SetFocus
End If
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtTelephone_Change()
On Error GoTo errFix
txtTelephone.Text = UCase(txtTelephone.Text)
txtTelephone.SelStart = Len(txtTelephone.Text)
Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub

Private Sub txtTelephone_KeyPress(KeyAscii As Integer)
On Error GoTo errFix
If KeyAscii = 13 Then
    cmdupdate.SetFocus
End If
Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc("-")
    Case Asc("+")
    Case Asc(")")
    Case Asc(" ")
    Case Asc(",")
    Case Asc("(")
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
  End Select
If Len(txtTelephone.Text) > 14 Then
      Beep
      MsgBox "Can't enter more than 15 characters", vbExclamation
      KeyAscii = 8
End If

  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc("-")
    Case Asc("+")
    Case Asc(")")
    Case Asc(" ")
    Case Asc(",")
    Case Asc("(")
    Case Is = 8
    Case Else
        Beep
        KeyAscii = 0
  End Select
  Exit Sub
errFix:
    MsgBox err.description, vbOKOnly, "Bank Setup"
End Sub
Sub LoadBanks()
    lvwSummary.Visible = True
    lvwSummary.ListItems.Clear
    Set rst = oSaccoMaster.GetRecordset("select BankCode,BankName,BranchName " _
    & " from d_banks")
    With rst
        While Not .EOF
            Set li = lvwSummary.ListItems.Add(, , !BankCode)
            'li.ListSubItems.Add , , !BankAccNo
            'li.ListSubItems.Add , , !ACCNO & ""
            li.ListSubItems.Add , , !BankName
            li.ListSubItems.Add , , !branchname
            'li.ListSubItems.Add , , !GlAccName
            .MoveNext
        Wend
    End With
End Sub
 Private Sub Load()
      PositionForm Me
     cboAccno.Clear
    sql = "Select accno from glsetup order by accno asc"
    Set rst = oSaccoMaster.GetRecordset(sql)
    While Not rst.EOF
        cboAccno.AddItem (rst(0))
        rst.MoveNext
    Wend
            txtBankCode.Text = ""
           txtBankName.Text = ""
            txtBankAccNO.Text = ""
            txtBranchName.Text = ""
            txtTelephone.Text = ""
            txtAddress.Text = ""
            cboAccType = ""
Exit Sub
errFix:
        MsgBox err.description, vbOKOnly, "Bank Setup"
 End Sub
