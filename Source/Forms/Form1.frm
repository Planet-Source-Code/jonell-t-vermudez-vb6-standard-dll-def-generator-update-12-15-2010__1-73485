VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB6 Standard DLL .DEF Generator"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowseOutput 
      Caption         =   "..."
      Height          =   300
      Left            =   3855
      TabIndex        =   30
      Top             =   7785
      Width           =   300
   End
   Begin VB.ListBox lstOrdinals 
      Height          =   450
      Left            =   4845
      Sorted          =   -1  'True
      TabIndex        =   29
      Top             =   3915
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtOutput 
      Height          =   300
      Left            =   750
      TabIndex        =   17
      Top             =   7785
      Width           =   3120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   3975
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "pxModule"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwModules 
      Height          =   5385
      Left            =   165
      TabIndex        =   2
      Top             =   735
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   9499
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwExportList 
      Height          =   2460
      Left            =   2355
      TabIndex        =   4
      Top             =   3675
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   4339
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Alias"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ordinal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "No Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Private"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Data"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwPublic 
      Height          =   2685
      Left            =   2355
      TabIndex        =   3
      Top             =   750
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   4736
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Frame fraExportOptions 
      Caption         =   "Export Options"
      Height          =   1320
      Left            =   180
      TabIndex        =   23
      Top             =   6255
      Width           =   10515
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   360
         Left            =   7680
         TabIndex        =   31
         Top             =   285
         Width           =   1290
      End
      Begin VB.CommandButton cmdClearExtDLL 
         Caption         =   "X"
         Height          =   300
         Left            =   6660
         TabIndex        =   9
         Top             =   390
         Width           =   285
      End
      Begin VB.CommandButton cmdClearAlias 
         Caption         =   "X"
         Height          =   300
         Left            =   3585
         TabIndex        =   7
         Top             =   390
         Width           =   285
      End
      Begin VB.TextBox txtExtDLL 
         Height          =   300
         Left            =   4395
         TabIndex        =   8
         Top             =   390
         Width           =   2280
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   360
         Left            =   9120
         TabIndex        =   16
         Top             =   780
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   360
         Left            =   9120
         TabIndex        =   15
         Top             =   285
         Width           =   1215
      End
      Begin VB.CheckBox chkData 
         Caption         =   "&Data"
         Height          =   405
         Left            =   5865
         TabIndex        =   14
         Top             =   735
         Width           =   660
      End
      Begin VB.CheckBox ChkPrivate 
         Caption         =   "&Private"
         Height          =   495
         Left            =   4545
         TabIndex        =   13
         Top             =   690
         Width           =   810
      End
      Begin VB.TextBox txtOrdinal 
         Height          =   300
         Left            =   1485
         TabIndex        =   11
         Top             =   765
         Width           =   570
      End
      Begin VB.TextBox txtAlias 
         Height          =   300
         Left            =   1485
         TabIndex        =   6
         Top             =   390
         Width           =   2100
      End
      Begin VB.CheckBox chkForceOrdinal 
         Caption         =   "&Force Ordinal"
         Height          =   225
         Left            =   2340
         TabIndex        =   12
         Top             =   825
         Width           =   1245
      End
      Begin VB.OptionButton optNameOrd 
         Caption         =   "&Ordinal"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   10
         Top             =   795
         Width           =   930
      End
      Begin VB.OptionButton optNameOrd 
         Caption         =   "&Name"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   5
         Top             =   420
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.Label Label8 
         Caption         =   "from"
         Height          =   180
         Left            =   3990
         TabIndex        =   26
         Top             =   435
         Width           =   300
      End
      Begin VB.Label Label6 
         Caption         =   "as"
         Height          =   180
         Left            =   1230
         TabIndex        =   25
         Top             =   795
         Width           =   270
      End
      Begin VB.Label Label5 
         Caption         =   "as"
         Height          =   180
         Left            =   1230
         TabIndex        =   24
         Top             =   420
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   405
      Left            =   9480
      TabIndex        =   18
      Top             =   7770
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowsePrj 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   315
      Left            =   10380
      TabIndex        =   1
      Top             =   195
      Width           =   300
   End
   Begin VB.TextBox txtPrj 
      Height          =   315
      Left            =   1095
      TabIndex        =   0
      Top             =   195
      Width           =   9300
   End
   Begin VB.Label Label9 
      Caption         =   "Copyright (c) 2004 - 2010 by Jonell T. Vermudez. All rights reserved."
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   4410
      TabIndex        =   28
      Top             =   7860
      Width           =   4770
   End
   Begin VB.Label Label7 
      Caption         =   "Output"
      Height          =   195
      Left            =   180
      TabIndex        =   27
      Top             =   7830
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Export List"
      Height          =   210
      Left            =   2385
      TabIndex        =   22
      Top             =   3465
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Public"
      Height          =   180
      Left            =   2370
      TabIndex        =   21
      Top             =   540
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "Module"
      Height          =   195
      Left            =   195
      TabIndex        =   20
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Project File"
      Height          =   240
      Left            =   195
      TabIndex        =   19
      Top             =   255
      Width           =   915
   End
   Begin VB.Menu lv1 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu lv1mnuAdd 
         Caption         =   "Add To Export List"
      End
      Begin VB.Menu lv1mnuAddAll 
         Caption         =   "Add All To Export List"
      End
   End
   Begin VB.Menu lv2mnu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu lv2mnuRemove 
         Caption         =   "Remove Item from list"
      End
      Begin VB.Menu lv2mnuRemoveAll 
         Caption         =   "Remove All Items from list!"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim OrigColor As Long
Dim X As Integer
Dim NewProjLoaded As Boolean
Dim pubSelectedIdx As Integer


Dim txtPrjTip As New clsTooltips
Dim cmdBrowsePrjTip As New clsTooltips
Dim optNameOrd0Tip As New clsTooltips
Dim optNameOrd1Tip As New clsTooltips
Dim txtAliasTip As New clsTooltips
Dim txtOrdinalTip As New clsTooltips
Dim txtExtDLLTip As New clsTooltips
Dim cmdClearAliasTip As New clsTooltips
Dim cmdClearExtDLLTip As New clsTooltips
Dim cmdAddTip As New clsTooltips
Dim cmdRemoveTip As New clsTooltips
Dim cmdUpdateTip As New clsTooltips
Dim chkNonameTip As New clsTooltips
Dim chkPrivateTip As New clsTooltips
Dim chkDataTip As New clsTooltips
Dim txtOutputTip As New clsTooltips
Dim cmdBrowseOutputTip As New clsTooltips
Dim cmdGenerateTip As New clsTooltips


Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub chkForceOrdinal_Click()
    If chkForceOrdinal.Value = vbChecked Then
        optNameOrd(0).Enabled = False
        optNameOrd(1).Value = True
        optNameOrd(1).Enabled = True
        Label5.Enabled = False
        txtAlias.Enabled = False
        OrigColor = txtAlias.BackColor
        txtAlias.BackColor = Me.BackColor
        txtOrdinal.SetFocus
    Else
        optNameOrd(0).Enabled = True
        Label5.Enabled = True
        txtOrdinal.Enabled = True
        txtOrdinal.BackColor = txtPrj.BackColor
        txtOrdinal.SetFocus
    End If
End Sub


Private Sub cmdRemove_Click()
    If lvwExportList.ListItems.item(pubSelectedIdx).SubItems(3) <> "" Then _
        lstOrdinals.AddItem lvwExportList.ListItems.item(pubSelectedIdx).SubItems(3)
        
    lvwExportList.ListItems.Remove pubSelectedIdx
    If lvwExportList.ListItems.Count = 0 Then cmdGenerate.Enabled = False
    cmdRemove.Enabled = False
End Sub


Private Sub cmdAdd_Click()
    Dim lstItmX As ListItem
    Set lstItmX = lvwExportList.ListItems.Add
    
    lstItmX.Text = GetFuncName(lvwPublic.SelectedItem)
    cmdGenerate.Enabled = True
    
    If optNameOrd(0).Value = True And chkForceOrdinal.Value <> vbChecked Then
        If txtAlias.Text <> "" Then
            lstItmX.SubItems(1) = txtAlias.Text
            If txtExtDLL.Text <> "" Then lstItmX.SubItems(2) = txtExtDLL.Text
        End If
    End If
    
    If optNameOrd(1).Value = True Then
        If txtOrdinal.Text <> "" Then
            lstItmX.SubItems(3) = txtOrdinal.Text
        End If
        If lstOrdinals.ListCount >= 1 Then
            lstOrdinals.RemoveItem 0
        Else
            X = X + 1
        End If
    End If
    
    If chkData.Value = vbChecked Then
        lstItmX.SubItems(6) = "Yes"
    Else
        lstItmX.SubItems(6) = "No"
    End If
    
    If ChkPrivate.Value = vbChecked Then
        lstItmX.SubItems(5) = "Yes"
    Else
        lstItmX.SubItems(5) = "No"
    End If
    
    If chkForceOrdinal.Value = vbChecked Then
        lstItmX.SubItems(4) = "Yes"
    Else
        lstItmX.SubItems(4) = "No"
    End If
    
    txtOrdinal.Text = X
    txtAlias.Text = ""
    txtExtDLL.Text = ""
End Sub


Private Sub cmdBrowsePrj_Click()
    On Error Resume Next

    If txtPrj.Text <> "" Then cCOMDLG32.FileName = txtPrj.Text
    cCOMDLG32.DefaultExt = ".vbp"
    cCOMDLG32.DialogTitle = "Load a VB6 ActiveX DLL Project..."
    cCOMDLG32.InitDir = App.Path
    cCOMDLG32.Filter = "VB6 ActiveX DLL Project | *.vbp"
    cCOMDLG32.ShowOpen
    cCOMDLG32.CancelError = True
    
    If Err.Number = 32755 Then Exit Sub
    If ProjectType(cCOMDLG32.FileName) = True Then
        txtPrj.Text = cCOMDLG32.FileName
        txtPrjTip.CreateBalloon txtPrj, ProjectDesc(txtPrj.Text)
        txtOutput.Text = App.Path & "\" & prjTitle(txtPrj.Text) & ".def"
        txtOutput.Enabled = True
        txtOutput.BackColor = txtPrj.BackColor
        Label7.Enabled = True
        cmdBrowseOutput.Enabled = True
        
        Me.Caption = "VB6 Standard DLL .DEF Generator - [" & _
        cCOMDLG32.FileTitle & " : " & prjTitle(txtPrj.Text) & "]"
        lvwModules.ListItems.Clear
        lvwPublic.ListItems.Clear
        lvwExportList.ListItems.Clear
        
        Call FindModules(txtPrj.Text)
        
        NewProjLoaded = True
        txtAlias.Text = ""
        txtOrdinal.Text = ""
        cmdAdd.Enabled = False
        cmdRemove.Enabled = False
        cmdUpdate.Enabled = False
        cmdGenerate.Enabled = False
        
        fraExportOptions.Enabled = False
        optNameOrd(0).Enabled = False
        optNameOrd(1).Enabled = False
        Label5.Enabled = False
        Label6.Enabled = False
        txtAlias.Enabled = False
        txtAlias.BackColor = Me.BackColor
        ChkPrivate.Enabled = False
        chkData.Enabled = False
    End If
End Sub


Private Sub cmdGenerate_Click()
    CreateDEF
End Sub


Private Sub cmdClearAlias_Click()
    txtAlias.Text = ""
    txtExtDLL.Text = ""
    txtAlias.SetFocus
    cmdClearAlias.Enabled = False
    cmdClearExtDLL.Enabled = False
End Sub

Private Sub cmdClearExtDLL_Click()
    txtExtDLL.Text = ""
    txtExtDLL.SetFocus
    cmdClearExtDLL.Enabled = False
End Sub


Private Sub cmdBrowseOutput_Click()
    Dim oldpath As String

    On Error Resume Next
    
    cCOMDLG32.FileName = prjTitle(txtPrj.Text) & ".def"
    cCOMDLG32.DefaultExt = ".def"
    cCOMDLG32.DialogTitle = "Save to..."
    cCOMDLG32.InitDir = App.Path
    cCOMDLG32.Filter = "Standard DLL Export Definition File | *.def"
    oldpath = txtOutput.Text
    cCOMDLG32.ShowSave
    cCOMDLG32.CancelError = True
    
    If Err.Number = 32755 Then
        txtOutput.Text = oldpath
    Else
        txtOutput.Text = cCOMDLG32.FileName
    End If
End Sub


Private Sub cmdUpdate_Click()
    With lvwExportList.SelectedItem
        .ListSubItems(1).Text = txtAlias.Text
        .ListSubItems(2).Text = txtExtDLL.Text
        If optNameOrd(1).Value = True Then
            .ListSubItems(3).Text = txtOrdinal.Text
        Else
            .ListSubItems(3).Text = ""
        End If
        If chkForceOrdinal.Value = vbChecked Then
            .ListSubItems(4).Text = "Yes"
        Else
            .ListSubItems(4).Text = "No"
        End If
        If ChkPrivate.Value = vbChecked Then
            .ListSubItems(5).Text = "Yes"
        Else
            .ListSubItems(5).Text = "No"
        End If
        If chkData.Value = vbChecked Then
            .ListSubItems(6).Text = "Yes"
        Else
            .ListSubItems(6).Text = "No"
        End If
    End With
    If optNameOrd(1).Value = True Then X = X + 1
    lvwExportList.Refresh
End Sub


Private Sub Form_Load()
    InitCommonControls
    
    If optNameOrd(0).Value = True Then
        txtOrdinal.BackColor = Me.BackColor
        txtOrdinal.Enabled = False
    End If
    
    lvwModules.ColumnHeaders(1).Width = lvwModules.Width * 95 / 100
    lvwPublic.ColumnHeaders(1).Width = lvwPublic.Width * 5
    
    lvwExportList.ColumnHeaders(1).Width = lvwExportList.Width * 25 / 100
    lvwExportList.ColumnHeaders(2).Width = lvwExportList.Width * 25 / 100
    lvwExportList.ColumnHeaders(3).Width = lvwExportList.Width * 25 / 100
    lvwExportList.ColumnHeaders(4).Width = lvwExportList.Width * 11 / 100
    lvwExportList.ColumnHeaders(5).Width = lvwExportList.Width * 11 / 100
    lvwExportList.ColumnHeaders(6).Width = lvwExportList.Width * 11 / 100
    
    lvwExportList_LostFocus
    chkForceOrdinal.Enabled = False
    
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = False
    cmdGenerate.Enabled = False
    cmdClearAlias.Enabled = False
    cmdClearExtDLL.Enabled = False
    txtExtDLL.BackColor = Me.BackColor
    txtExtDLL.Enabled = False
    Label8.Enabled = False
    txtOutput.Enabled = False
    txtOutput.BackColor = Me.BackColor
    Label7.Enabled = False
    cmdBrowseOutput.Enabled = False
    
    fraExportOptions.Enabled = False
    optNameOrd(0).Enabled = False
    optNameOrd(1).Enabled = False
    Label5.Enabled = False
    Label6.Enabled = False
    txtAlias.Enabled = False
    txtAlias.BackColor = Me.BackColor
    ChkPrivate.Enabled = False
    chkData.Enabled = False
    
    SystrayOn Me, Me.Caption
    
    cmdBrowsePrjTip.CreateBalloon cmdBrowsePrj, "Browse ActiveX DLL project"
    cmdClearAliasTip.CreateBalloon cmdClearAlias, "Clear alias field"
    cmdClearExtDLLTip.CreateBalloon cmdClearExtDLL, "Clear from field"
    cmdAddTip.CreateBalloon cmdAdd, "Add this public to export list"
    cmdRemoveTip.CreateBalloon cmdRemove, "Remove this public from export list"
    cmdUpdateTip.CreateBalloon cmdUpdate, "Update selected item in the export list"
    optNameOrd0Tip.CreateBalloon optNameOrd(0), "Export as Name"
    optNameOrd1Tip.CreateBalloon optNameOrd(1), "Export Name with ordinal"
    txtAliasTip.CreateBalloon txtAlias, "Instead of name, export symbol as an alias"
    txtExtDLLTip.CreateBalloon txtExtDLL, "External DLL source of symbol (export redirection)"
    txtOrdinalTip.CreateBalloon txtOrdinal, "Export symbol with ordinal"
    chkNonameTip.CreateBalloon chkForceOrdinal, "Force export as ordinal only"
    chkPrivateTip.CreateBalloon ChkPrivate, "Exclude export from import library"
    chkDataTip.CreateBalloon chkData, "Flag this export as data"
    txtOutputTip.CreateBalloon txtOutput, "Set this path the same as to where you will" _
        & vbCrLf & "produce your final DLL"
    cmdBrowseOutputTip.CreateBalloon cmdBrowseOutput, "Browse output folder"
    cmdGenerateTip.CreateBalloon cmdGenerate, "Generate definition file!"
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    txtPrjTip.Remove
    cmdBrowsePrjTip.Remove
    optNameOrd0Tip.Remove
    optNameOrd1Tip.Remove
    cmdAddTip.Remove
    cmdRemoveTip.Remove
    cmdUpdateTip.Remove
    txtAliasTip.Remove
    cmdClearAliasTip.Remove
    txtExtDLLTip.Remove
    cmdClearExtDLLTip.Remove
    txtOrdinalTip.Remove
    chkNonameTip.Remove
    chkPrivateTip.Remove
    chkDataTip.Remove
    txtOutputTip.Remove
    cmdBrowseOutputTip.Remove
    cmdGenerateTip.Remove
End Sub


Private Sub Form_Terminate()
    SystrayOff Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SystrayOff Me
    
    Set txtPrjTip = Nothing
    Set cmdBrowsePrjTip = Nothing
    Set optNameOrd0Tip = Nothing
    Set optNameOrd1Tip = Nothing
    Set cmdAddTip = Nothing
    Set cmdRemoveTip = Nothing
    Set cmdUpdateTip = Nothing
    Set txtAliasTip = Nothing
    Set cmdClearAliasTip = Nothing
    Set txtExtDLLTip = Nothing
    Set cmdClearExtDLLTip = Nothing
    Set txtOrdinalTip = Nothing
    Set chkNonameTip = Nothing
    Set chkPrivateTip = Nothing
    Set chkDataTip = Nothing
    Set txtOutputTip = Nothing
    Set cmdBrowseOutputTip = Nothing
    Set cmdGenerateTip = Nothing
End Sub


Private Sub lvwPublic_Click()
    cmdRemove.Enabled = False
    cmdUpdate.Enabled = False
End Sub


Private Sub lvwPublic_DblClick()
    If Not lvwPublic.SelectedItem Is Nothing Then cmdAdd_Click
End Sub


Private Sub lvwPublic_ItemClick(ByVal item As MSComctlLib.ListItem)
    If X = 0 Then X = 1
    
    If lstOrdinals.ListCount >= 1 Then
        ReassignX
    Else
        txtOrdinal.Text = X
    End If
    
    If item.Selected = True Then
        cmdAdd.Enabled = True
        cmdRemove.Enabled = False
        cmdUpdate.Enabled = False
    
        fraExportOptions.Enabled = True
        optNameOrd(0).Enabled = True
        optNameOrd(1).Enabled = True
        Label5.Enabled = True
        Label6.Enabled = True
        txtAlias.Enabled = True
        txtAlias.BackColor = txtPrj.BackColor
        If optNameOrd(1).Value = True Then
            chkForceOrdinal.Enabled = True
            txtOrdinal.Enabled = True
            txtOrdinal.BackColor = txtPrj.BackColor
        End If
        ChkPrivate.Enabled = True
        chkData.Enabled = True
    Else
        cmdAdd.Enabled = False
        cmdRemove.Enabled = False
        cmdUpdate.Enabled = False
    End If
    
    chkForceOrdinal.Value = vbUnchecked
End Sub


Public Sub ReassignX()
    If lstOrdinals.ListCount >= 1 Then
        txtOrdinal.Text = Val(lstOrdinals.List(0))
    End If
End Sub


Private Sub lvwPublic_LostFocus()
    If lvwPublic.SelectedItem Is Nothing Then
        cmdAdd.Enabled = False
        
        fraExportOptions.Enabled = False
        optNameOrd(0).Enabled = False
        optNameOrd(1).Enabled = False
        Label5.Enabled = False
        Label6.Enabled = False
        txtAlias.Enabled = False
        txtAlias.BackColor = Me.BackColor
        ChkPrivate.Enabled = False
        chkData.Enabled = False
    Else
        cmdAdd.Enabled = True
    End If
End Sub


Private Sub lvwPublic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Not lvwPublic.SelectedItem Is Nothing Then PopupMenu lv1
    End If
End Sub


Private Sub lvwExportList_Click()
    cmdAdd.Enabled = False
End Sub


Private Sub lvwExportList_DblClick()
    If Not lvwExportList.SelectedItem Is Nothing Then cmdRemove_Click
End Sub


Private Sub lvwExportList_GotFocus()
    If lvwExportList.SelectedItem Is Nothing Then
        cmdRemove.Enabled = False
    Else
        cmdRemove.Enabled = True
    End If
End Sub


Private Sub lvwExportList_ItemClick(ByVal item As MSComctlLib.ListItem)
    If item.Selected = True Then
        cmdRemove.Enabled = True
        cmdUpdate.Enabled = True
        cmdAdd.Enabled = False
        pubSelectedIdx = item.Index
        GetExportItemSetting item.Index
    Else
        cmdAdd.Enabled = False
        cmdRemove.Enabled = False
    End If
End Sub


Private Sub lvwExportList_LostFocus()
    If lvwExportList.SelectedItem Is Nothing Then
        cmdRemove.Enabled = False
    Else
        cmdRemove.Enabled = True
    End If
End Sub


Private Sub lvwExportList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Not lvwExportList.SelectedItem Is Nothing Then PopupMenu lv2mnu
    End If
End Sub


Private Sub lvwModules_Click()
    fraExportOptions.Enabled = False
    optNameOrd(0).Enabled = False
    optNameOrd(1).Enabled = False
    Label5.Enabled = False
    Label6.Enabled = False
    txtAlias.Enabled = False
    txtAlias.BackColor = Me.BackColor
    txtOrdinal.Enabled = False
    txtOrdinal.BackColor = Me.BackColor
    chkForceOrdinal.Enabled = False
    ChkPrivate.Enabled = False
    chkData.Enabled = False
    
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    cmdUpdate.Enabled = False

    
    If lvwModules.ListItems.Count >= 1 Then
        lvwPublic.ListItems.Clear
        
        'since we removed the module extensions, we have to append it
        FindPublics (lvwModules.SelectedItem.Tag & lvwModules.SelectedItem & ".bas")
        Set lvwPublic.SelectedItem = Nothing
        
        If NewProjLoaded = True Then
            NewProjLoaded = False
            If lvwPublic.ListItems.Count >= 1 Then
                If X > 0 Then X = 0
            End If
        End If
    End If
End Sub


Private Sub lv1mnuAdd_Click()
    cmdAdd_Click
End Sub


Private Sub lv1mnuAddAll_Click()
    Dim i As Integer
    Dim lstItm As ListItem
    
    With lvwPublic.ListItems
        For i = 1 To lvwPublic.ListItems.Count
            Set lstItm = lvwExportList.ListItems.Add
            lstItm.Text = GetFuncName(.item(i).Text)
            lstItm.SubItems(4) = "No"
            lstItm.SubItems(5) = "No"
            lstItm.SubItems(6) = "No"
        Next i
    End With
    
    cmdGenerate.Enabled = True
    Set lstItm = Nothing
End Sub


Private Sub lv2mnuRemove_Click()
    cmdRemove_Click
End Sub


Private Sub lv2mnuRemoveAll_Click()
    lvwExportList.ListItems.Clear
    cmdGenerate.Enabled = False
    X = 0
End Sub


Private Sub optNameOrd_Click(Index As Integer)
    Select Case Index
        Case 0: txtAlias.BackColor = txtPrj.BackColor
                txtAlias.Enabled = True
                If lvwPublic.SelectedItem Is Nothing Then
                Else
                    txtAlias.SetFocus
                End If
                If txtAlias.Text <> "" Then
                    Label8.Enabled = True
                    txtExtDLL.BackColor = txtPrj.BackColor
                    txtExtDLL.Enabled = True
                Else
                    Label8.Enabled = False
                    txtExtDLL.BackColor = Me.BackColor
                    txtExtDLL.Enabled = False
                End If
                txtOrdinal.BackColor = Me.BackColor
                txtOrdinal.Enabled = False
                If txtAlias.Text <> "" Then
                    cmdClearAlias.Enabled = True
                Else
                    cmdClearAlias.Enabled = False
                End If
                If txtExtDLL.Text <> "" Then
                    cmdClearExtDLL.Enabled = True
                Else
                    cmdClearExtDLL.Enabled = False
                End If
        Case 1: txtAlias.BackColor = Me.BackColor
                txtAlias.Enabled = False
                Label8.Enabled = False
                txtExtDLL.BackColor = Me.BackColor
                txtExtDLL.Enabled = False
                txtOrdinal.BackColor = txtPrj.BackColor
                txtOrdinal.Enabled = True
                If lvwPublic.SelectedItem Is Nothing Then
                Else
                    txtOrdinal.SetFocus
                End If
                cmdClearAlias.Enabled = False
                cmdClearExtDLL.Enabled = False
    End Select
End Sub


Private Sub txtAlias_Change()
    optNameOrd(0).Value = True
    If txtAlias.Text <> "" Then
        cmdClearAlias.Enabled = True
        Label8.Enabled = True
        txtExtDLL.Enabled = True
        txtExtDLL.BackColor = txtPrj.BackColor
    Else
        cmdClearAlias.Enabled = False
        Label8.Enabled = False
        txtExtDLL.Enabled = False
        txtExtDLL.BackColor = Me.BackColor
    End If
End Sub


Private Sub txtOrdinal_Change()
    If txtOrdinal.Text <> "" Then
        chkForceOrdinal.Enabled = True
    Else
        chkForceOrdinal.Value = False
        chkForceOrdinal.Enabled = False
    End If
End Sub


Private Sub CreateDEF()
    Dim oFS As New Scripting.FileSystemObject
    Dim ts As TextStream
    
    On Error GoTo Err_
    
    Set ts = oFS.CreateTextFile(txtOutput.Text)
    
    ts.WriteLine ";**********************************************************************************************"
    ts.WriteLine "; " & prjTitle(txtPrj.Text) & ".def"
    ts.WriteLine ";"
    ts.WriteLine "; Generated with VB6 Standard DLL .DEF Generator v" _
        & App.Major & "." & App.Minor & ".2." & App.Revision & " on " & Date & " @ " & Time
    ts.WriteLine "; Copyright (c) 2004 - 2010 by Jonell T. Vermudez. All rights reserved."
    ts.WriteLine "; Pinili, Ilocos Norte, Philippines."
    ts.WriteLine "; email : pinili_boy_2003@yahoo.com"
    ts.WriteLine ";"
    ts.WriteLine "; Last updated : October 2010"
    ts.WriteLine ";**********************************************************************************************"
    ts.WriteLine ";"
    ts.WriteLine ";"
    ts.WriteLine ";"
    ts.WriteLine "; Exerpts from : http://msdn.microsoft.com/en-us/library/28d6s79h(v=VS.80).aspx"
    ts.WriteLine ";              : http://www.vb-helper.com/howto_make_standard_dll.html"
    ts.WriteLine ";              : www.jsware.net"
    ts.WriteLine ";"
    ts.WriteLine ";"
    ts.WriteLine ";"
    ts.WriteLine "; As described in an article written by Ron Petrusha in the second URL above, link your DLL"
    ts.WriteLine "; project with the linker commandline modifier (Linker-Go-Between, from the third URL above)."
    ts.WriteLine ";"
    ts.WriteLine "; In addition, this shall serve as an update to Joe Priestly's article from the third URL above"
    ts.WriteLine "; stating that he has not yet tested a .def file with multiple (.bas) modules. And this imple-"
    ts.WriteLine "; mentation proves to be working. Thus, this simple program automates .def creation for you."
    ts.WriteLine ";"
    ts.WriteLine "; And finally, this version includes an update to support DLL export redirection. For example,"
    ts.WriteLine "; from this DLL (which you are creating a .def for), you export a name (or function) contained"
    ts.WriteLine "; from another external DLL. This name may be an actual internal name or even an alias (if you"
    ts.WriteLine "; prefer) to the name exported by the other external DLL, which may also be an actual name for"
    ts.WriteLine "; the internal funtion or even another alias! (Refer to the article from the first URL above.)"
    ts.WriteLine ";"
    ts.WriteLine "; To illustrate it more clearly, take a look at the next lines:"
    ts.WriteLine ";"
    ts.WriteLine "; DLL1.DLL has these exported functions: Plus() (alias to Add()), and Subtract() (actual inter-"
    ts.WriteLine "; nal function name). DLL2.DLL (in this case is our sample project which we are creating a .def"
    ts.WriteLine "; file for), will export a function Minus(), an alias to exported Subtract from DLL1, and Plus(),"
    ts.WriteLine "; the actual exported function from DLL1, which is just another alias to the internal function"
    ts.WriteLine "; Add(), and Sum(), our local (DLL2) function, exported as is defined internally."
    ts.WriteLine ";"
    ts.WriteLine "; LIBRARY DLL2"
    ts.WriteLine "; EXPORTS"
    ts.WriteLine ";    Minus=DLL1.Subtract"
    ts.WriteLine ";    Plus"
    ts.WriteLine ";    Sum"
    ts.WriteLine ";"
    ts.WriteLine "; The following lines is the actual .def file generated for our DLL project " _
                    & prjTitle(txtPrj.Text) & "."
    ts.WriteLine ";"
    
    ts.WriteBlankLines 2
    ts.WriteLine "LIBRARY " & prjTitle(txtPrj.Text)
    
    If ProjectDesc(txtPrj.Text) <> "" Then
        ts.WriteLine "DESCRIPTION " & ProjectDesc(txtPrj.Text)
    Else
        ts.WriteLine "DESCRIPTION " & prjTitle(txtPrj.Text)
    End If
    
    ts.WriteLine "EXPORTS"
    
    Dim X As Integer
    
    'Lets get the longest name in the list to beautify our exports list in the .def file
    Dim LongestName As Integer: LongestName = 0
    
    For X = 1 To lvwExportList.ListItems.Count
        If Len(Trim(lvwExportList.ListItems.item(X).SubItems(1) _
            & "=" & lvwExportList.ListItems.item(X).SubItems(2) & "." _
            & lvwExportList.ListItems.item(X).Text)) > LongestName Then
            
            LongestName = Len(Trim(lvwExportList.ListItems.item(X).SubItems(1) _
                & "=" & lvwExportList.ListItems.item(X).SubItems(2) & "." _
                & lvwExportList.ListItems.item(X).Text)) + 2
                
        ElseIf Len(Trim(lvwExportList.ListItems.item(X).SubItems(1) _
            & "=" & lvwExportList.ListItems.item(X).SubItems(2))) > LongestName Then
            
            LongestName = Len(Trim(lvwExportList.ListItems.item(X).SubItems(1) _
                & "=" & lvwExportList.ListItems.item(X).SubItems(2)))
                
        ElseIf Len(lvwExportList.ListItems.item(X).Text) > LongestName Then
            LongestName = Len(lvwExportList.ListItems.item(X).Text) - 2
        End If
    Next X
    
    For X = 1 To lvwExportList.ListItems.Count
        If lvwExportList.ListItems.item(X).SubItems(1) <> "" Then
            If lvwExportList.ListItems.item(X).SubItems(2) <> "" Then
                ts.Write "   " & lvwExportList.ListItems.item(X).SubItems(1) _
                & "=" & lvwExportList.ListItems.item(X).SubItems(2) & "." _
                & lvwExportList.ListItems.item(X).Text
                
                If lvwExportList.ListItems.item(X).SubItems(3) = "" Then
                    ts.Write (Space(LongestName - (Len(lvwExportList.ListItems.item(X).SubItems(1) _
                    & "=" & lvwExportList.ListItems.item(X).SubItems(2) & "." _
                    & lvwExportList.ListItems.item(X).Text) + 1)))
                Else
                    ts.Write (Space(LongestName - (Len(lvwExportList.ListItems.item(X).SubItems(1) _
                    & "=" & lvwExportList.ListItems.item(X).SubItems(2) & "." _
                    & lvwExportList.ListItems.item(X).Text) + 1))) & "@" _
                    & lvwExportList.ListItems.item(X).SubItems(3) & " "
                End If
            Else
                ts.Write "   " & lvwExportList.ListItems.item(X).SubItems(1) _
                & "=" & lvwExportList.ListItems.item(X).Text & _
                (Space(LongestName - (Len(lvwExportList.ListItems.item(X).SubItems(1) _
                & "=" & lvwExportList.ListItems.item(X).Text) + 1)))
                
                If lvwExportList.ListItems.item(X).SubItems(3) <> "" Then
                    ts.Write (Space(LongestName - (Len(lvwExportList.ListItems.item(X).SubItems(1) _
                    & "=" & lvwExportList.ListItems.item(X).Text) + 1))) & "@" _
                    & lvwExportList.ListItems.item(X).SubItems(3) & " "
                End If
            End If
        Else
            ts.Write "   " & lvwExportList.ListItems.item(X).Text _
                & (Space(LongestName - Len(lvwExportList.ListItems.item(X).Text)))
            
            If lvwExportList.ListItems.item(X).SubItems(3) <> "" Then
                ts.Write "@" & lvwExportList.ListItems.item(X).SubItems(3) _
                & " "
            End If
        End If
        
        If lvwExportList.ListItems.item(X).SubItems(4) = "No" And _
            lvwExportList.ListItems.item(X).SubItems(5) = "No" And _
            lvwExportList.ListItems.item(X).SubItems(6) = "No" Then
            
            ts.WriteLine
        Else
            If lvwExportList.ListItems.item(X).SubItems(4) = "Yes" And _
                lvwExportList.ListItems.item(X).SubItems(5) = "Yes" And _
                lvwExportList.ListItems.item(X).SubItems(6) = "Yes" Then
                ts.Write "NONAME "
                ts.Write "PRIVATE "
                ts.WriteLine "DATA"
            ElseIf lvwExportList.ListItems.item(X).SubItems(4) = "Yes" And _
                lvwExportList.ListItems.item(X).SubItems(5) = "No" And _
                lvwExportList.ListItems.item(X).SubItems(6) = "No" Then ts.WriteLine "NONAME"
            ElseIf lvwExportList.ListItems.item(X).SubItems(4) = "No" And _
                lvwExportList.ListItems.item(X).SubItems(5) = "Yes" And _
                lvwExportList.ListItems.item(X).SubItems(6) = "No" Then ts.WriteLine "PRIVATE"
            ElseIf lvwExportList.ListItems.item(X).SubItems(4) = "No" And _
                lvwExportList.ListItems.item(X).SubItems(5) = "No" And _
                lvwExportList.ListItems.item(X).SubItems(6) = "Yes" Then ts.WriteLine "DATA"
            ElseIf lvwExportList.ListItems.item(X).SubItems(4) = "Yes" And _
                lvwExportList.ListItems.item(X).SubItems(5) = "No" And _
                lvwExportList.ListItems.item(X).SubItems(6) = "Yes" Then
                ts.Write "NONAME "
                ts.WriteLine "DATA"
            ElseIf lvwExportList.ListItems.item(X).SubItems(4) = "Yes" And _
                lvwExportList.ListItems.item(X).SubItems(5) = "Yes" And _
                lvwExportList.ListItems.item(X).SubItems(6) = "No" Then
                ts.Write "NONAME "
                ts.WriteLine "PRIVATE"
            Else
                ts.Write "PRIVATE "
                ts.WriteLine "DATA"
            End If
        End If
    Next X
    
    ts.WriteBlankLines 2
    ts.Close
    Set ts = Nothing
    
    PopupBalloon Me, prjTitle(txtPrj.Text) & _
    ".def has been created successfully!" & _
    vbCrLf & "At least " & lvwExportList.ListItems.Count _
    & " public symbols were exported.", _
    "VB6 Standard DLL .DEF Generator", tipIconInfo
    
    Exit Sub
Err_:
    PopupBalloon Me, "There was an error creating the definition file " _
    & vbCrLf & prjTitle(txtPrj.Text) & ".def", _
    "VB6 Standard DLL .DEF Generator", tipIconError

    ts.Close
    Set ts = Nothing
End Sub


Private Sub GetExportItemSetting(idx As Integer)
    fraExportOptions.Enabled = True
    
    txtAlias.Text = lvwExportList.ListItems.item(idx).SubItems(1)
    txtExtDLL.Text = lvwExportList.ListItems.item(idx).SubItems(2)
    If lvwExportList.ListItems.item(idx).SubItems(3) <> "" Then
        txtOrdinal.Text = lvwExportList.ListItems.item(idx).SubItems(3)
    Else
        txtOrdinal = X
    End If
    
    If lvwExportList.ListItems.item(idx).SubItems(1) <> "" Then
        optNameOrd(0).Value = True
        optNameOrd(0).Enabled = True
        optNameOrd(1).Enabled = True
        If txtAlias.Text <> "" Then
            txtAlias.Enabled = True
            txtAlias.BackColor = txtPrj.BackColor
            cmdClearAlias.Enabled = True
            If txtExtDLL.Text <> "" Then
                txtExtDLL.Enabled = True
                txtExtDLL.BackColor = txtPrj.BackColor
                cmdClearExtDLL.Enabled = True
            End If
        End If
    ElseIf lvwExportList.ListItems.item(idx).SubItems(3) <> "" Then
            optNameOrd(1).Value = True
            optNameOrd(1).Enabled = True
            txtOrdinal.Enabled = True
            txtOrdinal.BackColor = txtPrj.BackColor
    Else
            optNameOrd(0).Value = True
            optNameOrd(0).Enabled = True
            optNameOrd(1).Enabled = True
    End If
    
    If lvwExportList.ListItems.item(idx).SubItems(4) = "No" Then
        chkForceOrdinal.Value = vbUnchecked
        optNameOrd(0).Enabled = True
    Else
        chkForceOrdinal.Value = vbChecked
        optNameOrd(0).Enabled = False
    End If
    If lvwExportList.ListItems.item(idx).SubItems(5) = "No" Then
        ChkPrivate.Value = vbUnchecked
    Else
        ChkPrivate.Value = vbChecked
    End If
    If lvwExportList.ListItems.item(idx).SubItems(6) = "No" Then
        chkData.Value = vbUnchecked
    Else
        chkData.Value = vbChecked
    End If
End Sub


Private Sub txtExtDLL_Change()
    If txtExtDLL.Text <> "" Then
        cmdClearExtDLL.Enabled = True
    Else
        cmdClearExtDLL.Enabled = False
    End If
End Sub

