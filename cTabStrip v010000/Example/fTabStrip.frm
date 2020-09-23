VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fTabStrip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test cTabStrip Container Management Class Wrapper"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fTabStrip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      Caption         =   "Testing Options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   5670
      TabIndex        =   4
      Top             =   4830
      Width           =   3795
      Begin VB.Frame fraOptContainer 
         Caption         =   "TabStrip Container:"
         Height          =   960
         Left            =   210
         TabIndex        =   5
         ToolTipText     =   "Note: When the container control is altered, the attached controls are automatically resized."
         Top             =   315
         Width           =   3165
         Begin VB.OptionButton optOptions 
            Appearance      =   0  'Flat
            Caption         =   "Frame Container Control"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   630
            TabIndex        =   7
            ToolTipText     =   "Set the TabStrip container to a Frame Control"
            Top             =   630
            Width           =   2325
         End
         Begin VB.OptionButton optOptions 
            Appearance      =   0  'Flat
            Caption         =   "Dialog Box (This Form)"
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   630
            TabIndex        =   6
            ToolTipText     =   "Set the TabStrip container to the Form"
            Top             =   315
            Value           =   -1  'True
            Width           =   2325
         End
      End
      Begin VB.ComboBox cboOptions 
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Set how all attached control are automatically aligned"
         Top             =   1680
         Width           =   2220
      End
      Begin VB.CheckBox chkOption 
         Appearance      =   0  'Flat
         Caption         =   "AutoFit container/control within TabStrip"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   210
         TabIndex        =   8
         ToolTipText     =   "Automatically fit all attached controls to fit the TabStrip's client area"
         Top             =   1365
         Width           =   3270
      End
      Begin VB.PictureBox picDialog 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3990
         Left            =   3150
         Picture         =   "fTabStrip.frx":27A2
         ScaleHeight     =   3990
         ScaleWidth      =   6060
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1785
         Visible         =   0   'False
         Width           =   6060
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alignment:"
         Height          =   315
         Left            =   210
         TabIndex        =   9
         ToolTipText     =   "Set how all attached control are automatically aligned"
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Frame fraDialog 
      Caption         =   "Frame Container"
      Height          =   3900
      Left            =   105
      TabIndex        =   0
      Top             =   4725
      Visible         =   0   'False
      Width           =   5475
      Begin VB.ListBox lstDialog 
         Height          =   3375
         Left            =   2520
         TabIndex        =   3
         Top             =   210
         Width           =   3480
      End
   End
   Begin MSComctlLib.TabStrip tabDialog 
      Height          =   4530
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   7990
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tab with PictureBox"
            Object.ToolTipText     =   "Demonstrates using a PictureBox as the attached container"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tab with ListBox"
            Object.ToolTipText     =   "Demonstrates using a ListBox as the attached control"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tab with Nothing"
            Object.ToolTipText     =   "Demonstrates having a Tab with nothing associated with it"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Configure"
            Key             =   "TestOptTab"
            Object.ToolTipText     =   "Demonstrates using a Frame as the attached container"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "fTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'
' Form Name:    fTabStrip
' Author:       Graeme Grant        (a.k.a. Slider)
' Date:         10/08/2002
' Version:      01.00.00
' Description:  Demonstrate how easy the cTrapStrip Management Class is to use.
' Edit History: 01.00.00 10/08/2002 Initial Release
'
' Notes:        To view all GUI features, set form Height = 10000,
'               Width = 10000; to hide, set form Height = 5055,
'               Width = 6615
'
'===========================================================================

Option Explicit

'===========================================================================
' Private: Demonstration declarations
'
Private Type tPosDEF
    Width As Single
    Height As Single
End Type

Private mtPos(2) As tPosDEF                 '## Used to store original control size

Private moTabStrip As cTabStrip             '## Declare reference to class wrapper

'===========================================================================
' Form Events
'
Private Sub Form_Load()

    Set moTabStrip = New cTabStrip          '## Initialise class wrapper

    With moTabStrip
        .HookCtrl tabDialog                 '## Advises wrapper which TabStrip control to manage
        .AutoFit = False                    '## Don't automatically resize attached controls to
                                            '   fit the client area of the TabStrip control
        .Attach 1, picDialog                '## Attach PicDialog PictureBox to Tab 1
        .Attach 2, lstDialog                '## Attach lstDialog ListBox to Tab 2
        .Attach "TestOptTab", fraOptions    '## Attach fraOptions Frame to Tab 4 (Key = TestOptTab)
        '
        '## No more is needed! The class wrapper 'cTabStrip' will take care of the rest!
        '
    End With

    pInitData                               '## Initialise sample data for ListBox control &
                                            '   data for configuration options
End Sub

'===========================================================================
' Control Events used for demonstration purposes only.
'
Private Sub cboOptions_Click()
    '
    '## Change Alignment to user-specified
    '
    moTabStrip.Alignment = cboOptions.ItemData(cboOptions.ListIndex)

End Sub

Private Sub chkOption_Click()

    moTabStrip.AutoFit = (chkOption.Value = vbChecked)      '## Toggle AutoFit State
    lblOptions.Visible = (chkOption.Value = vbUnchecked)    '## Toggle visible State
    cboOptions.Visible = (chkOption.Value = vbUnchecked)    '## Toggle visible State
    '
    '## AutoFit off, therefore restore alignment & sizes
    '
    If (chkOption = vbUnchecked) Then pResetSize

End Sub

Private Sub optOptions_Click(Index As Integer)

    Select Case Index
        Case 0
            '
            '## Set TabStrip container (and attached controls) to the Form
            '
            With fraDialog
                '
                '## Hide the Frame Control
                '
                .Visible = False
                '
                '## Reset TabStrip co-ordinates
                '
                tabDialog.Move .Left, .Top, .Width, .Height
                '
                '## Set TabStrip Container (Note: As the TabStrip control does not have a
                '                                 Resize event, the control is resized before
                '                                 the container is changed)
                '
                Set moTabStrip.TabStipContainer = Me

            End With

        Case 1
            '
            '## Set the TabStrip container (and attached controls) to Frame Control
            '
            With fraDialog
                '
                '## Set Size and positioning of the frame control
                '
                .Move tabDialog.Left, tabDialog.Top, tabDialog.Width, tabDialog.Height
                '
                '## Fit TabStrip inside Frame control
                '
                tabDialog.Move .Left + 100, .Top + 100, .Width - 400, .Height - 400
                '
                '## Set TabStrip Container (Note: As the TabStrip control does not have a
                '                                 Resize event, the control is resized before
                '                                 the container is changed)
                '
                Set moTabStrip.TabStipContainer = fraDialog
                '
                '## Show Frame Control
                '
                .Visible = True

            End With

    End Select

End Sub

'===========================================================================
' Private: Demonstration methods
'
Private Sub pInitData()

    Dim iLoop As Long

    With lstDialog
        For iLoop = 1 To 30
            .AddItem "List Item " + CStr(iLoop)
            .ItemData(.NewIndex) = iLoop
        Next
        mtPos(0).Width = .Width
        mtPos(0).Height = .Height
    End With
    With picDialog
        mtPos(1).Width = .Width
        mtPos(1).Height = .Height
    End With
    With fraOptions
        mtPos(2).Width = .Width
        mtPos(2).Height = .Height
    End With

    With cboOptions
        .AddItem "Align TopLeft"
        .ItemData(.NewIndex) = [Align TopLeft]          '## ListIndex = 0
        .AddItem "Align TopHCenter"
        .ItemData(.NewIndex) = [Align TopHCenter]       '## ListIndex = 1
        .AddItem "Align TopRight"
        .ItemData(.NewIndex) = [Align TopRight]         '## ListIndex = 2
        .AddItem "Align LeftVCenter"
        .ItemData(.NewIndex) = [Align LeftVCenter]      '## ListIndex = 3
        .AddItem "Align VCenterHCenter"
        .ItemData(.NewIndex) = [Align VCenterHCenter]   '## ListIndex = 4
        .AddItem "Align RightVCenter"
        .ItemData(.NewIndex) = [Align RightVCenter]     '## ListIndex = 5
        .AddItem "Align BottomLeft"
        .ItemData(.NewIndex) = [Align BottomLeft]       '## ListIndex = 6
        .AddItem "Align BottomHCenter"
        .ItemData(.NewIndex) = [Align BottomHCenter]    '## ListIndex = 7
        .AddItem "Align BottomRight"
        .ItemData(.NewIndex) = [Align BottomRight]      '## ListIndex = 8
        .ListIndex = 4
    End With

End Sub

Private Sub pResetSize()

    With lstDialog
        .Width = mtPos(0).Width
        .Height = mtPos(0).Height
    End With
    With picDialog
        .Width = mtPos(1).Width
        .Height = mtPos(1).Height
    End With
    With fraOptions
        .Width = mtPos(2).Width
        .Height = mtPos(2).Height
    End With

    cboOptions_Click    '## Force the Alignment property to be refreshed causing the attached
                        '   controls to be re-positioned
End Sub
