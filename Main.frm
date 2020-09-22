VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   Caption         =   "Ant colony simulation"
   ClientHeight    =   10905
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10815
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   727
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   721
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frmGraph 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Analysis"
      Height          =   3615
      Left            =   8280
      TabIndex        =   16
      Top             =   6120
      Width           =   2535
      Begin VB.PictureBox GraphCargo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1525
         Left            =   120
         Picture         =   "Main.frx":49E2
         ScaleHeight     =   102
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   26
         ToolTipText     =   "Cargo transit efficiency"
         Top             =   2040
         Width           =   2285
      End
      Begin VB.PictureBox GraphCount 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1525
         Left            =   120
         Picture         =   "Main.frx":4D71
         ScaleHeight     =   102
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   17
         ToolTipText     =   "Colony size as percentage of starting value"
         Top             =   360
         Width           =   2285
      End
   End
   Begin VB.PictureBox ZoomExtends 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4560
      Picture         =   "Main.frx":5134
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      ToolTipText     =   "Recenter terrarium"
      Top             =   4800
      Width           =   300
   End
   Begin VB.Frame frmControls 
      Caption         =   "Controls and settings"
      Height          =   6060
      Left            =   8280
      TabIndex        =   4
      Top             =   0
      Width           =   2535
      Begin VB.CheckBox RenderOptions 
         Caption         =   "Render ants"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Top             =   5040
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox RenderOptions 
         Caption         =   "Render values"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   4800
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox RenderOptions 
         Caption         =   "Render food"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   4560
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox RenderOptions 
         Caption         =   "Render scent"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   4320
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox RenderOptions 
         Caption         =   "Render grid"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   4080
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "Help..."
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "Info ..."
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   5400
         Width           =   2295
      End
      Begin VB.CheckBox chkRedraw 
         Caption         =   "High frequency redraw"
         Height          =   285
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Toggle master iteration refresh"
         Top             =   3800
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CommandButton cmdStopSimulation 
         Caption         =   "Stop simulation"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   3420
         Width           =   2295
      End
      Begin VB.CommandButton cmdPauseSimulation 
         Caption         =   "Pause simulation"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   3150
         Width           =   2295
      End
      Begin VB.TextBox txtIterationRatio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   28
         ToolTipText     =   "Density of food-distribution..."
         Top             =   2250
         Width           =   975
      End
      Begin VB.TextBox txtBirth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   25
         ToolTipText     =   "Density of food-distribution..."
         Top             =   1980
         Width           =   975
      End
      Begin VB.TextBox txtMaxCargo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         ToolTipText     =   "Density of food-distribution..."
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton cmdStartSimulation 
         Caption         =   "Start simulation"
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtFoodDensity 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         ToolTipText     =   "Density of food-distribution..."
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtAntAge 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         ToolTipText     =   "Lifespan of ants in iterations..."
         Top             =   1170
         Width           =   975
      End
      Begin VB.TextBox txtColonySize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         ToolTipText     =   "Antcount at colony start..."
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox txtExtend 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         ToolTipText     =   "Number of terrafields in both directions..."
         Top             =   630
         Width           =   975
      End
      Begin VB.TextBox txtGridSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         ToolTipText     =   "The terrarium grid size in pixels..."
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdRecreate 
         Caption         =   "Recreate terrarium"
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   2530
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Steps per loop"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   2250
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Birth threshold"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maximum cargo"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Food density"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Average ant age"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1170
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Colony size"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Terra extend"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Terra grid size"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.PictureBox hScroll 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   3
      Top             =   4800
      Width           =   4575
   End
   Begin VB.PictureBox vScroll 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   4560
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   0
      Width           =   300
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   10650
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   3969
            MinWidth        =   3969
            Text            =   "Number of ants:"
            TextSave        =   "Number of ants:"
            Object.ToolTipText     =   "Total amount of live ants"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   3969
            MinWidth        =   3969
            Text            =   "Food carriers:"
            TextSave        =   "Food carriers:"
            Object.ToolTipText     =   "Total amount of ants with food"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   3969
            MinWidth        =   3969
            Text            =   "Cargo in transit:"
            TextSave        =   "Cargo in transit:"
            Object.ToolTipText     =   "Total amount of food carried by ants"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   3969
            MinWidth        =   3969
            Text            =   "Cycle time:"
            TextSave        =   "Cycle time:"
            Object.ToolTipText     =   "Average time per iterations"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   3969
            MinWidth        =   3969
            Text            =   "Master iterations:"
            TextSave        =   "Master iterations:"
            Object.ToolTipText     =   "Number of master iterations"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Terra 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4800
      Left            =   0
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   0
      Width           =   4560
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnStopSolution As Boolean
Public blnPauseSolution As Boolean

Private Sub chkRedraw_Click()
    Select Case chkRedraw.Value
    Case 0
        RedrawFrequency = 50
    Case Else
        RedrawFrequency = 1
    End Select
End Sub

Private Sub cmdHelp_Click()
    Unload Banner
    Help.Show 1
End Sub

Private Sub cmdInfo_Click()
    Banner.tmrBanner.Interval = 1000
    Banner.Show
End Sub

Private Sub Form_Load()
    Randomize
    Terra.BackColor = RGB(230, 230, 230)
    With Settings
        .AntAge = 1000
        .AntCount = 0
        .BioMatter = 0.1
        .Birth = 5
        .ColFood = 0
        .ColonySize = 50
        .CycleTime = 1
        .GridSize = 30
        .HomePoint = CreatePoint
        .IterationRatio = 10
        .MaxCargo = 3
        .RenderAnts = True
        .RenderFood = True
        .RenderGrid = True
        .RenderLabels = True
        .RenderScent = True
        .TerraExtend = 30
        .Transit = 0
    End With
    RedrawFrequency = 1
    Call Interface.Settings2Interface
    Call Colony.InitiateTables
    Call Colony.InitiateTerrarium
    Main.Show
    Call Interface.CenterViewport
End Sub

Private Sub Form_Resize()
    Call Interface.DistributeForm
    Call Interface.DrawVerticalSlider
    Call Interface.DrawHorizontalSlider
    Terra.Cls
    Call Interface.RefreshViewport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub cmdRecreate_Click()
    Call Colony.InitiateTerrarium
    Terra.Cls
    Call Interface.RefreshViewport
    Call Interface.DrawVerticalSlider
    Call Interface.DrawHorizontalSlider
End Sub

Private Sub cmdStartSimulation_Click()
    If Settings.AntCount = 0 Then Call Colony.InitiateAnts(Val(txtColonySize))
    cmdStartSimulation.Enabled = False
    cmdPauseSimulation.Enabled = True
    cmdStopSimulation.Enabled = True
    cmdPauseSimulation.SetFocus
    blnStopSolution = False
    blnPauseSolution = False
    Call Colony.SimulateTerrarium
End Sub

Private Sub cmdPauseSimulation_Click()
    blnPauseSolution = True
    cmdStartSimulation.Caption = "Resume simulation"
    cmdStartSimulation.Enabled = True
    cmdPauseSimulation.Enabled = False
    cmdStartSimulation.SetFocus
End Sub

Public Sub cmdStopSimulation_Click()
    blnStopSolution = True
    blnPauseSolution = False
    cmdStopSimulation.Enabled = False
    cmdPauseSimulation.Enabled = False
    cmdStartSimulation.Enabled = True
    cmdStartSimulation.Caption = "Start simulation"
    cmdStartSimulation.SetFocus
    Call Colony.NukeAnts
    Call Colony.InitiateTables
    Terra.Cls
    Call Interface.RefreshViewport
End Sub

Private Sub RenderOptions_Click(Index As Integer)
    Select Case Index
    Case 0
        Settings.RenderGrid = False
        If RenderOptions(Index).Value = 1 Then Settings.RenderGrid = True
    Case 1
        Settings.RenderScent = False
        If RenderOptions(Index).Value = 1 Then Settings.RenderScent = True
    Case 2
        Settings.RenderFood = False
        If RenderOptions(Index).Value = 1 Then Settings.RenderFood = True
    Case 3
        Settings.RenderLabels = False
        If RenderOptions(Index).Value = 1 Then Settings.RenderLabels = True
    Case 4
        Settings.RenderAnts = False
        If RenderOptions(Index).Value = 1 Then Settings.RenderAnts = True
    End Select
End Sub

Private Sub Terra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 And Button <> 2 Then Exit Sub
    Dim xStep As Double, yStep As Double
    Dim Xt As Double, Yt As Double
    Dim Xo As Double, Yo As Double
    Dim i As Double
    
    Xt = (X - (Terra.ScaleWidth / 2)) / 20
    Yt = (Y - (Terra.ScaleHeight / 2)) / 20
    Xo = Offset(0)
    Yo = Offset(1)
    
    For i = 20 To 1 Step -1
        Xo = Xo + Xt
        Yo = Yo + Yt
        Offset(0) = Xo
        Offset(1) = Yo
        Call Interface.DrawVerticalSlider
        Call Interface.DrawHorizontalSlider
        Terra.Cls
        Call Interface.DrawTerraGrid
        Call Interface.DrawFood
        DoEvents
    Next
    Call Interface.RefreshViewport
End Sub

Private Sub txtAntAge_LostFocus()
    Call Interface.Interface2Settings
End Sub

Private Sub txtBirth_LostFocus()
    Call Interface.Interface2Settings
End Sub

Private Sub txtColonySize_LostFocus()
    Call Interface.Interface2Settings
End Sub

Private Sub txtExtend_LostFocus()
    Call Interface.Interface2Settings
End Sub

Private Sub txtFoodDensity_LostFocus()
    Call Interface.Interface2Settings
End Sub

Private Sub txtGridSize_LostFocus()
    Call Interface.Interface2Settings
End Sub

Private Sub txtIterationRatio_LostFocus()
    Call Interface.Interface2Settings
End Sub

Private Sub txtMaxCargo_LostFocus()
    Call Interface.Interface2Settings
End Sub

Private Sub ZoomExtends_Click()
    Call Interface.CenterViewport
End Sub
