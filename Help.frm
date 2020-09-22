VERSION 5.00
Begin VB.Form Help 
   AutoRedraw      =   -1  'True
   Caption         =   "Info..."
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "close"
      Height          =   255
      Left            =   5990
      TabIndex        =   1
      Top             =   5490
      Width           =   600
   End
   Begin VB.TextBox txtHelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strHelp As String
    
    strHelp = "This application aims to simulate the behaviour " & _
              "of an ant colony. This behaviour is goverened by " & _
              "the enviroment (food-amount and location) and " & _
              "settings such as average ant age, ant strength and " & _
              "the speed at which food can regenerate." & vbNewLine & vbNewLine & _
              "Into this changing environment we allow individual " & _
              "agents to operate freely. They both respond to and " & _
              "influence the environment according to as few rules " & _
              "as possible." & vbNewLine & vbNewLine & _
              "Of course in the end the goal of every ant is to " & _
              "find as much food as it can carry and then run home with " & _
              "it. But individual ants do not have sufficient data " & _
              "to plan this on their own. Ants have to communicate with " & _
              "each other using scent-trails. The goal of the experiment " & _
              "is to gain insight into the relation and significance " & _
              "of different settings and behavioural rules such as:" & vbNewLine & vbNewLine
    strHelp = strHelp & _
              "  » The extend of the available territory" & vbNewLine & _
              "  » The resolution of scent" & vbNewLine & _
              "  » The initial size of the colony" & vbNewLine & _
              "  » The life expectancy of ants" & vbNewLine & _
              "  » The initial density of food-resources" & vbNewLine & _
              "  » The amount of food that one ant can carry" & vbNewLine & _
              "  » The amount of food that is required for a new ant" & vbNewLine & _
              "  » The hierarchy of different scent-trails" & vbNewLine & vbNewLine
    strHelp = strHelp & _
              "Some of these properties are hard-coded into the algorithms " & _
              "while others are accessible through the interface. " & _
              "During simulation the user is provided with several kinds " & _
              "of feedback. The statusbar lists several colony-properties " & _
              "such as size, cargo in transit and colony age. In addition to " & _
              "these numeric values there are also two graphs on the screen:" & vbNewLine & vbNewLine
    strHelp = strHelp & _
              "  » The size of the colony in relation to the initial size" & vbNewLine & _
              "  » The efficiency of the colony" & vbNewLine & _
              "    (Cargo in transit, in relation to the maximum)" & vbNewLine & vbNewLine & _
              "The graphs shift to the left every master iteration. They provide " & _
              "clear visual feedback on the progress of the colony."
              
    txtHelp.Text = strHelp

End Sub

Private Sub Form_Resize()
    txtHelp.Left = 0
    txtHelp.Top = 0
    txtHelp.Width = Me.ScaleWidth
    txtHelp.Height = Me.ScaleHeight
    cmdClose.Left = Me.ScaleWidth - 865
    cmdClose.Top = Me.ScaleHeight - 285
End Sub
