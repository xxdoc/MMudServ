VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCommands 
   Caption         =   "Commands Listing"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6300
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   6300
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   360
      Left            =   1125
      TabIndex        =   2
      Top             =   7485
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   15
      TabIndex        =   1
      Top             =   7515
      Width           =   1065
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   13150
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    ListView1.Width = Me.Width - 7
    ListView1.Height = Me.Height - 7
    
    Me.ListView1.ColumnHeaders.Add , , "Command"
    Me.ListView1.ColumnHeaders.Add , , "Code"
    Me.ListView1.ColumnHeaders.Add , , "Level"
    Me.ListView1.ColumnHeaders.Add , , "Pos"
    Me.ListView1.ColumnHeaders.Add , , "Log"
    
    Dim i As Integer
    
    For i = 1 To UBound(HUB.Commands)
        
        ListView1.ListItems.Add , HUB.Commands(i).Name, HUB.Commands(i).Name
        ListView1.ListItems.Item(HUB.Commands(i).Name).SubItems(1) = HUB.Commands(i).Code
        
    Next i
    
End Sub

Private Sub Form_Resize()
    ListView1.Width = Me.Width - 7
    ListView1.Height = Me.Height - 7
End Sub
