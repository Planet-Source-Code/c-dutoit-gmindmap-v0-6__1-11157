VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVueArbre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vue en arbre"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   3405
   Begin MSComctlLib.TreeView TreeView 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   11033
      _Version        =   393217
      Indentation     =   441
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
End
Attribute VB_Name = "frmVueArbre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmVueArbre
'Added by C.Dutoit 24 August 2000
'dutoitc@hotmail.com - http://www.home.ch/~spaw4758
Option Explicit

Sub RefreshTree()
    'Clear existing nodes
    TreeView.Nodes.Clear
    
    'Add the root node
    TreeView.Nodes.Add , , "N" & Str$(0), Arbre(0).Legende
    
    'Add the children
    RefreshTreeRec 0
End Sub 'RefreshTree


Sub RefreshTreeRec(Root)
    Dim i, n
    
    TreeView.Nodes("N" & Str$(Root)).Expanded = Arbre(Root).Expanded
    
    'Add the children of the root
    For i = 0 To Arbre(Root).NbSuivants - 1
        'Find the current node
        n = Arbre(Root).Suivants(i)
        
        'Add the current node
        TreeView.Nodes.Add "N" & Str$(Root), tvwChild, "N" & Str$(n), Arbre(n).Legende
        'TreeView.Nodes("N" & Str$(n)).Expanded = Arbre(n).Expanded
        
        'Add all of the children nodes
        RefreshTreeRec (n)
    Next i
End Sub 'RefreshTreeRec


'Refresh the current tree
Private Sub Form_Load()
    RefreshTree
End Sub 'Form_Load


Private Sub Form_Resize()
    'Resize the tree control
    TreeView.Left = 0
    TreeView.Top = 0
    TreeView.Width = Width
    TreeView.Height = Height
End Sub


'Save node state
Private Sub TreeView_Collapse(ByVal Node As MSComctlLib.Node)
    Dim n
    n = Right$(Node.Key, Len(Node.Key) - 1)
    Arbre(Val(n)).Expanded = Node.Expanded
End Sub 'TreeView_Collapse


'Save node state
Private Sub TreeView_Expand(ByVal Node As MSComctlLib.Node)
    Dim n
    n = Right$(Node.Key, Len(Node.Key) - 1)
    Arbre(Val(n)).Expanded = Node.Expanded
End Sub 'TreeView_Expand


'Define the node as root
Private Sub TreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim n
    n = Right$(Node.Key, Len(Node.Key) - 1)
    Racine = Val(n)
    DessinerAllMindMap
End Sub 'TreeView_NodeClick
