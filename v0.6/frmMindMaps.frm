VERSION 5.00
Begin VB.Form frmMap 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   Caption         =   "MindMap"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   692
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmMap
Option Explicit

Dim NoeudAccroche As Long 'contient le N° du noeud accroché pour déplacement. -1 si aucun noeud accroché


Private Sub Form_Click()
   DessinerAllMindMap
End Sub 'Form_Click


'Edition d'un noeud
Private Sub Form_DblClick()
    'Editer le noeud et redessiner le mindmap
    frmProperties.EditerNoeud (NoeudSelectionne)
    DessinerAllMindMap
    
    'Définir le titre de la fenêtre principale + ...
    If Not MyApp.Modifie Then
        MyApp.Modifie = True
        SetAppCaption
    End If
End Sub 'Form_DblClick


'Supprimer le noeud sélectionné
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete:
            SupprimerNoeud (NoeudSelectionne)
            DessinerAllMindMap
        
            'Définir le titre de la fenêtre principale + ...
            If Not MyApp.Modifie Then
                MyApp.Modifie = True
                SetAppCaption
            End If
        Case vbKeyInsert:
            frmMDI.mnuNoeudInsererFils_Click
        Case vbKeyRight:
            If NoeudSelectionne > -1 Then SelectionnerLeNoeudADroite Arbre(NoeudSelectionne).x, Arbre(NoeudSelectionne).y
        Case vbKeyLeft:
            If NoeudSelectionne > -1 Then SelectionnerLeNoeudAGauche Arbre(NoeudSelectionne).x, Arbre(NoeudSelectionne).y
        Case vbKeyUp:
            If NoeudSelectionne > -1 Then SelectionnerLeNoeudEnHaut Arbre(NoeudSelectionne).x, Arbre(NoeudSelectionne).y
        Case vbKeyDown:
            If NoeudSelectionne > -1 Then SelectionnerLeNoeudEnBas Arbre(NoeudSelectionne).x, Arbre(NoeudSelectionne).y
            
    End Select
End Sub 'Form_KeyDown


'Initialiser le mindmap
Private Sub Form_Load()
   'ScaleLeft = -1000
   'ScaleWidth = 2000
   'ScaleHeight = 2000
   'ScaleTop = -1000
   'DoEvents
   NouveauFichier
   ResizeWindows
   'DessinerAllMindMap
   NoeudAccroche = -1
End Sub 'Form_Load


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   'Retrouver une positions vectorielle
   x = GetVecPosXfrmMap(Int(x))
   y = GetVecPosYfrmMap(Int(y))
   
   NoeudSelectionne = NoeudLePlusProcheXY(Int(x), Int(y))
   If Shift And vbShiftMask Then NoeudAccroche = NoeudSelectionne
   
   'Afficher le menu popup 1
   If Button = vbRightButton And NoeudSelectionne > -1 Then
       If Arbre(NoeudSelectionne).PositionForcee Then
           frmMDI.mnuPopFrmMapForcerPos.Caption = LoadResString(OfsLanguage + 720)
       Else
           frmMDI.mnuPopFrmMapForcerPos.Caption = LoadResString(OfsLanguage + 721)
       End If
       frmMap.PopupMenu frmMDI.mnuPopFrmMap
       'Définir le titre de la fenêtre principale + ...
       If Not MyApp.Modifie Then
           MyApp.Modifie = True
           SetAppCaption
       End If
   End If
End Sub 'Form_MouseDown


'Déplacer le noeud accroché
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   'Retrouver une positions vectorielle
   x = GetVecPosXfrmMap(Int(x))
   y = GetVecPosYfrmMap(Int(y))

    'minimum 5 pixels de déplacement pour accrocher le noeud
    If NoeudAccroche > -1 Then
        'recentrer par rapport à 0, puis limites -1000 => +1000
        Arbre(NoeudAccroche).x = x
        Arbre(NoeudAccroche).y = y
        Arbre(NoeudAccroche).PositionForcee = True
        MyApp.Modifie = True
        NoeudSelectionne = NoeudAccroche
    End If
End Sub 'Form_MouseMove


Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If NoeudAccroche > -1 Then
        NoeudAccroche = -1  'Désaccrocher le noeud
        DessinerAllMindMap  'mettre à jour l'affichage
    End If
End Sub 'Form_MouseUp


Private Sub Form_Paint()
    'Mettre à jour l'affichage
    DessinerAllMindMap
End Sub 'Form_Paint


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Cancel = 1
End Sub 'Form_QueryUnload


Private Sub Form_Resize()
    'Mettre à jour l'affichage
    DessinerAllMindMap
End Sub 'Form_Resize
