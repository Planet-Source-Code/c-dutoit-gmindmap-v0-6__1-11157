Attribute VB_Name = "modMap"
'modMap : Gestion de l'affichage du mindmap + structure de donnée
'Par C.Dutoit, 1er Août 2000 (dutoitc@hotmail.com)
'http://www.home.ch/~spaw4758
Option Explicit

'max 10 fils !


'Un Noeud
Type TNoeud
    Legende As String       'Légende du noeud
    URL As String           'URL
    x As Long
    y As Long               'Position centrale
    NbSuivants As Byte      'Nombre de fils
    Suivants() As Long      'Liste des fils
    Expanded As Boolean     'vrai si ce noeud est ouvert (+) dans l'arbre
    PositionForcee As Boolean 'true si la positions (x,y) est forcée par l'utilisateur
End Type 'TNoeud


Global Arbre() As TNoeud         'L'arbre du mindmap
Global NoeudSelectionne As Long  'Noeud sélectionné
Global Racine As Long          'racine de l'arbre



'Dessiner un noeud
Private Sub DessinerNoeud(x As Long, y As Long, index As Long)
    Dim txtW As Long
    Dim txtH As Long
    Dim w As Long           'Largeur
    Dim h As Long           'Hauteur
    
    'Calculer la hauteur et la largeur
    txtW = frmMap.TextWidth(Arbre(index).Legende)
    txtH = frmMap.TextHeight(Arbre(index).Legende)
    w = txtW * 0.5 + frmMap.TextWidth("OO")
    h = txtH * 0.5 + frmMap.TextHeight("O") / 2
    
    'Dessiner le centre
    frmMap.FillColor = RGB(255, 255, 255)
    frmMap.FillStyle = 0 'solide
    frmMap.DrawWidth = 2
    frmMap.Circle (GetRealPosXfrmMap(x), GetRealPosYfrmMap(y)), w, , , , h / w
    frmMap.DrawWidth = 1
    
    'Sélectionné ? => tracer un cadre traitillé autour de l'ellipse
    If index = NoeudSelectionne Then
        frmMap.ForeColor = 0
        frmMap.DrawStyle = 2
        frmMap.FillStyle = 1 'transparent
        frmMap.Line (GetRealPosXfrmMap(x) - txtW / 2 - 2, GetRealPosYfrmMap(y) - txtH / 2 - 2)-(frmMap.ScaleWidth / 2 + x + txtW / 2 + 2, frmMap.ScaleHeight / 2 + y + txtH / 2 + 2), , B
        frmMap.DrawStyle = 0
    End If
    
    'Afficher le label
    frmMap.CurrentX = GetRealPosXfrmMap(x) - txtW / 2
    frmMap.CurrentY = GetRealPosYfrmMap(y) - txtH / 2
    frmMap.ForeColor = 0 'Couleur du cadre
    'frmMap.BackColor = RGB(255, 255, 200)
    'frmMap.FillColor = RGB(0, 255, 0)
    frmMap.Print Arbre(index).Legende & vbCrLf & Arbre(index).URL
End Sub 'DessinerNoeud



Private Sub DessinerNoeudEtFils(NoeudDepart As Long, Etape)
 Dim NewX As Long, NewY As Long, AngleTexte As Single, text As String, hcar As Byte, i, x As Long, y As Long
    x = Arbre(NoeudDepart).x '***/ 1000 * (frmMap.ScaleWidth / 2) + frmMap.ScaleWidth / 2
    y = Arbre(NoeudDepart).y '***/ 1000 * (frmMap.ScaleHeight / 2) + frmMap.ScaleHeight / 2
    
    
    'Dessiner les suivants
    If Arbre(NoeudDepart).NbSuivants > 0 Then
        'Afficher chaque suivant
        For i = 0 To Arbre(NoeudDepart).NbSuivants - 1
            'Coordonnées
            NewX = Arbre(Arbre(NoeudDepart).Suivants(i)).x '***/ 1000 * (frmMap.ScaleWidth / 2) + frmMap.ScaleWidth / 2
            NewY = Arbre(Arbre(NoeudDepart).Suivants(i)).y '***/ 1000 * (frmMap.ScaleHeight / 2) + frmMap.ScaleHeight / 2
            
                        
            'Forcer la position ?
            If Racine = 0 And Arbre(Arbre(NoeudDepart).Suivants(i)).PositionForcee Then
                'Afficher un rond => pos forcée ?
                If frmMDI.mnuNoeudsAffNPosForcee.Checked = True Then
                    frmMap.FillStyle = 0 'solide
                    frmMap.FillColor = RGB(0, 0, 255)
                    frmMap.Circle (frmMap.ScaleWidth / 2 + NewX, frmMap.ScaleHeight / 2 + NewY), 5, RGB(0, 0, 255)
                    frmMap.FillStyle = 1 'transparent
                End If
            End If
            
            'ReCalculer l'angle du texte
            If x = NewX Then
                AngleTexte = 90
            Else
                AngleTexte = -Atn((GetRealPosYfrmMap(NewY) - GetRealPosYfrmMap(y)) / (GetRealPosXfrmMap(NewX) - GetRealPosXfrmMap(x))) * 180 / 3.1415926535
            End If
            
            'Tracer une ligne
            frmMap.ForeColor = RGB(Etape * 64 Mod 256, Etape * 128 Mod 256, Etape * 32 Mod 256)
            frmMap.DrawWidth = ((HauteurArbre(0) - Etape) / HauteurArbre(0) * 3) ^ 2 + 1
            frmMap.Line (GetRealPosXfrmMap(x), GetRealPosYfrmMap(y))-(GetRealPosXfrmMap(NewX), GetRealPosYfrmMap(NewY))
            frmMap.DrawWidth = 1
           
            'Trouver la position (x,y) de début du texte
            hcar = ((HauteurArbre(0) - Etape) * 3 / HauteurArbre(0)) ^ 2 + 8
            text = Arbre(Arbre(NoeudDepart).Suivants(i)).Legende
            Dim XTexte As Long, YTexte As Long, Angle As Single
            If Etape = 1 Then
                XTexte = (3 * NewX + 2 * x) / 5 '- Cos(AngleTexte) * Dist
                YTexte = (3 * NewY + 2 * y) / 5  '- Sin(AngleTexte) * Dist
            Else
                XTexte = (NewX + x) / 2  '- Cos(AngleTexte) * Dist
                YTexte = (NewY + y) / 2  '- Sin(AngleTexte) * Dist
            End If
            
            'If NewX - x < 0 Then Angle = AngleTexte + 180
            
            XTexte = XTexte + GetVecHfrmMap(frmMap.TextHeight("O") / 4 * Cos((90 - Angle) * 3.1415926535 / 180) * 2)
            YTexte = YTexte + GetVecLfrmMap(frmMap.TextHeight("O") / 4 * Sin((90 - Angle) * 3.1415926535 / 180) * 2)
            PrintRotfrmMap GetRealPosXfrmMap(XTexte), GetRealPosYfrmMap(YTexte), AngleTexte, text, hcar
                                      
            DessinerNoeudEtFils Arbre(NoeudDepart).Suivants(i), Etape + 1
        Next i
    End If
    
    'Dessiner la racine
    If Etape = 1 Then DessinerNoeud x, y, NoeudDepart
    
    'Noeud sélectionné => tracer un cercle
    If NoeudSelectionne = NoeudDepart And NoeudSelectionne <> Racine Then
        frmMap.FillColor = RGB(255, 255, 255)
        frmMap.ForeColor = RGB(255, 0, 0)
        frmMap.FillStyle = 0 'solide
        frmMap.Circle (GetRealPosXfrmMap(x), GetRealPosYfrmMap(y)), 5, RGB(255, 0, 0)
    End If
End Sub 'DessinerNoeudEtFils



'Dessiner tous le mindmap
Sub DessinerAllMindMap()
    frmMap.Cls
    CalculerCoordonnees
    DessinerNoeudEtFils Racine, 1
    frmVueArbre.RefreshTree
End Sub 'DessinerAllMindMap


'Calculer les coordonnées de tous les noeuds par récursion
Private Sub CalculerCoordonneesRec(NoeudDepart As Long, AngleDeb, AngleFin, x, y, Etape)
    If NoeudDepart > 0 Then
        Arbre(NoeudDepart).x = x '***(x - frmMap.ScaleWidth / 2) / (frmMap.ScaleWidth / 2) * 1000
        Arbre(NoeudDepart).y = y '***(y - frmMap.ScaleHeight / 2) / (frmMap.ScaleHeight / 2) * 1000
    End If
    

    'Dessiner les suivants
    If Arbre(NoeudDepart).NbSuivants > 0 Then
        'Normaliser les angles
        Dim IncAngle
        If AngleDeb < 0 Then AngleDeb = AngleDeb + 360
        If AngleFin < AngleDeb Then AngleFin = AngleFin + 360
    
        'Calculer l'incrément
        If Arbre(NoeudDepart).NbSuivants = 1 Then
            IncAngle = 0
            AngleDeb = (AngleDeb + AngleFin) / 2
        Else
            If AngleDeb Mod 360 = AngleFin Mod 360 Then
                IncAngle = (AngleFin - AngleDeb) / (Arbre(NoeudDepart).NbSuivants)
            Else
                IncAngle = (AngleFin - AngleDeb) / (Arbre(NoeudDepart).NbSuivants - 1)
            End If
        End If
    
        Dim i
        Dim NewAngleDeb
        Dim NewAngleFin
        Dim Delta
        Dim NewX, NewY
        Dim Dist, Angle As Single '***modifié
        Dim Xp, Yp

    
        'Afficher chaque suivant
        For i = 0 To Arbre(NoeudDepart).NbSuivants - 1
            'Calculer les angles limites
            Delta = (90 - Etape * 9)
            NewAngleDeb = IncAngle * i + AngleDeb - Delta / 2
            NewAngleFin = IncAngle * i + AngleDeb + Delta / 2
        
            'Calculer l'angle (en radian)
            Angle = (IncAngle * i + AngleDeb) / 180 * 3.1415926535
            
            'Calculer la pos. finale
            Dim texte As String
            Dim AngleTexte As Long
            Dim hcar As Byte
            AngleTexte = Angle * 180 / 3.1415926535 '-Atn((NewY - Y) / (NewX - X)) * 180 / 3.1415926535
            If AngleTexte Mod 360 > 90 And AngleTexte Mod 360 < 270 Then AngleTexte = AngleTexte Mod 360 - 180
            texte = Arbre(Arbre(NoeudDepart).Suivants(i)).Legende
            hcar = ((HauteurArbre(0) - Etape) * 3 / HauteurArbre(0)) ^ 2 + 8
            
            'Forcer la position ?
            If Racine = 0 And Arbre(Arbre(NoeudDepart).Suivants(i)).PositionForcee Then
                NewX = Arbre(Arbre(NoeudDepart).Suivants(i)).x
                NewY = Arbre(Arbre(NoeudDepart).Suivants(i)).y
                
                'ReCalculer l'angle du texte
                AngleTexte = -Atn((NewY - y) / (NewX - x + 0.000001)) * 180 / 3.1415926535
            Else
                NewX = x + GetVecLfrmMap(LongueurTexteRot(texte & "OO", hcar) * Cos(Angle))  ' * Dist '((HauteurArbre(0) - Etape + 1) / HauteurArbre(0) * Dist + 10)
                NewY = y - GetVecHfrmMap(LongueurTexteRot(texte & "OO", hcar) * Sin(Angle))  '* Dist '((HauteurArbre(0) - Etape + 1) / HauteurArbre(0) * Dist + 10)
                
                If NoeudDepart = 0 Then 'fils de racine ? => agrandir
                    NewX = NewX + GetVecLfrmMap(LongueurTexteRot(Arbre(0).Legende & "OO", hcar) / 2 * Cos(Angle))
                    NewY = NewY - GetVecHfrmMap(LongueurTexteRot(Arbre(0).Legende, hcar) / 2 * Sin(Angle))
                End If
            End If
                           
                   
            CalculerCoordonneesRec Arbre(NoeudDepart).Suivants(i), NewAngleDeb, NewAngleFin, NewX, NewY, Etape + 1
        Next i
    End If
End Sub 'CalculerCoordonneesRec



'Calculer les coordonnées de tous les noeuds (sauf les noeuds fixés)
Sub CalculerCoordonnees()
    CalculerCoordonneesRec Racine, 0, 360, 0, 0, 1
End Sub 'CalculerCoordonnees


'Retourner la hauteur de l'arbre depuis la racine
Function HauteurArbre(Racine) As Long
    Dim h As Long       'Hauteur de l'arbre
    h = 0               'Hauteur à 0
    
    'Hauteur des fils
    Dim i, HTemp
    For i = 0 To Arbre(Racine).NbSuivants - 1
        HTemp = HauteurArbre(Arbre(Racine).Suivants(i))
        If HTemp > h Then h = HTemp
    Next i
    
    'Retourner la hauteur + 1 pour cet étage
    HauteurArbre = h + 1
End Function 'HauteurArbre


'Retourne la position réel en pixel d'une position vectorielle
Function GetRealPosXfrmMap(x As Long) As Long
    GetRealPosXfrmMap = x / (Largeur / 2) * frmMap.ScaleWidth / 2 + frmMap.ScaleWidth / 2
End Function 'GetRealPosXfrmMap

'Retourne la position réel en pixel d'une position vectorielle
Function GetRealPosYfrmMap(y As Long) As Long
    GetRealPosYfrmMap = y / (Hauteur / 2) * frmMap.ScaleHeight / 2 + frmMap.ScaleHeight / 2
End Function 'GetRealPosYfrmMap


'Retourne la position réel en pixel d'une position vectorielle
Function GetRealPosXprinter(x As Long) As Long
    GetRealPosXprinter = x / (Largeur / 2) * Printer.ScaleWidth / 2 + Printer.ScaleWidth / 2
End Function 'GetRealPosXPrinter

'Retourne la position réel en pixel d'une position vectorielle
Function GetRealPosYPrinter(y As Long) As Long
    GetRealPosYPrinter = y / (Hauteur / 2) * Printer.ScaleHeight / 2 + Printer.ScaleHeight / 2
End Function 'GetRealPosYPrinter



'Retourne la position vectorielle en pixel d'une position réelle
Function GetVecPosXfrmMap(x As Long) As Long
    GetVecPosXfrmMap = (x - frmMap.ScaleWidth / 2) * (Largeur / 2) / (frmMap.ScaleWidth / 2)
End Function 'GetVecPosXfrmMap

'Retourne la position vectorielle en pixel d'une position réelle
Function GetVecPosYfrmMap(y As Long) As Long
    GetVecPosYfrmMap = (y - frmMap.ScaleHeight / 2) * (Hauteur / 2) / (frmMap.ScaleHeight / 2)
End Function 'GetVecPosYfrmMap


'Retourner une largeur en position vectorielle
Function GetVecLfrmMap(L As Long) As Long
    GetVecLfrmMap = L * Largeur / frmMap.ScaleWidth
End Function 'GetVecLfrmMap


'Retourner une hauteur en position vectorielle
Function GetVecHfrmMap(h As Long) As Long
    GetVecHfrmMap = h * Hauteur / frmMap.ScaleHeight
End Function 'GetVecHfrmMap


'Retourner une largeur en position vectorielle
Function GetVecLPrinter(L As Long) As Long
    GetVecLPrinter = L * Largeur / Printer.ScaleWidth
End Function 'GetVecLPrinter


'Retourner une hauteur en position vectorielle
Function GetVecHPrinter(h As Long) As Long
    GetVecHPrinter = h * Hauteur / Printer.ScaleHeight
End Function 'GetVecHPrinter

