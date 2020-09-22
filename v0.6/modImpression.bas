Attribute VB_Name = "modImpression"
'modImpression : gestion de l'impression
'Par C.Dutoit, 3 Août 2000 (dutoitc@hotmail.com)
'http://www.home.ch/~spaw4758
Option Explicit

Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long



'Préparer la feuille pour l'impression
Private Sub DessinerCartouche()
    Dim BordGauche, BordHaut
    Dim HauteurLigne 'Hauteur d'une ligne
    Dim Intervale    'Hauteur entre 2 lignes
    
    BordGauche = Printer.ScaleWidth - Printer.TextWidth("OOOOOOOOOOOOOOOOOOOOOOOOOOOOOO") - 1 '(30 caractères de large)
    HauteurLigne = Printer.TextHeight("O")
    Intervale = HauteurLigne / 2
    BordHaut = Printer.ScaleHeight - (HauteurLigne + 2 * Intervale) * 3 - 1 'place pour 3 lignes
    
    'Cartouche
    Printer.Line (BordGauche, BordHaut)- _
                 (Printer.ScaleWidth - 1, Printer.ScaleHeight - 1), , B
                
    'Trait horizontal entre le titre et "G-Mindmap..."
    Printer.Line (BordGauche, BordHaut + HauteurLigne + 2 * Intervale)- _
                 (Printer.ScaleWidth, BordHaut + HauteurLigne + 2 * Intervale)
                
    'Trait horizontal entre "G-Mindmap..." et l'auteur+date
    Printer.Line (BordGauche, BordHaut + (HauteurLigne + 2 * Intervale) * 2)- _
                 (Printer.ScaleWidth, BordHaut + (HauteurLigne + 2 * Intervale) * 2)
     
                
    'Trait vertical entre la version et (date-auteur)
    Printer.Line ((Printer.ScaleWidth + BordGauche) / 2, BordHaut + (HauteurLigne + 2 * Intervale) * 2)- _
                 ((Printer.ScaleWidth + BordGauche) / 2, Printer.ScaleHeight)
                                
    'Afficher le titre
    Printer.CurrentX = BordGauche + Intervale
    Printer.CurrentY = BordHaut + Intervale
    If Len(Arbre(0).Legende) > 20 Then '20 premiers car. uniquement
        Printer.Print Left$(Arbre(0).Legende, 20)
    Else
        Printer.Print Arbre(0).Legende
    End If
                             
    'Afficher la version
    Printer.CurrentX = BordGauche + Intervale
    Printer.CurrentY = BordHaut + (HauteurLigne + 3 * Intervale)
    Printer.Print "G-Mindmap v" & App.Major & "." & App.Minor & "." & App.Revision
                     
    'Afficher l'auteur
    Printer.CurrentX = BordGauche + Intervale
    Printer.CurrentY = BordHaut + 2 * HauteurLigne + 5 * Intervale
    Printer.Print InputBox(LoadResString(OfsLanguage + 740), LoadResString(OfsLanguage + 741), LoadResString(OfsLanguage + 742))
        
    'Afficher la date
    Printer.CurrentX = (BordGauche + Printer.ScaleWidth) / 2 + Intervale
    Printer.CurrentY = BordHaut + 2 * HauteurLigne + 5 * Intervale
    Printer.Print Date
End Sub 'DessinerCartouche


'Imprimer le mindmap
Sub ImprimerMindmap()
    Dim NbreCopies As Integer
    Dim i As Integer
    
    On Error GoTo annuler
    frmMDI.cmDlgImprimer.CancelError = True
    frmMDI.cmDlgImprimer.ShowPrinter
        
    
    On Error Resume Next
    NbreCopies = frmMDI.cmDlgImprimer.Copies
    For i = 1 To NbreCopies
        ImprimerUnMindMap
    Next i
    Exit Sub
    
annuler:
    
End Sub 'ImprimerMindmap



'Impression d'un mindmap, avec les options de la boîte de dialogue d'impression
Private Sub ImprimerUnMindMap()
    'frmMap.PrintForm
    PrinterDessinerAllMindMap
End Sub 'ImprimerUnMindMap






'Dessiner un noeud
Private Sub printerDessinerNoeud(x As Long, y As Long, index As Long)
    Dim txtW As Long
    Dim txtH As Long
    Dim w As Long           'Largeur
    Dim h As Long           'Hauteur
    
    'Calculer la hauteur et la largeur
    txtW = Printer.TextWidth(Arbre(index).Legende)
    txtH = Printer.TextHeight(Arbre(index).Legende)
    w = txtW * 0.5 + Printer.TextWidth("OO")
    h = txtH * 0.5 + Printer.TextHeight("O") / 2
    
    'Dessiner le centre
    Printer.FillColor = RGB(255, 255, 255)
    Printer.FillStyle = 0 'solide
    Printer.DrawWidth = 2
    Printer.Circle (GetRealPosXprinter(x), GetRealPosYPrinter(y)), w, , , , h / w
    Printer.DrawWidth = 1
    
    'Afficher le label
    Printer.CurrentX = GetRealPosXprinter(x) - txtW / 2
    Printer.CurrentY = GetRealPosYPrinter(y) - txtH / 2
    Printer.ForeColor = 0 'Couleur du cadre
    'Printer.BackColor = RGB(255, 255, 200)
    'Printer.FillColor = RGB(0, 255, 0)
    Printer.Print Arbre(index).Legende & vbCrLf & Arbre(index).URL
End Sub 'printerDessinerNoeud



Private Sub printerDessinerNoeudEtFils(NoeudDepart As Long, Etape)
 Dim NewX As Long, NewY As Long, AngleTexte As Single, text As String, hcar As Byte, i, x As Long, y As Long
    x = Arbre(NoeudDepart).x '***/ 1000 * (Printer.ScaleWidth / 2) + Printer.ScaleWidth / 2
    y = Arbre(NoeudDepart).y '***/ 1000 * (Printer.ScaleHeight / 2) + Printer.ScaleHeight / 2
    
    
    'Dessiner les suivants
    If Arbre(NoeudDepart).NbSuivants > 0 Then
        'Afficher chaque suivant
        For i = 0 To Arbre(NoeudDepart).NbSuivants - 1
            'Coordonnées
            NewX = Arbre(Arbre(NoeudDepart).Suivants(i)).x '***/ 1000 * (Printer.ScaleWidth / 2) + Printer.ScaleWidth / 2
            NewY = Arbre(Arbre(NoeudDepart).Suivants(i)).y '***/ 1000 * (Printer.ScaleHeight / 2) + Printer.ScaleHeight / 2
            
            
            'ReCalculer l'angle du texte
            If x = NewX Then
                AngleTexte = 90
            Else
                AngleTexte = -Atn((GetRealPosYPrinter(NewY) - GetRealPosYPrinter(y)) / (GetRealPosXprinter(NewX) - GetRealPosXprinter(x))) * 180 / 3.1415926535
            End If
            
            'Tracer une ligne
            Printer.ForeColor = RGB(Etape * 64 Mod 256, Etape * 128 Mod 256, Etape * 32 Mod 256)
            Printer.DrawWidth = ((HauteurArbre(0) - Etape) / HauteurArbre(0) * 3) ^ 2 + 1
            Printer.Line (GetRealPosXprinter(x), GetRealPosYPrinter(y))-(GetRealPosXprinter(NewX), GetRealPosYPrinter(NewY))
            Printer.DrawWidth = 1
           
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
            
            XTexte = XTexte + GetVecHPrinter(Printer.TextHeight("O") / 4 * Cos((90 - Angle) * 3.1415926535 / 180) * 2)
            YTexte = YTexte + GetVecLPrinter(Printer.TextHeight("O") / 4 * Sin((90 - Angle) * 3.1415926535 / 180) * 2)
            PrintRotprinter GetRealPosXprinter(XTexte), GetRealPosYPrinter(YTexte), AngleTexte, text, hcar
                                      
            printerDessinerNoeudEtFils Arbre(NoeudDepart).Suivants(i), Etape + 1
        Next i
    End If
    
    'Dessiner la racine
    If Etape = 1 Then printerDessinerNoeud x, y, NoeudDepart
End Sub 'printerDessinerNoeudEtFils



'Dessiner tous le mindmap
Private Sub PrinterDessinerAllMindMap()
    Printer.ScaleMode = vbPixels
    'printer.Orientation = vbPRORLandscape
    Printer.PrintQuality = 600 ' vbPRPQHigh
    'Printer.TwipsPerPixelX
    
    'Dessiner un cadre
    Printer.FillStyle = vbTransparent
    Printer.FillColor = RGB(255, 255, 255)
    Printer.Line (0, 0)-(Printer.ScaleWidth - 1, Printer.ScaleHeight - 1), , B
    SetBkMode Printer.hdc, 1 'transparent=1; opaque=2
    DessinerCartouche
        
    printerDessinerNoeudEtFils Racine, 1
    Printer.EndDoc
End Sub 'PrinterDessinerAllMindMap

