Attribute VB_Name = "modLanguage"
'modLanguage
'Ajouté par C.Dutoit le 19 Août 2000
'Supporte les différents language du fichier de ressource
Option Explicit

Public Const OfsFR = 1000
Public Const OfsUS = 2000
Public OfsLanguage


'Afficher les différentes zones de texte dans le bon language
Sub SetLanguagesText(Ofs As Long)
    OfsLanguage = Ofs
    
    'frmmdi
    Dim obj
    For Each obj In frmMDI.Controls
        'Gérer les menus
        If TypeName(obj) = "Menu" Then
            If Val(obj.Tag) > 0 Then
                obj.Caption = LoadResString(OfsLanguage + obj.Tag)
            End If
        End If
    Next obj
End Sub 'SetLanguagesText
