' SPDX-License-Identifier: Apache-2.0
'
' Copyright 2024 Lou Lerner
'
' This source code is licensed under the Apache License, Version 2.0.
' A copy of the license is included in the root directory of this project as LICENSE.txt.

Function CLD(PlageModaValeurs As Range, PlagePostHoc As Range, CelluleModa As Range, Optional Alpha As Variant, Optional Ordre As Variant)

'------------------------------------------------------------
On Error GoTo E:

Dim Echec As Boolean
'------------------------------------------------------------

' Déclaration des variables

Dim OrdreTri As String

Dim i As Long, j As Long, k As Long, l As Long

Dim cle As Variant

Dim TableauModaValeurs() As Variant

Dim TempArray1 As Variant
Dim TempArray2 As Variant

Dim Matrice() As Variant
Dim MatriceTri() As Variant

Dim ListeDifferences As Object

Dim Compteur As Integer

Dim CompteurTraitement As Integer
Dim CompteurTraitementPresent As Integer

Dim TableauLettres() As Variant

Dim ModaCible As Integer
Dim LettresModa As String

Dim ModaPresente As Boolean

' Configuration de l'argument "Alpha" qui permet à l'utilisateur de choisir un seuil de pvalue

    ' Si l'utilisateur n'a pas spécifié de valeur alpha -> valeur par défault (5%, CàD pvalue à 0,05)

    If IsMissing(Alpha) Then
    
        Alpha = 5 ' Valeur par défaut
        
    End If

    ' Sinon vérifier que la valeur spécifiée se situe bien entre 1 et 100 inclus, et qu'elle est numérique

    If Not IsNumeric(Alpha) Or Alpha < 1 Or Alpha > 100 Then
    
        CLD = "" ' Retourne une valeur vide si Alpha n'est pas correct
        
        Exit Function
        
    End If
    
' Configuration de l'argument "Ordre" qui permet à l'utilisateur de choisir si les modalités aux
' moyennes les plus hautes on les lettre les plus basses (ordre descendant)

    ' Vérifier si l'argument "Ordre" est présent et affecter une valeur par défaut si nécessaire
    
    If IsMissing(Ordre) Then
    
        Ordre = "descendant"
        
    End If

    ' Vérifier la valeur de l'argument "Ordre" et affecter la valeur appropriée
    
    Select Case UCase(Ordre)
    
        Case "DESCENDANT"
        
            OrdreTri = "descendant"
            
        Case "NORMAL"
        
            OrdreTri = "normal"
            
        Case Else
        
            ' Si l'argument "Ordre" n'est pas égal à "descendant" ou "normal", renvoyer une valeur vide
            
            CLD = ""
            
            Exit Function
            
    End Select

' Créer une array à 2 colonnes "TableauModaValeurs" contenant la liste des modalités et leur valeur (définie par
' le 1er argument de la fonction)

TableauModaValeurs = PlageModaValeurs.Value
   
' Creer une array "Matrice" contenant le tableau indiquant les résultats du post-hoc pour chaque couple de modalités
' (définit par le 2eme argument de la fonction)

    ' Definir l'array depuis la plage
 
    Matrice() = PlagePostHoc.Value
    
    ' Remplacer les valeurs vides en dessous de la diagonale par la valeur "CelluleVide", ce sera utile plus tard
    ' pour ne pas prendre en compte ces valeurs
    
    For i = 2 To UBound(Matrice, 1)
        For j = 2 To UBound(Matrice, 2) - (UBound(Matrice, 2) - i)
            Matrice(i, j) = "CelluleVide"
        Next j
    Next i

' Supprimer les modalités qui n'ont pas été testées: cela permet au code de fonctionner même si une modalité
' n'a pas pu être testée

Dim nbLignes As Long, nbColonnes As Long
Dim ligneVide As Boolean, colonneVide As Boolean

nbLignes = UBound(Matrice, 1)
nbColonnes = UBound(Matrice, 2)

    Do
        ligneVide = False
        
        colonneVide = False
        
        For i = nbLignes To 2 Step -1
        
            ' Vérifier si la ligne i est vide
            
            ligneVide = True
            
            For j = 2 To nbColonnes
            
                If Matrice(i, j) <> "" And Matrice(i, j) <> "CelluleVide" Then
                
                    ligneVide = False
                    
                    Exit For
                    
                End If
                
            Next j
            
            ' Vérifier si la colonne correspondante est également vide
            
            If ligneVide Then
            
                colonneVide = True
                
                For j = 2 To nbColonnes
                
                    If Matrice(j, i) <> "" And Matrice(j, i) <> "CelluleVide" Then
                    
                        colonneVide = False
                        
                        Exit For
                        
                    End If
                    
                Next j
                
            End If
            
            ' Suppression de la ligne et de la colonne si elles sont vides
            
            If ligneVide And colonneVide Then
            
                Matrice = SupprimerLigne(Matrice, i)
                
                Matrice = SupprimerColonne(Matrice, i)
                
                TableauModaValeurs = SupprimerLigne(TableauModaValeurs, i - 1) ' Supprimer également la modalité dans "TableauModaValeurs" pour qu'elle soit considérée comme absente
                
                nbLignes = nbLignes - 1
                
                nbColonnes = nbColonnes - 1

                ligneVide = True
                
            End If
  
        Next i
        
    Loop While ligneVide Or colonneVide

' Si l'argument "Ordre" est définit par "normal", ne pas modifier l'ordre des modalités

If OrdreTri = "normal" Then

' Si l'argument "Ordre" est définit en "descendant", trier l'array dans l'ordre décroissant à partir des valeurs
' (cela permet qu'à la fin par exemple la lettre "a" soit assignée à la modalité la plus élevée)

ElseIf OrdreTri = "descendant" Then

    For i = LBound(TableauModaValeurs) To UBound(TableauModaValeurs) - 1
    
        For j = i + 1 To UBound(TableauModaValeurs)
        
            If TableauModaValeurs(i, 2) > TableauModaValeurs(j, 2) Then
            
                TempArray1 = TableauModaValeurs(j, 1)
                TempArray2 = TableauModaValeurs(j, 2)
                
                TableauModaValeurs(j, 1) = TableauModaValeurs(i, 1)
                TableauModaValeurs(j, 2) = TableauModaValeurs(i, 2)
                
                TableauModaValeurs(i, 1) = TempArray1
                TableauModaValeurs(i, 2) = TempArray2
                
            End If
            
        Next j
        
    Next i
    
    ' Trier la Matrice en fonction de l'array triée, cela permet que plus tard la modalité la plus élevée ait la lettre "a".
     
        ' Creer une copie de Matrice appelée "MatriceTri" et la trier pour que les modalités soient dans le même ordre que
        ' l'array triée "TableauModaValeurs"
    
        ReDim MatriceTri(1 To UBound(Matrice, 1), 1 To UBound(Matrice, 2))
            
            For i = 2 To UBound(Matrice, 1)
            
                For k = LBound(TableauModaValeurs) To UBound(TableauModaValeurs)
                
                    If Matrice(i, 1) = TableauModaValeurs(k, 1) Then
                    
                        MatriceTri(k + 1, 1) = Matrice(i, 1)
                        
                        Exit For
                        
                    End If
                    
                Next k
                
                For j = 2 To UBound(Matrice, 2)
                
                    For l = LBound(TableauModaValeurs) To UBound(TableauModaValeurs)
                    
                        If Matrice(1, j) = TableauModaValeurs(l, 1) Then
                        
                            MatriceTri(1, l + 1) = Matrice(1, j)
                            
                            Exit For
                            
                        End If
                        
                    Next l
                    
                    MatriceTri(k + 1, l + 1) = Matrice(i, j)
                    
                Next j
                
            Next i
            
        ' Remplacer le contenu de Matrice par celui de MatriceTri
        
        For i = LBound(Matrice, 1) To UBound(Matrice, 1)
        
            For j = LBound(Matrice, 2) To UBound(Matrice, 2)
            
                Matrice(i, j) = MatriceTri(i, j)
                
            Next j
            
        Next i
        
End If

' Creer un scripting dictionary pour lister les différences significatives. Chaque élément clé de cette
' liste comprend une modalité et une modalités différente à cette dernière (listées dans une array)

Set ListeDifferences = CreateObject("Scripting.Dictionary")

    ' Peupler la liste de différences en lisant la Matrice
    
    For i = 2 To UBound(Matrice, 1) ' lire chaque ligne de la matrice
   
            For j = 2 To UBound(Matrice, 2) ' lire chaque colonne de la matrice
            
                ' Si une des cellules est vide, renvoyer une valeur vide
                
                If Matrice(i, j) = "" Then
                
                    CLD = ""
                    
                    Exit Function
                   
                ' Si la valeur observée est "CelluleVide" ne rien faire car il s'agit de la zone hors-tableau
            
                ElseIf Matrice(i, j) = "CelluleVide" Then
            
                ' Si une modalité est différente d'une autre, ajouter les deux dans une clé du dictionnaire
            
                ElseIf (IsNumeric(Matrice(i, j)) And Matrice(i, j) <= Alpha / 100) Or (Matrice(i, j) = "<0,001" Or _
                                                                                    Matrice(i, j) = "<0,005" Or _
                                                                                    Matrice(i, j) = "<0,05" Or _
                                                                                    Matrice(i, j) = "<0,01" Or _
                                                                                    Matrice(i, j) = "< 0,001" Or _
                                                                                    Matrice(i, j) = "< 0,005" Or _
                                                                                    Matrice(i, j) = "< 0,05" Or _
                                                                                    Matrice(i, j) = "< 0,01") Then
                
                    Compteur = ListeDifferences.Count + 1
                    
                    ListeDifferences.Add Compteur, Array()
                    
                    ' Utilisation de la fonction perso "PeuplerArray" pour ajouter un élément à une clé du dictionaire
                    
                    ListeDifferences(Compteur) = PeuplerArray(ListeDifferences(Compteur), Matrice(i, 1))

                    ListeDifferences(Compteur) = PeuplerArray(ListeDifferences(Compteur), Matrice(1, j))
                    
                ' Si une modalité n'est pas significativement différente d'une autre ne rien faire
                    
                ElseIf IsNumeric(Matrice(i, j)) And Matrice(i, j) > Alpha / 100 Then
                
                ' Si la case est remplie avec une valeur inatendue, renvoyer une cellule vide
                                
                Else
                
                    CLD = ""
                    
                    Exit Function
                    
                End If
   
            Next j
       
    Next i
      
' Phase "insérer et absorber": créer une array avec pour chaque colonne des "1" partagés entre les modalités

    ' Créer une première colonne dans l'array avec uniquement des "1"
    
    ReDim TableauLettres(1 To UBound(TableauModaValeurs, 1), 1 To 1)
       
    For i = 1 To UBound(TableauLettres)

    TableauLettres(i, 1) = "1"
    
    Next i
        
    ' Insérer et absorber pour chaque couple de modalités inégales
        
    For Each cle In ListeDifferences.Keys ' Pour chaque couple de moda significativement différentes

        i = 1
    
        ' Vérifier pour chaque colonne du tableau si elle vérifie la différence entre les deux moda.
    
        Do Until i = UBound(TableauLettres, 2) + 1 ' Pour chaque colonne
        
                'Si l'affirmation n'est pas vérifiée insérer/absorber la colonne grace à la fonction
  
                If AffirmationVerifiee(ListeDifferences(cle), TableauLettres, i, TableauModaValeurs) = False Then
                
                    InsererAbsorber ListeDifferences(cle), TableauLettres, i, TableauModaValeurs
                                       
                    i = 1
                    
                Else
                
                    i = i + 1
                
                End If
        Loop

    Next cle

' Phase "Nettoyage": après la précédente étape il reste certains "1" inutiles qui peuvent être supprimés
    
For i = 1 To UBound(TableauLettres, 2) ' Pour chaque colonne

    For j = 1 To UBound(TableauLettres, 1) ' Pour chaque modalité (cellule)
        
        If TableauLettres(j, i) = 0 Then ' Si la case est égal à 0, ne rien faire
        
        Else ' Si la case n'est pas égal à 0
        
            CompteurTraitement = 1
            CompteurTraitementPresent = 1
            
            For k = 1 To UBound(TableauLettres, 1) ' Pour chaque traitement à comparer (ligne)
            
                If k = j Then ' (ne pas comparer la modalité à elle même)
                
                Else
                
                    If TableauLettres(k, i) = 1 Then ' Si la case du traitement ET la case du traitement à comparer sont égales à 1
                                            
                        CompteurTraitement = CompteurTraitement + 1
                                            
                        For l = 1 To UBound(TableauLettres, 2) ' Pour chaque colonnes
                        
                            If l = i Then ' (Ne pas comparer la colonne à elle-même)
                            
                            Else
                        
                                If TableauLettres(j, l) = 1 And TableauLettres(k, l) = 1 Then ' Si les deux traitement sont dans une autre colonne
                                
                                    CompteurTraitementPresent = CompteurTraitementPresent + 1 ' Décompter
                                
                                    Exit For
                                
                                End If
                                
                            End If
                        
                        Next l
                    
                    End If
                
                End If
                         
            Next k
            
            If CompteurTraitementPresent = 1 And CompteurTraitement = 1 Then ' S'il n'y a qu'un "1"
            
                For m = 1 To UBound(TableauLettres, 2) ' Pour chaque colonne
                
                    If m = i Then ' (Ne pas comparer la colonne à elle même)
                    
                    Else
                    
                        If TableauLettres(j, m) = 1 Then ' Si un autre "1" est présent sur la même ligne, le "1" isolé peut être supprimé
                        
                            TableauLettres(j, i) = 0
                            
                            Exit For
                            
                        End If
                    
                    End If
                                    
                Next m
        
            ElseIf CompteurTraitementPresent = CompteurTraitement Then ' Si toutes les égalités sont déjà indiquées dans d'autre colonnes, le "1" peut être supprimé
            
                TableauLettres(j, i) = 0
                
            End If
            
            CompteurTraitementPresent = 1 ' Réinitialiser le compteur
        
        End If

    Next j

Next i

' Suppression des colonnes vides: parfois après nettoyage des colonnes n'ont aucun "1", hors ce n'est pas utile.

i = 1

Do Until i > UBound(TableauLettres, 2) ' Pour chaque colonne

    CompteurTraitement = 0

    For j = 1 To UBound(TableauLettres, 1) ' Pour chaque modaltié (ligne)

        If TableauLettres(j, i) = 1 Then ' Si la case est égal à "1"

            CompteurTraitement = CompteurTraitement + 1 ' Décompter

        End If

    Next j

    If CompteurTraitement = 0 Then ' Si la colonne ne contient aucun "1"
        
        TableauLettres = SupprimerColonne(TableauLettres, i) ' Supprimer la colonne
                
        i = 1
        
    Else
    
    i = i + 1
    
    End If

Loop

' Ordonner les colonnes en fonction de l'odre des modalités: par exemple pour des lettres descendantes,
' pour que la modalité la plus élevée ait par exemple la lettre "a", le tableau doit être trié de manière
' à ce que les modalités les plus basses dans le tableau aient les lettres les plus basses

If OrdreTri = "descendant" Then

TrierLettresDescendant TableauLettres ' Utilisation de la fonction "TrierLettres" pour trier les colonnes

ElseIf OrdreTri = "normal" Then

TrierLettresNormal TableauLettres

End If

' Création des lettres concaténées

' Définition de la modalité cible en fonction de la cellule selectionnée dans la feuille

ModaPresente = False

For i = 1 To UBound(TableauModaValeurs, 1)

    If CelluleModa.Value = TableauModaValeurs(i, 1) Then
    
        ModaCible = i
        
        ModaPresente = True ' Vérification de l'existance de la modalité indiquée dans la cellule dans le tableau
    
    End If

Next i

' Si la modalité n'a pas été testée (absente dans les tableaux), ne rien affichier

If ModaPresente = False Then

CLD = "" ' Retourne une valeur vide

Exit Function

End If

' Dans la ligne de la modalité cible, si une ou des colonne contiennent 1, alors ajouter une lettre selon le n° de colonne (colonne 1 = b, colonne 1+3 = ac...)

For i = 1 To UBound(TableauLettres, 2)

If TableauLettres(ModaCible, i) = 1 Then

LettresModa = LettresModa + (Chr(96 + i))

End If

Next i

' Fin: restituer les lettres en tant que résultat de la fonction

CLD = LettresModa

'------------------------------------------------------------ Gestion des erreurs inatendues

Finalisation:

On Error Resume Next:

If Not Echec Then

Else

CLD = ""

End If

Exit Function

'------------------------------------------------------------

E:

Echec = True

Resume Finalisation:
      
End Function

Function PeuplerArray(ByVal arr As Variant, ByVal element As Variant) As Variant

' Fonction servant à ajouter un élément à un tableau array (en l'occurence les array contenues dans chaques clés d'un dictionnaire)

Dim i As Integer
Dim ArrayModif() As Variant

' Création d'une nouvelle array à partir de l'array de base (définie par le 1er argument de la fonction), avec un espace en plus.

ReDim ArrayModif(LBound(arr) To UBound(arr) + 1)

' Copie des éléments depuis la premiere array vers la nouvelle

For i = LBound(arr) To UBound(arr)
    ArrayModif(i) = arr(i)
Next i

' Ajouter l'élément (définit par le 2eme argument de la fonction) à la fin de l'array

ArrayModif(UBound(ArrayModif)) = element

' La valeur donnée par la fonction est maintenant l'array modifiée

PeuplerArray = ArrayModif
    
End Function

Function AffirmationVerifiee(ByVal Couple As Variant, ByRef TableauLettres As Variant, ByVal ColonneCible As Integer, ByVal TableauModaValeurs) As Boolean

' Cette fonction vérifie si les deux modalités ont un "1" dans une même colonne.
' Si tel est le cas cela va contre l'affirmation indiquant une différence significative entre les deux modalités
' La fonction renvoie donc AffirmationVerifiee = False, sinon, True.

Dim i As Integer
Dim Moda1Presente As Boolean
Dim Moda2Presente As Boolean

' Vérification de la présence de la 1ère modalité du couple dans la colonne cible

For i = 1 To UBound(TableauLettres, 1)
    If TableauModaValeurs(i, 1) = Couple(0) Then
        If TableauLettres(i, ColonneCible) = 1 Then
            Moda1Presente = True
        Else
            Moda1Presente = False
        End If
    End If
Next i

' Vérification de la présence de la 2ème modalité du couple dans la colonne cible

For i = 1 To UBound(TableauLettres, 1)
    If TableauModaValeurs(i, 1) = Couple(1) Then
        If TableauLettres(i, ColonneCible) = 1 Then
            Moda2Presente = True
        Else
            Moda2Presente = False
        End If
    End If
Next i

' Si les deux modalités sont présentes, passer la variable en "False", sinon "True".

If Moda1Presente = True And Moda2Presente = True Then
    AffirmationVerifiee = False
Else
    AffirmationVerifiee = True
End If

End Function

Sub InsererAbsorber(ByVal Couple As Variant, ByRef TableauLettres As Variant, ByVal ColonneCible As Integer, ByVal TableauModaValeurs)

' Fonction permettant de duppliquer une colonne dans un tableau array, l'absorber dans une autre si besoin puis supprimer les "1" correspondant aux modalites

Dim i As Integer, j As Integer

Dim TableauModif() As Variant

Dim Absorber As Boolean

' Copie de la colonne

    ' Créer un tableau avec une colonne en plus
    
    ReDim TableauModif(LBound(TableauLettres, 1) To UBound(TableauLettres, 1), LBound(TableauLettres, 2) To UBound(TableauLettres, 2) + 1)
    
    ' Duppliquer la colonne cible
    
    For i = LBound(TableauLettres, 1) To UBound(TableauLettres, 1)
        For j = LBound(TableauLettres, 2) To UBound(TableauLettres, 2)
            If j < ColonneCible Then
                TableauModif(i, j) = TableauLettres(i, j)
            ElseIf j = ColonneCible Then
                TableauModif(i, j) = TableauLettres(i, j)
                TableauModif(i, j + 1) = TableauLettres(i, j)
            ElseIf j > ColonneCible Then
                TableauModif(i, j + 1) = TableauLettres(i, j)
            End If
        Next j
    Next i
    
    ' Appliquer la modification au tableau original
    
    TableauLettres = TableauModif

' Dans la colonne originale supprimer le "1" de la premiere modalite

For i = 1 To UBound(TableauLettres, 1)
    If TableauModaValeurs(i, 1) = Couple(0) Then
        TableauLettres(i, ColonneCible) = 0
    End If
Next i

' Vérifier si la colonne originale modifiée peut être absorbée par une colonne précédente.

    ' La premiere colonne ne peut jamais être absorbée
    
    If ColonneCible <= 1 Then
    
    Absorber = False
    
    Else

        ' Parcourir les colonnes précédentes à la colonne cible
        
        For j = 1 To ColonneCible - 1
        
        Absorber = True
        
            ' Parcourir les lignes
            
            For i = 1 To UBound(TableauLettres, 1)
            
                ' Si pour une ligne, seule la cellule de la colonne cible est remplie, pas d'absorbtion.
    
                If TableauLettres(i, ColonneCible) = 1 And TableauLettres(i, j) = 0 Then
                
                    Absorber = False
                    
                End If
    
            Next i
            
        Next j
    
    End If

' Si la colonne originale peut bien être absorbée, l'absorber (la supprimer).

If Absorber = True Then

TableauLettres = SupprimerColonne(TableauLettres, ColonneCible) ' Uilisation de la fonction SupprimerColonne qui supprime une colonne et décale les colonnes en conséquence

ColonneCible = ColonneCible - 1 ' Si la colonne originale a été absorbée, la colonne ajoutée se retrouve reculée d'un cran

End If

' Dans la colonne ajoutée supprimer le "1" de la deuxième modalite

For i = 1 To UBound(TableauLettres, 1)

    If TableauModaValeurs(i, 1) = Couple(1) Then
    
        TableauLettres(i, ColonneCible + 1) = 0
        
    End If
    
Next i

' Vérifier si la colonne ajoutée modifiée peut être absorbée par une colonne précédente.

    ' La premiere colonne ne peut jamais être absorbée
    
    If ColonneCible + 1 <= 1 Then
    
    Absorber = False
    
    Else
    
    Absorber = True
    
        ' Parcourir les colonnes précédentes à la colonne cible
        
        For j = 1 To ColonneCible + 1 - 1
        
            ' Parcourir les lignes
            
            For i = 1 To UBound(TableauLettres, 1)
            
                ' Si pour une ligne, seule la cellule de la colonne cible est remplie, pas d'absorbtion.
    
                If TableauLettres(i, ColonneCible + 1) = 1 And TableauLettres(i, j) = 0 Then
                
                    Absorber = False
                    
                End If
    
            Next i
            
        Next j
    
    End If

' Si la colonne ajoutée peut bien être absorbée, l'absorber (la supprimer).

If Absorber = True Then

TableauLettres = SupprimerColonne(TableauLettres, ColonneCible + 1)

End If

End Sub

Function SupprimerColonne(ByRef TableauCible As Variant, ByVal ColonneCible As Long) As Variant

Dim TableauModif() As Variant

Dim i As Long, j As Long, k As Long

' Creer un tableau array avec une colonne en moins

ReDim TableauModif(LBound(TableauCible, 1) To UBound(TableauCible, 1), LBound(TableauCible, 2) To UBound(TableauCible, 2) - 1)

' Si la colonne cible est la dernière colonne du tableau, juste copier le contenu du tableau original

If ColonneCible = UBound(TableauCible, 2) Then

    For i = LBound(TableauCible, 1) To UBound(TableauCible, 1)
        For j = LBound(TableauCible, 2) To UBound(TableauCible, 2) - 1
            TableauModif(i, j) = TableauCible(i, j)
        Next j
    Next i
    
' Sinon, copier le contenu du tableau original jusqu'à la colonne précédent la colonne cible, puis copier la fin du tableau original de la colonne cible jusqu'à la fin

Else

    For i = LBound(TableauCible, 1) To UBound(TableauCible, 1)
        For j = LBound(TableauCible, 2) To ColonneCible - 1
            TableauModif(i, j) = TableauCible(i, j)
        Next j
    Next i
    
    For i = LBound(TableauCible, 1) To UBound(TableauCible, 1)
        For j = ColonneCible + 1 To UBound(TableauCible, 2)
            TableauModif(i, j - 1) = TableauCible(i, j)
        Next j
    Next i
    
End If

' Appliquer les modifications

SupprimerColonne = TableauModif
    
End Function

Function SupprimerLigne(ByRef TableauCible As Variant, ByVal LigneCible As Long) As Variant

Dim TableauModif() As Variant

Dim i As Long, j As Long, k As Long

' Créer un tableau array avec une ligne en moins

ReDim TableauModif(LBound(TableauCible, 1) To UBound(TableauCible, 1) - 1, LBound(TableauCible, 2) To UBound(TableauCible, 2))

' Si la ligne cible est la dernière ligne du tableau, juste copier le contenu du tableau original

If LigneCible = UBound(TableauCible, 1) Then

    For i = LBound(TableauCible, 1) To UBound(TableauCible, 1) - 1
        For j = LBound(TableauCible, 2) To UBound(TableauCible, 2)
            TableauModif(i, j) = TableauCible(i, j)
        Next j
    Next i

' Sinon, copier le contenu du tableau original jusqu'à la ligne précédant la ligne cible, puis copier la fin du tableau original de la ligne cible jusqu'à la fin

Else

    For i = LBound(TableauCible, 1) To LigneCible - 1
        For j = LBound(TableauCible, 2) To UBound(TableauCible, 2)
            TableauModif(i, j) = TableauCible(i, j)
        Next j
    Next i

    For i = LigneCible + 1 To UBound(TableauCible, 1)
        For j = LBound(TableauCible, 2) To UBound(TableauCible, 2)
            TableauModif(i - 1, j) = TableauCible(i, j)
        Next j
    Next i

End If

' Appliquer les modifications

SupprimerLigne = TableauModif

End Function

Sub TrierLettresDescendant(ByRef TableauLettres As Variant)

' Cette fonction trie le tableau de lettres pour que les lettres assignées dépendent de l'ordre des modalités

Dim TableauLettresTri As Variant

Dim i As Long, j As Long, k As Long

Dim NombreColonnes As Long

Dim SommeColonne() As Double

Dim tempArray() As Variant

' Plus une madalité est élevée, plus elle est basse dans le tableau. Pour que les premières lettres soient associées
' aux valeurs les plus hautes, il faut créer des rangs de colonnes selon la présence de lettre et leur position dans la colonne.
' Pour ce faire créer une matrice où les 1 sont remplacés par 2^x, où x est le n° de ligne, et la somme de ces valeurs dans
' chaque colonne permettera de trier les colonnes du tableau. /!\ 2^x étant une fonction exponentielle, le nombre de traitement
' impacte fortement la vitesse de calcul

' Création du tableau avec les valeurs 2^x
    
TableauLettresTri = TableauLettres

For i = 1 To UBound(TableauLettresTri, 1)

    For j = 1 To UBound(TableauLettresTri, 2)
    
        If TableauLettresTri(i, j) = 1 Then
        
            TableauLettresTri(i, j) = 2 ^ i
        
        End If
    
    Next j

Next i

' Trier le tableau grace à ce nouveau tableau

    NombreColonnes = UBound(TableauLettresTri, 2)
    
    ' Calculer la somme de chaque colonne
    
    ReDim SommeColonne(1 To NombreColonnes)
    
    For i = 1 To NombreColonnes
    
        For j = 1 To UBound(TableauLettresTri, 1)
        
            SommeColonne(i) = SommeColonne(i) + TableauLettresTri(j, i)
            
        Next j
        
    Next i
    
    ' Réorganiser les colonnes en fonction des sommes
    
    For i = 1 To NombreColonnes
    
        For j = i To NombreColonnes
        
            If SommeColonne(j) > SommeColonne(i) Then
            
                ' Échanger les sommes
                
                Dim temp As Double
                
                temp = SommeColonne(i)
                
                SommeColonne(i) = SommeColonne(j)
                
                SommeColonne(j) = temp
    
                ' Échanger les colonnes dans le tableau
                
                ReDim tempArray(1 To UBound(TableauLettres, 1))
                
                For k = 1 To UBound(TableauLettres, 1)
                
                    tempArray(k) = TableauLettres(k, i)
                    
                    TableauLettres(k, i) = TableauLettres(k, j)
                    
                    TableauLettres(k, j) = tempArray(k)
                    
                Next k
                
            End If
            
        Next j
        
    Next i

End Sub

Sub TrierLettresNormal(ByRef TableauLettres As Variant)

' Cette fonction trie le tableau de lettres pour que les lettres assignées dépendent de l'ordre des modalités

Dim TableauLettresTri As Variant

Dim i As Long, j As Long, k As Long

Dim NombreColonnes As Long

Dim SommeColonne() As Double

Dim tempArray() As Variant

' Plus une madalité est élevée, plus elle est basse dans le tableau. Pour que les premières lettres soient associées
' aux valeurs les plus hautes, il faut créer des rangs de colonnes selon la présence de lettre et leur position dans la colonne.
' Pour ce faire créer une matrice où les 1 sont remplacés par 2^x, où x est le n° de ligne, et la somme de ces valeurs dans
' chaque colonne permettera de trier les colonnes du tableau. /!\ 2^x étant une fonction exponentielle, le nombre de traitement
' impacte fortement la vitesse de calcul

' Création du tableau avec les valeurs 2^x
    
TableauLettresTri = TableauLettres

For i = 1 To UBound(TableauLettresTri, 1)

    For j = 1 To UBound(TableauLettresTri, 2)
    
        If TableauLettresTri(i, j) = 1 Then
        
            TableauLettresTri(i, j) = 2 ^ i
        
        End If
    
    Next j

Next i

' Trier le tableau grace à ce nouveau tableau

    NombreColonnes = UBound(TableauLettresTri, 2)
    
    ' Calculer la somme de chaque colonne
    
    ReDim SommeColonne(1 To NombreColonnes)
    
    For i = 1 To NombreColonnes
    
        For j = 1 To UBound(TableauLettresTri, 1)
        
            SommeColonne(i) = SommeColonne(i) + TableauLettresTri(j, i)
            
        Next j
        
    Next i
    
    ' Réorganiser les colonnes en fonction des sommes
    
    For i = 1 To NombreColonnes
    
        For j = i To NombreColonnes
        
            If SommeColonne(j) < SommeColonne(i) Then
            
                ' Échanger les sommes
                
                Dim temp As Double
                
                temp = SommeColonne(i)
                
                SommeColonne(i) = SommeColonne(j)
                
                SommeColonne(j) = temp
    
                ' Échanger les colonnes dans le tableau
                
                ReDim tempArray(1 To UBound(TableauLettres, 1))
                
                For k = 1 To UBound(TableauLettres, 1)
                
                    tempArray(k) = TableauLettres(k, i)
                    
                    TableauLettres(k, i) = TableauLettres(k, j)
                    
                    TableauLettres(k, j) = tempArray(k)
                    
                Next k
                
            End If
            
        Next j
        
    Next i

End Sub
