Option Explicit

Sub remboursementDette()

'Ce programme calucl pour un pret, le montant a rembourser chaque
'annee pour un taux d'interet fixe au depart ainsi qu'une periode
'de remboursement pre-etabli

'Raccourcis d'execution : Ctrl + y

Dim intTaux, intAmmount, nDuree
Dim yrBegBal, yrEndBal
Dim mAnnuite, iInteret, mAmmortissement
Dim outRow, rowNum, outSheet

'******************************************
' Valeurs d'entrees
'******************************************

outRow = 5 'Indique la ligne a laquelle le tableau commence
outSheet = "Amort"

Worksheets(outSheet).Activate

'Efface les donnees du tableau pour les mettres a jour
Rows(outRow + 3 & ":" & outRow + 100).Select
Selection.Clear

'******************************************
' Valeurs d'entree de l'utilisateur
'******************************************
intTaux = Cells(2, 2).Value
nDuree = Cells(3, 2).Value
intAmmount = Cells(4, 2).Value

'Le taux d'interet ne doit pas depasser 15%
'On fixe donc une condition

If intTaux > 0.15 Then
    MsgBox "Le taux d'interet du pret ne doit pas depasser 15%. "
    End
End If

'******************************************
' Calcul des valeurs de sortie
'******************************************

'Calcul de l'annuite
mAnnuite = Pmt(intTaux, nDuree, -intAmmount, , 0)

'En premiere annee, le montant restant a payer est egale au montant initial
yrBegBal = intAmmount

'Boucle pour le calcul des valeurs de sorties du tableau d'amortissement
For rowNum = 1 To nDuree
    iInteret = yrBegBal * intTaux
    mAmmortissement = mAnnuite - iInteret
    yrEndBal = yrBegBal - mAmmortissement
    
    Cells(outRow + rowNum + 3, 3).Value = rowNum
    Cells(outRow + rowNum + 3, 4).Value = yrBegBal
    Cells(outRow + rowNum + 3, 5).Value = mAnnuite
    Cells(outRow + rowNum + 3, 6).Value = iInteret
    Cells(outRow + rowNum + 3, 7).Value = mAmmortissement
    Cells(outRow + rowNum + 3, 8).Value = yrEndBal
    
    yrBegBal = yrEndBal
    
Next rowNum

'*****************************************
' Mise en forme de la table de sortie
'*****************************************
Range(Cells(outRow + 4, 4), Cells(outRow + nDuree + 3, 8)).Select
Selection.NumberFormat = "$#,##0"


End Sub
