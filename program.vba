Sub Bouton1_Cliquer() ' nombre de facture
    Dim resultat, objExcel

    resultat = InputBox("Entrez le nombre de facture que vous-voulez !", "Nombre de facture")

    If resultat = 0 Or resultat = "" Or resultat = " " Then
        Exit Sub
    Else
        objExcel = Sheets("home").Activate
            ActiveSheet.Range("c14") = resultat
            ActiveSheet.Range("c15") = 0
            
        Call choiceTemplate(1, CInt(resultat))
    End If

End Sub

Sub Bouton2_Cliquer() ' montant en dollars
    Dim resultat, objExcel

    resultat = InputBox("Entrez le montant que vous avez a blanchir !", "Montant a blanchir")

    If resultat = 0 Or resultat = "" Or resultat = " " Then
        Exit Sub
    Else
        objExcel = Sheets("home").Activate
            ActiveSheet.Range("c14") = 0
            ActiveSheet.Range("c15") = resultat

        Call factory(2, CInt(resultat))
    End If

End Sub

Sub choiceTemplate(ByVal choice As Integer, ByVal data As Integer) ' choix du template de la facture
    Dim template As Variant

    template = InputBox("Entrez le nombre de l'entreprise que vous-voulez pour les factures !" & _
                Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                "Tattoo            :   1" & _
                Chr(13) & Chr(10) & _
                "Coiffeur         :   2" & _
                Chr(13) & Chr(10) & _
                "Yellow Jack    :   3", "Nombre de facture")

    If template = 0 Or template = "" Or template = " " Then
        Exit Sub
    Else
        ' Call factory(choice, data, CInt(template))
    End If
End Sub

Sub factory(ByVal choice As Integer, ByVal data As Integer, ByVal template As Integer)

    Dim objExcel, objSheet, numberFacture, montantFacture, totalFacture, totalFactureAll, totalPrice, counter, loopWhile, dataInt, j
    Dim listOfMenu, listOfName As Variant
    Dim arrayMenu As Object
    
    Set arrayMenu = CreateObject("System.Collections.ArrayList")

    dataInt = data

    resultat = ThroughFiles

    If (resultat = 6) Then
        ' oui = 6
        objExcel = Sheets("home").Activate
            ActiveSheet.Range("c16") = 0
            ActiveSheet.Range("c17") = 0

        Kill ActiveWorkbook.Path & "\path\*.pdf"
    ' ElseIf (resultat = 7) Then
        ' non = 7
    ElseIf (resultat = 2) Then
        ' annuler = 2
        Exit Sub
    End If

    totalPrice = 0
    counter = 0

    objExcel = Sheets("home").Activate
        ' numberFacture = ActiveSheet.Range("c14").Value
        ' montantFacture = ActiveSheet.Range("c15").Value

        totalFacture = ActiveSheet.Range("c16").Value
        totalFactureAll = ActiveSheet.Range("c21").Value
   
    objExcel = Sheets("bdd_menu").Activate
        listOfMenu = ActiveSheet.Range("b2:c8").Value
        
    objExcel = Sheets("bdd_name").Activate
        listOfName = ActiveSheet.Range("a2:f3884").Value

    objExcel = Sheets("template").Activate
        Set objSheet = ActiveSheet
        
        ActiveSheet.Shapes.Range(Array("Picture 2")).Select
        Selection.Copy

        ActiveSheet.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
        ActiveSheet.Name = "FactureFactory"

        ActiveSheet.Paste Range("F2")
        
    Randomize

    While counter < dataInt

        randomPersonnalData = Int(UBound(listOfName) * Rnd) + 1
        Set arrayMenu = CreateObject("System.Collections.ArrayList")

        objExcel = Sheets("FactureFactory").Activate
            Set objSheet = ActiveSheet

        LineMenu = Int(3 * Rnd)
        ' quantite = Int(6 * Rnd) + 1
        totalFacture = totalFacture + 1
        j = 0
        
        For i = 1 To LineMenu - 1
            Rows("23:23").Select
            Selection.Copy
            Selection.Insert Shift:=xlDown
            Rows("24:24").Select
            Application.CutCopyMode = False
        Next i
        
        While j < LineMenu
            randomMenu = Int(UBound(listOfMenu) * Rnd) + 1
            quantite = Int(6 * Rnd) + 1

            If (Not arrayMenu.Contains(listOfMenu(randomMenu, 1))) Then
    
                arrayMenu.Add listOfMenu(randomMenu, 1)
                
                objSheet.Range("A" & (23 + j)) = listOfMenu(randomMenu, 1)
                objSheet.Range("B" & (23 + j)) = quantite
                objSheet.Range("F" & (23 + j)) = listOfMenu(randomMenu, 2)
                
                totalPrice = totalPrice + (listOfMenu(randomMenu, 2) * quantite)

                j = j + 1

            End If
        Wend
        
        objSheet.Range("G" & (25 + LineMenu)).Formula = "=SUM($G$23:G" & (23 + LineMenu) & ")"

        objSheet.Range("C12") = totalFactureAll + totalFacture
        objSheet.Range("E10") = listOfName(randomPersonnalData, 2) & " " & listOfName(randomPersonnalData, 3) & " " & listOfName(randomPersonnalData, 4)
        objSheet.Range("G11") = listOfName(randomPersonnalData, 5)
        objSheet.Range("C11") = Format(Date, "dd/mm/yyyy")

        objExcel = Sheets("bdd_name").Activate
            data = ActiveSheet.Range("f" & randomPersonnalData + 1).Value
            ActiveSheet.Range("f" & randomPersonnalData + 1) = data + 1

        ' totalPrice = totalPrice + (listOfMenu(randomMenu, 2) * quantite)
    
        objSheet.ExportAsFixedFormat 0, ActiveWorkbook.Path & "\path\ticket_client_" & totalFacture & "_" & listOfName(randomPersonnalData, 1) & ".pdf", 0, 1, 0, , , 0

        objExcel = Sheets("home").Activate
            ActiveSheet.Range("c16") = totalFacture
            
            dollars = ActiveSheet.Range("c17").Value + totalPrice
            ActiveSheet.Range("c17") = dollars
            
            ActiveSheet.Range("c21") = ActiveSheet.Range("c21").Value + 1

        If choice = "1" Then
            counter = counter + 1
        ElseIf choice = "2" Then
            counter = dollars
        End If
    Wend
    
    Application.DisplayAlerts = False
        objSheet.Delete
    Application.DisplayAlerts = True
End Sub

Public Function ThroughFiles() As Integer
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim i As Integer

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(ActiveWorkbook.Path & "\path\")
    
    If (oFolder.Files.Count > 0) Then

        ThroughFiles = MsgBox("Le dossier contient déjà des factures. (" & oFolder.Files.Count & ")" & Chr(10) & Chr(10) & "Voulez-vous les effacer ?", 3 + 48, "Fichier de facture existante (" & oFolder.Files.Count & ")")
        Exit Function

    End If
    
    LoopThroughFiles = -1
End Function

Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Public Sub CopySelectedSheets()
    Dim objExcel, objSheet

    objExcel = Sheets("template").Activate
        Set objSheet = ActiveSheet

        ActiveSheet.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
        ActiveSheet.Name = "FactureFactory"
        
        objExcel = Sheets("FactureFactory").Activate
        Set objSheet = ActiveSheet
End Sub

