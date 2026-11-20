Attribute VB_Name = "modInstallateur"
Option Explicit
'=====================================================================
' modInstallateur : construit automatiquement l'application LaboStock
' à partir des feuilles CODE_VBA (code source) et SPEC_USERFORMS
' (formulaires, contrôles et propriétés) du classeur.
'
' PRÉREQUIS (une seule fois) :
'   Fichier > Options > Centre de gestion de la confidentialité >
'   Paramètres du Centre de gestion de la confidentialité >
'   Paramètres des macros > cocher
'   « Accès approuvé au modèle d'objet du projet VBA ».
'
' UTILISATION : Alt+F8 > InstallerLaboStock > Exécuter
'=====================================================================

Private Const FEUILLE_CODE As String = "CODE_VBA"
Private Const FEUILLE_SPEC As String = "SPEC_USERFORMS"
Private mEtape As String

Public Sub InstallerLaboStock()
    Dim vbp As Object, n As Long, codes As Object, typesComp As Object, ordre As Collection
    Dim nE As Long, dE As String

    ' 1. Vérification de l'accès au projet VBA
    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    n = vbp.VBComponents.Count
    If Err.Number <> 0 Or vbp Is Nothing Then
        On Error GoTo 0
        MsgBox "L'accès au projet VBA est refusé par Excel." & vbCrLf & vbCrLf & _
               "Activez-le une seule fois :" & vbCrLf & _
               "Fichier > Options > Centre de gestion de la confidentialité > " & _
               "Paramètres du Centre de gestion de la confidentialité > Paramètres des macros > " & _
               "cocher « Accès approuvé au modèle d'objet du projet VBA »." & vbCrLf & vbCrLf & _
               "Puis relancez InstallerLaboStock.", vbCritical, "Installation LaboStock"
        Exit Sub
    End If
    On Error GoTo 0

    If Not FeuilleExiste(FEUILLE_CODE) Or Not FeuilleExiste(FEUILLE_SPEC) Then
        MsgBox "Les feuilles " & FEUILLE_CODE & " et " & FEUILLE_SPEC & " sont introuvables." & vbCrLf & _
               "Exécutez cette macro dans le classeur LaboStock fourni.", vbCritical, "Installation LaboStock"
        Exit Sub
    End If

    If MsgBox("L'installation va créer (ou recréer) :" & vbCrLf & _
              "  - 8 modules standards et 1 module de classe," & vbCrLf & _
              "  - 10 UserForms avec tous leurs contrôles et leur code," & vbCrLf & _
              "  - le code du classeur (ThisWorkbook) et la page d'accueil," & vbCrLf & _
              "puis supprimer les feuilles d'installation et enregistrer le classeur au format .xlsm." & vbCrLf & vbCrLf & "Continuer ?", _
              vbQuestion + vbYesNo, "Installation LaboStock") <> vbYes Then Exit Sub

    On Error GoTo ErrH
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    mEtape = "référence Microsoft Forms 2.0"
    AjouterReferenceMSForms vbp

    mEtape = "lecture du code source"
    LireCode codes, typesComp, ordre

    mEtape = "modules et classe"
    InstallerModules vbp, codes, typesComp, ordre

    mEtape = "formulaires"
    InstallerFormulaires vbp, codes

    Application.EnableEvents = True
    mEtape = "finalisation (boutons, stock, protection)"
    Application.Run "FinaliserInstallation"

    mEtape = "enregistrement au format .xlsm"
    EnregistrerEnXlsm
    Application.ScreenUpdating = True

    MsgBox "Installation terminée avec succès." & vbCrLf & vbCrLf & _
           "Classeur enregistré sous :" & vbCrLf & ThisWorkbook.FullName & vbCrLf & vbCrLf & _
           "Mot de passe administrateur : " & Application.Run("Param", "MOT_DE_PASSE") & vbCrLf & vbCrLf & _
           "Fermez puis rouvrez ce fichier .xlsm et activez les macros : l'application démarre seule." & vbCrLf & _
           "Vous pouvez supprimer le module modInstallateur (clic droit > Supprimer).", vbInformation, "Installation LaboStock"
    Exit Sub
ErrH:
    nE = Err.Number: dE = Err.Description
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "L'installation a échoué à l'étape : " & mEtape & vbCrLf & vbCrLf & _
           "Erreur " & nE & " : " & dE & vbCrLf & vbCrLf & _
           "Corrigez la cause puis relancez InstallerLaboStock (l'installation peut être relancée sans risque).", _
           vbCritical, "Installation LaboStock"
End Sub

Private Function FeuilleExiste(ByVal nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    FeuilleExiste = Not ws Is Nothing
    On Error GoTo 0
End Function

' Ajoute la référence Microsoft Forms 2.0 (FM20.DLL) si absente
Private Sub AjouterReferenceMSForms(vbp As Object)
    Dim r As Object, i As Long
    ' Retire les références cassées (« MANQUANT ») qui provoquent des erreurs de compilation
    For i = vbp.References.Count To 1 Step -1
        Set r = vbp.References(i)
        If r.IsBroken Then
            On Error Resume Next
            vbp.References.Remove r
            On Error GoTo 0
        End If
    Next i
    For Each r In vbp.References
        If UCase$(r.Name) = "MSFORMS" Then Exit Sub
    Next r
    On Error Resume Next
    vbp.References.AddFromGuid "{0D452EE1-E08F-101A-852E-02608C4D0BB4}", 2, 0
    On Error GoTo 0
End Sub

' Lit la feuille CODE_VBA : colonne A = composant, B = type, C = n° de ligne, D = ligne de code
Private Sub LireCode(ByRef codes As Object, ByRef typesComp As Object, ByRef ordre As Collection)
    Dim ws As Worksheet, derLig As Long, i As Long, nom As String, arr As Variant
    Set ws = ThisWorkbook.Worksheets(FEUILLE_CODE)
    Set codes = CreateObject("Scripting.Dictionary")
    Set typesComp = CreateObject("Scripting.Dictionary")
    Set ordre = New Collection
    derLig = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    arr = ws.Range("A1:D" & derLig).Value
    For i = 2 To derLig
        nom = Trim$(CStr(arr(i, 1)))
        If Len(nom) > 0 Then
            If Not codes.Exists(nom) Then
                codes.Add nom, ""
                typesComp.Add nom, CLng(arr(i, 2))
                ordre.Add nom
            End If
            codes(nom) = codes(nom) & CStr(arr(i, 4)) & vbCrLf
        End If
    Next i
End Sub

Private Sub InstallerModules(vbp As Object, codes As Object, typesComp As Object, ordre As Collection)
    Dim nom As Variant, t As Long, c As Object
    For Each nom In ordre
        t = typesComp(nom)
        Select Case t
            Case 1, 2
                SupprimerComposant vbp, CStr(nom)
                Set c = vbp.VBComponents.Add(t)
                c.Name = CStr(nom)
                EcrireCode c, codes(nom)
            Case 100
                If Len(ThisWorkbook.CodeName) > 0 Then
                    Set c = vbp.VBComponents(ThisWorkbook.CodeName)
                Else
                    Set c = vbp.VBComponents("ThisWorkbook")
                End If
                EcrireCode c, codes(nom)
            Case Else
                ' type 3 = UserForm : traité dans InstallerFormulaires
        End Select
    Next nom
End Sub

Private Sub EcrireCode(c As Object, ByVal code As String)
    With c.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString code
    End With
End Sub

Private Sub SupprimerComposant(vbp As Object, ByVal nom As String)
    Dim c As Object
    On Error Resume Next
    Set c = vbp.VBComponents(nom)
    If c Is Nothing Then Exit Sub
    If c.Type = 100 Then Exit Sub
    ' Renommer avant suppression évite l'erreur 75 lors de la recréation sous le même nom
    c.Name = Left$(nom, 20) & "_old" & Format$(Int(Rnd() * 100000), "00000")
    vbp.VBComponents.Remove c
    On Error GoTo 0
End Sub

' Lit SPEC_USERFORMS : une ligne "UserForm" suivie des lignes de ses contrôles
' Colonnes : A Formulaire | B Type | C Name | D Caption | E Left | F Top | G Width | H Height
'            I BackColor | J ForeColor | K Taille police | L Gras | M Autres propriétés
Private Sub InstallerFormulaires(vbp As Object, codes As Object)
    Dim ws As Worksheet, derLig As Long, i As Long, ligDebut As Long, typ As String
    Dim comp As Object, nomForm As String
    Set ws = ThisWorkbook.Worksheets(FEUILLE_SPEC)
    derLig = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ligDebut = 0
    For i = 1 To 20
        If Trim$(CStr(ws.Cells(i, 1).Value)) = "Formulaire" Then
            ligDebut = i + 1
            Exit For
        End If
    Next i
    If ligDebut = 0 Then Err.Raise vbObjectError + 700, , "En-tête « Formulaire » introuvable dans " & FEUILLE_SPEC
    For i = ligDebut To derLig
        typ = Trim$(CStr(ws.Cells(i, 2).Value))
        If typ = "UserForm" Then
            If Not comp Is Nothing Then EcrireCode comp, codes(nomForm)
            nomForm = Trim$(CStr(ws.Cells(i, 3).Value))
            mEtape = "formulaire " & nomForm
            If Not codes.Exists(nomForm) Then Err.Raise vbObjectError + 701, , "Code introuvable pour " & nomForm
            Set comp = CreerFormulaire(vbp, nomForm, ws, i)
        ElseIf Len(typ) > 0 Then
            mEtape = "formulaire " & nomForm & ", contrôle " & CStr(ws.Cells(i, 3).Value)
            AjouterControle comp, ws, i
        End If
    Next i
    If Not comp Is Nothing Then EcrireCode comp, codes(nomForm)
End Sub

Private Function CreerFormulaire(vbp As Object, ByVal nom As String, ws As Worksheet, ByVal i As Long) As Object
    Dim c As Object
    SupprimerComposant vbp, nom
    Set c = vbp.VBComponents.Add(3)
    On Error Resume Next
    c.Name = nom
    If Err.Number <> 0 Then
        Err.Clear
        c.Properties("Name").Value = nom
    End If
    On Error GoTo 0
    If c.Name <> nom Then Err.Raise vbObjectError + 702, , "Impossible de nommer le formulaire " & nom & _
                                    " (fermez puis rouvrez Excel et relancez l'installation)."
    c.Properties("Caption").Value = CStr(ws.Cells(i, 4).Value)
    c.Properties("Width").Value = CDbl(ws.Cells(i, 7).Value)
    c.Properties("Height").Value = CDbl(ws.Cells(i, 8).Value)
    If Len(Trim$(CStr(ws.Cells(i, 9).Value))) > 0 Then c.Properties("BackColor").Value = CouleurVBA(CStr(ws.Cells(i, 9).Value))
    On Error Resume Next
    c.Properties("StartUpPosition").Value = 1
    On Error GoTo 0
    Set CreerFormulaire = c
End Function

Private Sub AjouterControle(comp As Object, ws As Worksheet, ByVal i As Long)
    Dim ctl As Object, typ As String, progId As String, cap As String, autres As String
    Dim p As Variant, kv() As String
    If comp Is Nothing Then Err.Raise vbObjectError + 703, , "Contrôle défini avant tout UserForm (ligne " & i & ")"
    typ = Trim$(CStr(ws.Cells(i, 2).Value))
    Select Case typ
        Case "Label": progId = "Forms.Label.1"
        Case "TextBox": progId = "Forms.TextBox.1"
        Case "ComboBox": progId = "Forms.ComboBox.1"
        Case "ListBox": progId = "Forms.ListBox.1"
        Case "CommandButton": progId = "Forms.CommandButton.1"
        Case "CheckBox": progId = "Forms.CheckBox.1"
        Case "OptionButton": progId = "Forms.OptionButton.1"
        Case "Frame": progId = "Forms.Frame.1"
        Case "Image": progId = "Forms.Image.1"
        Case Else: Err.Raise vbObjectError + 704, , "Type de contrôle inconnu : " & typ
    End Select
    Set ctl = comp.Designer.Controls.Add(progId)
    ctl.Name = Trim$(CStr(ws.Cells(i, 3).Value))
    ctl.Left = CDbl(ws.Cells(i, 5).Value)
    ctl.Top = CDbl(ws.Cells(i, 6).Value)
    ctl.Width = CDbl(ws.Cells(i, 7).Value)
    ctl.Height = CDbl(ws.Cells(i, 8).Value)
    cap = CStr(ws.Cells(i, 4).Value)
    If Len(cap) > 0 Then
        Select Case typ
            Case "Label", "CommandButton", "CheckBox", "OptionButton", "Frame": ctl.Caption = cap
        End Select
    End If
    If typ <> "Image" Then                      ' le contrôle Image n'a ni police ni couleur de texte
        ctl.Font.Name = "Segoe UI"
        If Len(Trim$(CStr(ws.Cells(i, 11).Value))) > 0 Then ctl.Font.Size = CDbl(ws.Cells(i, 11).Value)
        ctl.Font.Bold = (UCase$(Trim$(CStr(ws.Cells(i, 12).Value))) = "OUI")
        If Len(Trim$(CStr(ws.Cells(i, 10).Value))) > 0 Then ctl.ForeColor = CouleurVBA(CStr(ws.Cells(i, 10).Value))
    End If
    If Len(Trim$(CStr(ws.Cells(i, 9).Value))) > 0 Then ctl.BackColor = CouleurVBA(CStr(ws.Cells(i, 9).Value))
    autres = Trim$(CStr(ws.Cells(i, 13).Value))
    If Len(autres) > 0 Then
        For Each p In Split(autres, "|")
            kv = Split(CStr(p), "=", 2)
            If UBound(kv) = 1 Then AppliquerPropriete ctl, Trim$(kv(0)), Trim$(kv(1))
        Next p
    End If
End Sub

' Applique une propriété "Nom=Valeur" (conversion automatique Booléen / Entier / Texte)
Private Sub AppliquerPropriete(ctl As Object, ByVal prop As String, ByVal valeur As String)
    Dim v As Variant
    Select Case UCase$(prop)
        Case "COLUMNWIDTHS", "CONTROLTIPTEXT", "TAG", "GROUPNAME", "TEXT"
            v = valeur
        Case Else
            If UCase$(valeur) = "TRUE" Then
                v = True
            ElseIf UCase$(valeur) = "FALSE" Then
                v = False
            ElseIf EstEntierSigne(valeur) Then
                v = CLng(valeur)
            Else
                v = valeur
            End If
    End Select
    On Error Resume Next
    CallByName ctl, prop, VbLet, v
    If Err.Number <> 0 Then
        Debug.Print "Propriété non appliquée : " & ctl.Name & "." & prop & " = " & valeur & " (" & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Function EstEntierSigne(ByVal s As String) As Boolean
    Dim i As Long, c As String
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If Not ((c >= "0" And c <= "9") Or (i = 1 And c = "-")) Then Exit Function
    Next i
    EstEntierSigne = (Len(s) < 10)
End Function

' "1F4E79" -> couleur VBA
Private Function CouleurVBA(ByVal hexa As String) As Long
    hexa = Replace(Trim$(hexa), "#", "")
    CouleurVBA = RGB(CLng("&H" & Mid$(hexa, 1, 2)), CLng("&H" & Mid$(hexa, 3, 2)), CLng("&H" & Mid$(hexa, 5, 2)))
End Function

' Enregistre le classeur au format prenant en charge les macros (.xlsm)
Private Sub EnregistrerEnXlsm()
    Dim f As String, racine As String, p As Long
    If LCase$(Right$(ThisWorkbook.Name, 5)) = ".xlsm" Then
        ThisWorkbook.Save
        Exit Sub
    End If
    racine = ThisWorkbook.Name
    p = InStrRev(racine, ".")
    If p > 0 Then racine = Left$(racine, p - 1)
    If Len(ThisWorkbook.Path) > 0 Then
        f = ThisWorkbook.Path & "\" & racine & ".xlsm"
    Else
        f = Application.DefaultFilePath & "\" & racine & ".xlsm"
    End If
    If Len(Dir$(f)) > 0 Then
        If MsgBox("Le fichier suivant existe déjà (ancienne installation) :" & vbCrLf & f & vbCrLf & vbCrLf & _
                  "Le remplacer ? (Il doit être fermé.)" & vbCrLf & "Non = enregistrer sous un nouveau nom.", _
                  vbQuestion + vbYesNo, "Installation LaboStock") <> vbYes Then
            f = Left$(f, Len(f) - 5) & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
        End If
    End If
    Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=f, FileFormat:=52
    Application.DisplayAlerts = True
End Sub
