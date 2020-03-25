Private TicketFolder As String
Private AutoFolderCreation As Boolean
Private PatternCaseCreation As String
Private WithEvents MyIncs As Outlook.Items
Private WithEvents MySents As Outlook.Items
Private dbg As Boolean
Private SearchPattern As String


Private Sub Application_Startup()

  '#### VARIABLES CUSTOMISABLE PAR L'UTILISATEUR ####
  

  'TicketFolder : entre guillemet, le nom de votre dossier (DOIT SE TROUVER DANS LA BOITE DE RECEPTION) où vous avez un sous-dossier par case.
  TicketFolder = "Professional_Services"
    
  'dbg : Active ou désactive le mode debug (affiche l'output sur la fenêtre d'exécution). True ou False.
  dbg = True
  
  '#### FIN DE LA PARTIE CUSTOMISABLE PAR L'UTILISATEUR ####
   
  'SearchPattern : match des numéros de case
  '#### ((so)?(INC)?([0-9]{1,2}-)?(c)?(r)?(req1-)?[0-9]{5,12}(-[0-9])?(pw)?)|([0-9]{4}-[0-9]{4}-[0-9]{4})####
  
  SearchPattern = "\[PSV[0-9]{4}\]"

  Dim olApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Set olApp = Outlook.Application
  Set objNS = olApp.GetNamespace("MAPI")
  
  'Boite de réception par défaut :
  Set MyIncs = objNS.GetDefaultFolder(olFolderInbox).Items
  'Boite d'éléments envoyés par défaut :
  Set MySents = objNS.GetDefaultFolder(olFolderSentMail).Items
  
 
End Sub

Private Sub MyIncs_ItemAdd(ByVal Item As Object)

    
'Lorsqu'un message arrive dans la boîte de réception, le code suivant sera executé
'  On Error GoTo ErrorHandler
  
'Pour debug
'   If dbg Then
'    Debug.Print ""
'    Debug.Print "--- NOUVEAU MESSAGE BOITE RECEPTION ---"
'  End If
  
  'Vérifie que l'élément actif est bien de type mail
'  If TypeName(Item) = "MailItem" Then
    'Effectue le tri basé les numéros d'incident
'    a = MailSorting(Item)
'  End If
    
'ProgramExit:
'  Exit Sub
'ErrorHandler:
'  MsgBox Err.Number & " - " & Err.Description
'  Resume ProgramExit
  
End Sub


Private Sub MySents_ItemAdd(ByVal Item As Object)

  Dim sArray() As String
  Dim ns As Outlook.NameSpace
  Set ns = Application.GetNamespace("MAPI")
  dim ictr as Integer
  Dim FolderID As String
  dim karim as string
  
  sArray = OutlookFolderNames(ns.GetDefaultFolder(olFolderInbox).Folders(TicketFolder))
  For ictr = 0 To UBound(sArray)
          'Recherche un numéro de case dans le nom du folder
          FolderID = caseNumberSearch(sArray(ictr))
          karim = karim & FolderID
         
          'Compare le numéro de case trouvé à celui du sujet du mail initial
  Next
  msgBox ("the folder list is" & karim )
'  On Error GoTo ErrorHandler
  
  'Pour debug
'  If dbg Then
'    Debug.Print ""
'    Debug.Print "--- NOUVEAU MESSAGE BOITE ENVOIE ---"
'  End If
  
  'Vérifie que l'élément actif est bien de type mail
'  If TypeName(Item) = "MailItem" Then
    'Effectue le tri basé les numéros d'incident
'    a = MailSorting(Item)
'  End If
  
'ProgramExit:
'  Exit Sub
'ErrorHandler:
'  MsgBox Err.Number & " - " & Err.Description
'  Resume ProgramExit
'  
End Sub


' Public Function MailSorting(Msg As MailItem) As Boolean

  Dim ns As Outlook.NameSpace
  Set ns = Application.GetNamespace("MAPI")
  
  'Variables utilisateurs
  Dim sArray() As String
  Dim ictr As Integer
  Dim TicketID As String
  Dim FolderID As String
  Dim sorted As Boolean
  Dim searchFolder As folder
  Dim myItemRec, myItem As Object
  Dim myMailItem As Outlook.MailItem

  sorted = False

  
  'Vérifie que l'item choisi est bien un mail
  If TypeName(Msg) = "MailItem" Then
      
    'initialisation du numéro de ticket
    TicketID = ""
    
    'recherche du numéro de ticket dans le sujet du message
    TicketID = caseNumberSearch(Msg.subject)
    
    'Test si un numéro de ticket valide a été trouvé
    If TicketID <> "" Then
        
        'pour debug
        If dbg = True Then
          Debug.Print "Case number détecté dans le message : " & TicketID
        End If
        
        'Récupération du répertoire contenant les dossiers ticket
        sArray = OutlookFolderNames(ns.GetDefaultFolder(olFolderInbox).Folders(TicketFolder))
        On Error Resume Next
        
        For ictr = 0 To UBound(sArray)
          'Recherche un numéro de case dans le nom du folder
          FolderID = caseNumberSearch(sArray(ictr))
          'Compare le numéro de case trouvé à celui du sujet du mail initial
          If StrComp(FolderID, TicketID, vbTextCompare) = 0 Then
            'Si match, déplace le mail dans le dossier spécifié
            Set MoveToFolder = ns.GetDefaultFolder(olFolderInbox).Folders(TicketFolder).Folders(sArray(ictr))
              'Pour debug
              If dbg = True Then
                Debug.Print "TRI! - Dossier correspondant trouvé : " & MoveToFolder
              End If
            Msg.Move MoveToFolder
            sorted = True
          End If
        Next
        
        'Si le message n'a pas été trié lors de l'étape précédente
        If sorted = False Then
        
        If dbg = True Then
          Debug.Print "Aucun match trouvé dans les noms des dossier. Recherche dans les sujets ..."
        End If
        
        
        'Récupération du répertoire contenant les dossiers ticket
        sArray = OutlookFolderNames(ns.GetDefaultFolder(olFolderInbox).Folders(TicketFolder))
        On Error Resume Next
        
        'Pour chacun de ces répertoire, récupère les mails contenus dans le folder
        For ictr = 0 To UBound(sArray)
          Set searchFolder = ns.GetDefaultFolder(olFolderInbox).Folders(TicketFolder).Folders(sArray(ictr))
          For Each myItem In searchFolder.Items
            If sorted = False Then
              If TypeName(myItem) = "MailItem" Then
                'Recherche un numéro de case dans le sujet des mails récupérés
                subjectID = caseNumberSearch(myItem.subject)
                
                
                'Si le numéro corresponds, tri le mail dans le dossier, et marque le classement comme effectué
                If StrComp(subjectID, TicketID, vbTextCompare) = 0 Then
                  'Pour debug
                  If dbg = True Then
                    Debug.Print "TRI! - Sujet correspondant trouvé dans le dossier : " & searchFolder
                  End If
                  sorted = True
                  Msg.Move searchFolder
                End If
              End If
            End If
          Next
        Next
        
        End If
    End If
 End If

'pour debug
If dbg = True And sorted = False Then
  Debug.Print "Aucun tri effectué"
End If

'Retour de value
MailSorting = sorted

End Function





Private Function caseNumberSearch(subject As String) As String

  Dim caseID As String
  
  'Creation de la regex, en non-sensitive, recurrent sur toutes les occurences, avec le pattern souhaité
  Dim RegCaseNumber
  Set RegCaseNumber = New RegExp
      RegCaseNumber.IgnoreCase = True
      RegCaseNumber.Global = True
      RegCaseNumber.Pattern = SearchPattern

  Dim Match As VBScript_RegExp_55.Match
  Dim Matches As VBScript_RegExp_55.MatchCollection
  
  caseID = ""
  
  'test si la chaine de caractère match le pattern.
  If RegCaseNumber.test(subject) Then
    Set Matches = RegCaseNumber.Execute(subject)
    For Each Match In Matches
      caseID = Match.Value
    Next
  End If
  
  'Retourne le match si trouvé. En cas de multiples match au sein de la chaine de caractère, seul le dernier sera retourné
  caseNumberSearch = caseID

End Function

Public Function OutlookFolderNames(folder As Outlook.folder) As String()
'Retourne la liste des sous-folder d'un folder spécifié

Dim sArray() As String
Dim i As Integer
Dim iElement As Integer
ReDim sArray(0) As String
    
On Error GoTo ErrorHandler
Set oParentFolder = folder

If oParentFolder.Folders.Count Then
             
  For i = 1 To oParentFolder.Folders.Count
    If Trim(oParentFolder.Folders(i).Name) <> "" Then
        iElement = IIf(sArray(0) = "", 0, UBound(sArray) + 1)
        ReDim Preserve sArray(iElement) As String
        sArray(iElement) = oParentFolder.Folders(i).Name
    End If
  Next i
Else
  
  sArray(0) = oParentFolder.Name

End If

OutlookFolderNames = sArray
Set oMAPI = Nothing
Exit Function

ErrorHandler:
    Set oMAPI = Nothing
End Function






Sub GetMailProps()
    Dim currentFolder As folder
    Dim sArray() As String
    Dim myMail As Outlook.MailItem
    Dim today As Variant
    
    sArray = OutlookFolderNames(ns.GetDefaultFolder(olFolderInbox).Folders(TicketFolder))
    For ictr = 0 To UBound(sArray)
        Set currentFolder = ns.GetDefaultFolder(olFolderInbox).Folders(TicketFolder).Folders(sArray(ictr))
        Set myItems = objFolder.Items
        myItems.Sort "CreationTime", True
        Set myItem = myItems.Item(1)
        
        If TypeName(myItem) = "MailItem" Then
            myMail = myItem
            MsgBox "Mail was sent on: " & myMail.SentOn & vbCr & _
            "by: " & myMail.SenderName & vbCr & _
            "message was received at: " & myMail.ReceivedTime
            today = Date
            diff = DateDiff("d", myMail.SentOn, today)
            If diff > 2 Then
                myMail.UnRead = True
            End If
        End If
        
    Next ictr

End Sub