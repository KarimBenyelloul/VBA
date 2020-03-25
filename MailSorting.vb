Private WithEvents MyIncs As Outlook.Items
Private WithEvents MySents As Outlook.Items
Private gl_PSVPattern, gl_PSVFolder, gl_TRNPattern, gl_TRNFolder, gl_SPSPattern, gl_SPSFolder As String
Private createFolder As Boolean


Private Sub Application_Startup()
    '#### VARIABLES CUSTOMISABLE PAR L'UTILISATEUR ####
    'Folder names that should appear in side the Inbox folder
    gl_TRNFolder = "Trainer_Services"
    gl_SPSFolder = "Presales_Services"
    gl_PSVFolder = "Professional_Services"
    
    'Weather or not to create the SPS/PSV/TRN folder if not found
    createFolder = True
    
    '########### Fin de la customisation ##############


    'tickets may have 5 or 4 digits
    gl_PSVPattern = "(PSV[0-9]{5})|(PSV[0-9]{4})"
    gl_TRNPattern = "(TRN[0-9]{5})|(TRN[0-9]{4})"
    gl_SPSPattern = "(SPS[0-9]{5})|(SPS[0-9]{4})"
    
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
On Error GoTo ErrorHandler
  
    'check if the item is indeed an email
    If TypeName(Item) = "MailItem" Then
        'MsgBox ("the msg is " & Item.subject)
        MailSorting Item
    End If
    

ProgramExit:
Exit Sub
ErrorHandler:
MsgBox Err.Number & " - " & Err.Description
Resume ProgramExit
  
End Sub

Private Sub MySents_ItemAdd(ByVal Item As Object)
On Error GoTo ErrorHandler

    'check if the item is indeed an email
    If TypeName(Item) = "MailItem" Then
        'MsgBox ("the msg is " & Item.subject)
        MailSorting Item
    End If
  
ProgramExit:
Exit Sub
ErrorHandler:
MsgBox Err.Number & " - " & Err.Description
Resume ProgramExit

End Sub

Public Sub MailSorting(msg As MailItem)
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    
    Dim sArray() As String
    Dim i As Integer
    Dim newCreatedFolder As folder
  
    Set RegExp = CreateObject("vbscript.regexp")
    RegExp.Global = True
    RegExp.IgnoreCase = True
    
    RegExp.Pattern = gl_SPSPattern
    Set ticketsNumbers = RegExp.Execute(msg.subject)
    If ticketsNumbers.Count <> 0 Then
        'the ticket contains an SPS pattern
        Set ticketnumber = ticketsNumbers(0)
        
        sArray = OutlookFolderNames(ns.GetDefaultFolder(olFolderInbox).folders(gl_SPSFolder))
        RegExp.Pattern = ticketnumber
        For i = 0 To UBound(sArray)
            If RegExp.test(sArray(i)) Then
                Exit For
            End If
        Next i
        If (i = UBound(sArray) + 1) Then
            'Folder not found'
            If createFolder Then
                'MsgBox ("Sub Folder " & ticketnumber & " not found inside the " & gl_SPSFolder & " Folder")
                Set newCreatedFolder = ns.GetDefaultFolder(olFolderInbox).folders(gl_SPSFolder).folders.Add(ticketnumber & " : " & Replace(msg.subject, ticketnumber, ""))
                msg.Move newCreatedFolder
            End If
        Else
            'Folder Found
            'MsgBox (ns.GetDefaultFolder(olFolderInbox).folders(gl_SPSFolder).folders(sArray(i)))
            msg.Move ns.GetDefaultFolder(olFolderInbox).folders(gl_SPSFolder).folders(sArray(i))
            
        End If
   Else
        RegExp.Pattern = gl_PSVPattern
        Set ticketsNumbers = RegExp.Execute(msg.subject)
        If ticketsNumbers.Count <> 0 Then
            'the ticket contains a PSV pattern
            Set ticketnumber = ticketsNumbers(0)
            
            sArray = OutlookFolderNames(ns.GetDefaultFolder(olFolderInbox).folders(gl_PSVFolder))
            RegExp.Pattern = ticketnumber
            For i = 0 To UBound(sArray)
                If RegExp.test(sArray(i)) Then
                    Exit For
                End If
            Next i
            If (i = UBound(sArray) + 1) Then
                'Folder not found'
                If createFolder Then
                    'MsgBox ("Sub Folder " & ticketnumber & " not found inside the " & gl_PSVFolder & " Folder")
                    Set newCreatedFolder = ns.GetDefaultFolder(olFolderInbox).folders(gl_PSVFolder).folders.Add(ticketnumber & " : " & Replace(msg.subject, ticketnumber, ""))
                    msg.Move newCreatedFolder
                End If
            Else
                'MsgBox (ns.GetDefaultFolder(olFolderInbox).folders(gl_PSVFolder).folders(sArray(i)))
                msg.Move ns.GetDefaultFolder(olFolderInbox).folders(gl_PSVFolder).folders(sArray(i))
            End If
        Else
            RegExp.Pattern = gl_TRNPattern
            Set ticketsNumbers = RegExp.Execute(msg.subject)
            If ticketsNumbers.Count <> 0 Then
                'the ticket contains a TRN pattern
                Set ticketnumber = ticketsNumbers(0)
            
                sArray = OutlookFolderNames(ns.GetDefaultFolder(olFolderInbox).folders(gl_TRNFolder))
                RegExp.Pattern = ticketnumber
                For i = 0 To UBound(sArray)
                    If RegExp.test(sArray(i)) Then
                        Exit For
                    End If
                Next i
                If (i = UBound(sArray) + 1) Then
                    'Folder not found'
                    If createFolder Then
                        'MsgBox ("Sub Folder " & ticketnumber & " not found inside the " & gl_TRNFolder & " Folder")
                        Set newCreatedFolder = ns.GetDefaultFolder(olFolderInbox).folders(gl_TRNFolder).folders.Add(ticketnumber & " : " & Replace(msg.subject, ticketnumber, ""))
                        msg.Move newCreatedFolder
                    End If
                Else
                    'MsgBox (ns.GetDefaultFolder(olFolderInbox).folders(gl_TRNFolder).folders(sArray(i)))
                    msg.Move ns.GetDefaultFolder(olFolderInbox).folders(gl_TRNFolder).folders(sArray(i))
                End If
  
            Else
                'msgbox ("the email subject doesn't contain any of the " & gl_PSVPattern & gl_SPSPattern & gl_TRNPattern)
                'msgbox ("ie it's a standard email that doesn't require sorting")
            End If
        End If
    End If
End Sub



Public Function OutlookFolderNames(folder As Outlook.folder) As String()
'Retourne la liste des sous-folder d'un folder spécifié

Dim sArray() As String
Dim i As Integer
Dim iElement As Integer
ReDim sArray(0) As String
    
On Error GoTo ErrorHandler
Set oParentFolder = folder

If oParentFolder.folders.Count Then
             
  For i = 1 To oParentFolder.folders.Count
    If Trim(oParentFolder.folders(i).Name) <> "" Then
        iElement = IIf(sArray(0) = "", 0, UBound(sArray) + 1)
        ReDim Preserve sArray(iElement) As String
        sArray(iElement) = oParentFolder.folders(i).Name
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
