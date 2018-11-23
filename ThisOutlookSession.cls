VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents AM As MailItem
Attribute AM.VB_VarHelpID = -1
Dim objMsg As Object
Private Sub Application_Quit()

    MsgBox "Goodbye, " & Application.GetNamespace("MAPI").CurrentUser
 
End Sub
Private Sub Application_MAPILogonComplete()
    
    
End Sub

Private Sub AM_PropertyChange(ByVal Name As String)
    
    Set NS = GetNamespace("MAPI")
    Set Accounts = Application.Session.Accounts
    Set objDestFolder = Nothing
    
End Sub

Private Sub Application_ItemLoad(ByVal Item As Object)

    For Each oaccount In Application.Session.Accounts
      If oaccount = "pilotage_flux_entrants@ag2rlamondiale.fr" Then
            Set Store = oaccount.DeliveryStore
            Set Folder = Store.GetDefaultFolder(olFolderInbox)
            Set FolderCom = Store.GetDefaultFolder(olFolderInbox).Parent.Folders("01_COMMUNICATION")
            Set FolderEvol = Store.GetDefaultFolder(olFolderInbox).Parent.Folders("01_DEMANDE D'EVOLUTION")
            Set FolderIncident = Store.GetDefaultFolder(olFolderInbox).Parent.Folders("01_INCIDENTS_GENERIQUES")
            Set FolderInc = Store.GetDefaultFolder(olFolderInbox).Parent.Folders("01_INCIDENTS_GENERIQUES")
            Set FolderInc = FolderInc.Folders("DEMANDE_DEPANNAGE")
            Set FolderInc2 = Store.GetDefaultFolder(olFolderInbox).Parent.Folders("01_INCIDENTS_GENERIQUES")
            Set FolderInc2 = FolderInc2.Folders("INCIDENTS_PRODUCTION")
            Set Destfolder = Folder.Parent.Folders("01_MAIL_TRAITE")
            Set Destfolfer2 = Folder.Parent.Folders("01_MAIL_TRAITE")
            
            For intCount = Folder.Items.Count To 1 Step -1
                Set objVariant = Folder.Items.Item(intCount)
                DoEvents
                If objVariant.Class = olMail And objVariant.Categories = "Traité" Then
                        objVariant.Move Destfolder
                        lngMovedItems = lngMovedItems + 1
                End If
                
            Next
            
            For intCount = FolderCom.Items.Count To 1 Step -1
                Set objVariant = FolderCom.Items.Item(intCount)
                DoEvents
                If objVariant.Class = olMail And objVariant.Categories = "Traité" Then
                    objVariant.Move Destfolder
                    lngMovedItems = lngMovedItems + 1
                End If
            Next
            
            For intCount = FolderEvol.Items.Count To 1 Step -1
                Set objVariant = FolderEvol.Items.Item(intCount)
                DoEvents
                If objVariant.Class = olMail And objVariant.Categories = "Traité" Then
                    objVariant.Move Destfolder
                    lngMovedItems = lngMovedItems + 1
                End If
            Next
            
            For intCount = FolderIncident.Items.Count To 1 Step -1
                Set objVariant = FolderIncident.Items.Item(intCount)
                DoEvents
                If objVariant.Class = olMail And objVariant.Categories = "Traité" Then
                    objVariant.Move Destfolder
                    lngMovedItems = lngMovedItems + 1
                End If
            Next
            
            For intCount = FolderInc.Items.Count To 1 Step -1
                Set objVariant = FolderInc.Items.Item(intCount)
                DoEvents
                If objVariant.Class = olMail And objVariant.Categories = "Traité" Then
                 objVariant.Move Destfolder2
                 lngMovedItems = lngMovedItems + 1
                End If
            Next
            
            For intCount = FolderInc2.Items.Count To 1 Step -1
                Set objVariant = FolderInc2.Items.Item(intCount)
                DoEvents
                If objVariant.Class = olMail And objVariant.Categories = "Traité" Then
                    objVariant.Move Destfolder2
                    lngMovedItems = lngMovedItems + 1
                End If
            Next
            
         End If
    Next
    
    Set objDestFolder = Nothing
    ' move the mails from the inbox to the "communication" box after 2 days

    Set NS = GetNamespace("MAPI")
    Set Accounts = Application.Session.Accounts
    For Each oaccount In Application.Session.Accounts
      If oaccount = "pilotage_flux_entrants@ag2rlamondiale.fr" Then

            Set Store = oaccount.DeliveryStore
            Set Folder = Store.GetDefaultFolder(olFolderInbox)
            ' use a folder under Inbox
            Set Destfolder = Folder.Parent.Folders("01_COMMUNICATION")
            For intCount = Folder.Items.Count To 1 Step -1
                Set objVariant = Folder.Items.Item(intCount)
                DoEvents
                If objVariant.Class = olMail Then
            
                    intDateDiff = DateDiff("d", objVariant.ReceivedTime, Now)
             
                    ' I'm using 3 days
                    If intDateDiff > 3 Then

                        objVariant.Move Destfolder
              
                        'count the # of items moved
                        lngMovedItems = lngMovedItems + 1

                    End If
                   
                End If
    
            Next
       
         End If
        
    Next
    'Display the number of items that were moved.
   ' MsgBox " " & lngMovedItems & " messages(s)ont été déplacés ans le dossier communication."
    Set objDestFolder = Nothing
   
End Sub
