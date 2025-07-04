VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "omOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' http://www.slipstick.com/exchange/sending-email-from-a-secondary-exchange-mailbox/

'Dim objOutlook As New Outlook.Application
Dim objOutlook As Object

'outlook.OlImportance
Public Enum OlImportance
    olImportanceLow = 0
    olImportanceNormal = 1
    olImportanceHigh = 2
End Enum

'Outlook.OlItemType
Public Enum OlItemType
    olMailItem = 0
    olAppointmentItem = 1
    olContactItem = 2
    olTaskItem = 3
    olJournalItem = 4
    olNoteItem = 5
    olPostItem = 6
    olDistributionListItem = 7
    olMobileItemSMS = 11
    olMobileItemMMS = 12
End Enum

'Outlook.OlBodyFormat
Public Enum OlBodyFormat
    olFormatUnspecified = 0
    olFormatPlain = 1
    olFormatHTML = 2
    olFormatRichText = 3
End Enum

'Outlook.OlMailRecipientType
Public Enum OlMailRecipientType
    olOriginator = 0
    olTo = 1
    olCC = 2
    olBCC = 3
End Enum

'Outlook.OlDefaultFolders
Public Enum OlDefaultFolders
    olFolderDeletedItems = 3
    olFolderOutbox = 4
    olFolderSentMail = 5
    olFolderInbox = 6
    olFolderCalendar = 9
    olFolderContacts = 10
    olFolderJournal = 11
    olFolderNotes = 12
    olFolderTasks = 13
    olFolderDrafts = 16
    olPublicFoldersAllPublicFolders = 18
    olFolderConflicts = 19
    olFolderSyncIssues = 20
    olFolderLocalFailures = 21
    olFolderServerFailures = 22
    olFolderJunk = 23
    olFolderRssFeeds = 25
    olFolderToDo = 28
    olFolderManagedEmail = 29
    olFolderSuggestedContacts = 30
End Enum

'Public Sub Mail(EmailAddresses As String, Subject As String, Body As String, replyAddress As String, Importance As Outlook.OlImportance, ReadReceipt As Boolean, attachments() As String, Popup As Boolean, Optional BCCAddresses As String = "", Optional CCAddresses As String = "", Optional font As String = "Arial", Optional fontSize As String = "14.5px", Optional keepSignature As Boolean = True)
Public Sub Mail(EmailAddresses As String, subject As String, body As String, replyAddress As String, Importance As OlImportance, ReadReceipt As Boolean, attachments() As String, Popup As Boolean, Optional bccAddresses As String = "", Optional ccAddresses As String = "", Optional font As String = "Arial", Optional fontSize As String = "14.5px", Optional keepSignature As Boolean = True, Optional useReplyAddressAsOnBehalf As Boolean = False)
'Dim objOutlookMsg As Outlook.MailItem
'Dim objOutlookRecip As Outlook.Recipient
'Dim objOutlookRecipBCC As Outlook.Recipient

Dim objOutlookMsg As Object
Dim objOutlookRecip As Object
Dim objOutlookRecipBCC As Object
Dim mailbox As Object


Dim i As Integer
Dim filename As String
Dim bodyFilenameMissing As String
Dim address() As String
Dim signature As String
Dim bodyStart As Long
Dim bodyEnd As Long

    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)
    DoEvents
    With objOutlookMsg
        .Display
        If NotIsNullOrEmpty(.body) And keepSignature Then
            signature = .HTMLBody
        End If
        .BodyFormat = olFormatHTML
        If NotIsNullOrEmpty(bccAddresses) Then
            address = StringSplit(bccAddresses, ";")
            For i = 0 To omArrayFunctions.StringArrayCount(address) - 1
                Set objOutlookRecip = .Recipients.Add(address(i))
                objOutlookRecip.Type = olBCC
            Next
        End If
        If NotIsNullOrEmpty(ccAddresses) Then
            address = StringSplit(ccAddresses, ";")
            For i = 0 To omArrayFunctions.StringArrayCount(address) - 1
                Set objOutlookRecip = .Recipients.Add(address(i))
                objOutlookRecip.Type = olCC
            Next
        End If
        If NotIsNullOrEmpty(EmailAddresses) Then
            address = StringSplit(EmailAddresses, ";")
            For i = 0 To omArrayFunctions.StringArrayCount(address) - 1
                Set objOutlookRecip = .Recipients.Add(address(i))
                objOutlookRecip.Type = olTo
            Next
        End If
        .subject = subject
        If IsNullOrEmpty(signature) Then
            .HTMLBody = StringFormat("<HTML><BODY style='font-family:{0};font-size:{1}'>{2}{3}</BODY></HTML>", font, fontSize, body, IIf(NotIsNullOrEmpty(signature), "<BR/>" & signature, ""))
        Else
            bodyStart = InStr(1, signature, "<body ")
            bodyStart = InStr(bodyStart, signature, ">")

            bodyEnd = InStr(bodyStart, signature, "</body")
            .HTMLBody = StringFormat("{0}{1}{2}{3}", Left(signature, bodyStart), body, "<br/>", Mid(signature, bodyStart + 1))
        End If
        If NotIsNullOrEmpty(replyAddress) Then
            If useReplyAddressAsOnBehalf Then
                If DoesMailboxExist(replyAddress) Then
                    .SentOnBehalfOfName = replyAddress
                End If
            Else
                .ReplyRecipients.Add replyAddress
                Set mailbox = GetMailbox(replyAddress)
                If mailbox Is Nothing Then
                    Set .Sender = mailbox
                End If
            End If
        End If
        .Importance = Importance
        .ReadReceiptRequested = ReadReceipt

        On Error GoTo Mail_NoAttachments
        For i = 0 To omArrayFunctions.StringArrayCount(attachments) - 1
            bodyFilenameMissing = bodyFilenameMissing & attachments(i) & vbCrLf
        Next i
        For i = 0 To omArrayFunctions.StringArrayCount(attachments) - 1
            filename = attachments(i)
            If gFso.FileExists(filename) Then
                .attachments.Add filename
            Else
                MsgBox filename & " is missing.", vbCritical
                'gOpmaatMessaging.Save "accounts1@itecs.nl", "carole@itecs.nl", "File does not exist and did not get attached: " & filename, bodyFilenameMissing, olImportanceHigh
            End If
        Next i
Mail_NoAttachments:
        On Error GoTo 0
        For Each objOutlookRecip In .Recipients
            objOutlookRecip.Resolve
        Next
        If Popup Then
            .Display
        Else
            .Save
            .Send
        End If
    End With
    Set objOutlookMsg = Nothing
    Set objOutlookRecip = Nothing

End Sub
Public Function DoesMailboxExist(Email As String) As Boolean
'Dim objNS As Outlook.NameSpace
'Dim objFolder As Outlook.MAPIFolder 'get name of other persons folder
'Dim objRecip As Outlook.Recipient 'other persons name

Dim objNS As Object
Dim objFolder As Object 'get name of other persons folder
Dim objRecip As Object 'other persons name


On Error Resume Next 'will skip over errors


' ### email address of the Calendar/email etc you want to use ###
'email = "helpdesk@rietje.be"



Set objNS = objOutlook.GetNamespace("MAPI")

Set objRecip = objNS.CreateRecipient(Email)

Set objFolder = objNS.GetSharedDefaultFolder(objRecip, olFolderInbox)

DoesMailboxExist = Not (objFolder Is Nothing)

'Otherwise from this folder you can now make a new email, add contents and send.

Set objNS = Nothing
Set objFolder = Nothing
Set objRecip = Nothing

End Function

Public Function GetMailbox(mailaddress As String) As Object
Dim addrEntry As Object
Dim folder As Object
Dim store As Object
Dim accounts As Object
Dim account As Object
'Dim acc As Outlook.account
    ' Get the store for the current folder.
    Set folder = objOutlook.ActiveExplorer().CurrentFolder 'as Outlook.Folder;
    Set store = folder.store

    Set accounts = objOutlook.Session.accounts

    'Match the delivery store of each account with the
    'store for the current folder.
    For Each account In accounts
     'Set acc = account
        If (account.SmtpAddress = mailaddress) Then

            Set addrEntry = account.CurrentUser.AddressEntry
            Exit For
        End If
    Next
    Set GetMailbox = addrEntry
End Function
Private Sub Class_Initialize()
    Set objOutlook = CreateObject("Outlook.Application")
End Sub

Private Sub Class_Terminate()
    Set objOutlook = Nothing
End Sub
