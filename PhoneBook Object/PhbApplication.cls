VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Contacts"
Attribute VB_Ext_KEY = "Member1" ,"Memory_PHB"
Attribute VB_Ext_KEY = "Member2" ,"Version"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
'                   Phone Book Class
'
'   Created by:         Brian Yule
'   Last Modified:      21 Feb 2002
'   Last Modified by:   Brian Yule
'
'
Option Explicit

Private WithEvents PHB As PhonebookAdapterDS3.PhonebookSuite3
Attribute PHB.VB_VarHelpID = -1

Private mvarVersionInfo As Version
Private mvarPhoneMemory As Memory_PHB
Private mvarSimMemory As Memory_PHB
Private mvarPhoneContacts As Contacts
Private mvarSimContacts As Contacts

Public Event Added()
Public Event Deleted()
Public Event Modified()

Public Property Get VersionInfo() As Version
    If mvarVersionInfo Is Nothing Then Set mvarVersionInfo = New Version
    Set VersionInfo = mvarVersionInfo
End Property

Public Property Get phoneMemory() As Memory_PHB
    If mvarPhoneMemory Is Nothing Then Set mvarPhoneMemory = New Memory_PHB
    Set phoneMemory = mvarPhoneMemory
End Property

Public Property Get simMemory() As Memory_PHB
    If mvarSimMemory Is Nothing Then Set mvarSimMemory = New Memory_PHB
    Set simMemory = mvarSimMemory
End Property

Public Property Get PhoneContacts() As Contacts
    If mvarPhoneContacts Is Nothing Then Set mvarPhoneContacts = New Contacts
    Set PhoneContacts = mvarPhoneContacts
End Property

Public Property Get SimContacts() As Contacts
    If mvarSimContacts Is Nothing Then Set mvarSimContacts = New Contacts
    Set SimContacts = mvarSimContacts
End Property

Private Sub PHB_ContactAdded(ByVal PhysicalMemory As PhonebookAdapterDS3.PhysicalMem, ByVal LogicalMemory As PhonebookAdapterDS3.LogicalMem, ByVal Location As Long)
    RaiseEvent Added
End Sub

Private Sub PHB_ContactDeleted(ByVal PhysicalMemory As PhonebookAdapterDS3.PhysicalMem, ByVal LogicalMemory As PhonebookAdapterDS3.LogicalMem, ByVal Location As Long)
    RaiseEvent Deleted
End Sub

Private Sub PHB_ContactModified(ByVal PhysicalMemory As PhonebookAdapterDS3.PhysicalMem, ByVal LogicalMemory As PhonebookAdapterDS3.LogicalMem, ByVal Location As Long)
    RaiseEvent Modified
End Sub

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set PHB = New PhonebookAdapterDS3.PhonebookSuite3
    
    Set mvarVersionInfo = New Version
    Set mvarVersionInfo.phb2 = PHB
    
    Set mvarPhoneMemory = New Memory_PHB
    mvarPhoneMemory.MemoryLocationIndex = PhonebookAdapterDS3.PhonebookMemory.MEMORY_ME
    Set mvarPhoneMemory.phb2 = PHB
    Set mvarSimMemory = New Memory_PHB
    mvarSimMemory.MemoryLocationIndex = PhonebookAdapterDS3.PhonebookMemory.MEMORY_SIM
    Set mvarSimMemory.phb2 = PHB
    
    Dim i As Integer
    
    Set mvarPhoneContacts = New Contacts
    If mvarPhoneMemory.Full > 0 Then
        For i = 1 To mvarPhoneMemory.Full
            Dim tmpPhoneContact As New Contact
            Set tmpPhoneContact.PHB = PHB
            tmpPhoneContact.MemoryLocationIndex = PhonebookAdapterDS3.PhonebookMemory.MEMORY_ME
            tmpPhoneContact.ContactIndex = i
            Call mvarPhoneContacts.Add(tmpPhoneContact, i)
            Set tmpPhoneContact = Nothing
        Next i
    End If
    Set mvarSimContacts = New Contacts
    If mvarSimMemory.Full > 0 Then
        For i = 1 To mvarSimMemory.Full
            Dim tmpSimContact As New Contact
            Set tmpSimContact.PHB = PHB
            tmpSimContact.MemoryLocationIndex = PhonebookAdapterDS3.PhonebookMemory.MEMORY_SIM
            tmpSimContact.ContactIndex = i
            Call mvarSimContacts.Add(tmpSimContact, i)
            Set tmpSimContact = Nothing
        Next i
    End If
Exit Sub
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Sub

Public Property Get GetIMEI() As String
    On Error GoTo ErrorHandler
    
    Dim ppsIMEI As String, ppsType As String, ppsVerSW As String, ppsVerHW As String
    Dim phb2 As PhonebookAdapterDS3.IPhonebook2
    
    Set phb2 = PHB
    Call phb2.GetDevIdentificationInfo(ppsIMEI, ppsType, ppsVerSW, ppsVerHW)
    Set phb2 = Nothing
    
    GetIMEI = ppsIMEI
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Property

Public Property Get GetType() As String
    On Error GoTo ErrorHandler
    
    Dim ppsIMEI As String, ppsType As String, ppsVerSW As String, ppsVerHW As String
    Dim phb2 As PhonebookAdapterDS3.IPhonebook2
    
    Set phb2 = PHB
    Call phb2.GetDevIdentificationInfo(ppsIMEI, ppsType, ppsVerSW, ppsVerHW)
    Set phb2 = Nothing
    
    GetType = ppsType
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Property



Public Property Get PhonebookEditable() As Boolean
    On Error GoTo ErrorHandler
    
    Dim pPhonebookEditable As Boolean, pCallergroupsSupported As Boolean, pSpeedDialAccess As Boolean

    Call PHB.GetSupportCaps(pPhonebookEditable, pCallergroupsSupported, pSpeedDialAccess)
    
    PhonebookEditable = pPhonebookEditable
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Property

Public Property Get SpeedDialAccess() As Boolean
    On Error GoTo ErrorHandler
    
    Dim pPhonebookEditable As Boolean, pCallergroupsSupported As Boolean, pSpeedDialAccess As Boolean

    Call PHB.GetSupportCaps(pPhonebookEditable, pCallergroupsSupported, pSpeedDialAccess)
    
    SpeedDialAccess = pSpeedDialAccess
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Property

Public Property Get CallerGroupsSupported() As Boolean
    On Error GoTo ErrorHandler
    
    Dim pPhonebookEditable As Boolean, pCallergroupsSupported As Boolean, pSpeedDialAccess As Boolean

    Call PHB.GetSupportCaps(pPhonebookEditable, pCallergroupsSupported, pSpeedDialAccess)
    
    CallerGroupsSupported = pCallergroupsSupported
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Property

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Set mvarVersionInfo = Nothing
    Set mvarPhoneMemory = Nothing
    Set mvarSimMemory = Nothing
    Set mvarPhoneContacts = Nothing
    Set mvarSimContacts = Nothing
    Set PHB = Nothing
Exit Sub
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Sub
