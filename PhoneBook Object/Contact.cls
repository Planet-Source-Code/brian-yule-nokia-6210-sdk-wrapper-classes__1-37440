VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Contact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public PHB As PhonebookAdapterDS3.PhonebookSuite3
Public ContactIndex As Integer
Public MemoryLocationIndex As Integer

Public Property Get ContactMemory() As Long
    On Error GoTo ErrorHandler
     
    Dim Contact As PhonebookAdapterDS3.IContact
    
    Set Contact = PHB.GetContact(MemoryLocationIndex, ContactIndex)
    
    ContactMemory = Contact.Memory

    Set Contact = Nothing
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Property

Public Property Get ContactGroupId() As Long
    On Error GoTo ErrorHandler
     
    Dim Contact As PhonebookAdapterDS3.IContact
    
    Set Contact = PHB.GetContact(MemoryLocationIndex, ContactIndex)
    
    ContactGroupId = Contact.GroupID

    Set Contact = Nothing
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Property

Public Property Get ContactNumber() As String
    On Error GoTo ErrorHandler
     
    Dim Contact As PhonebookAdapterDS3.IContact
    
    Set Contact = PHB.GetContact(MemoryLocationIndex, ContactIndex)
    
    ContactNumber = Contact.Number
    
    Set Contact = Nothing
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Property

Public Property Get ContactTag(index As Integer) As Long
    On Error GoTo ErrorHandler
     
    Dim Contact As PhonebookAdapterDS3.IContact
        
    Set Contact = PHB.GetContact(MemoryLocationIndex, ContactIndex)
        
    ContactTag = Contact.Tag
    
    Set Contact = Nothing
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Property

Public Property Get ContactName() As String
    On Error GoTo ErrorHandler
     
    Dim Contact As PhonebookAdapterDS3.IContact
    
    Set Contact = PHB.GetContact(MemoryLocationIndex, ContactIndex)
    
    ContactName = Contact.Name
    
    Set Contact = Nothing
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print PHB.GetLastError, getPHBError(PHB.GetLastError)
End Property

