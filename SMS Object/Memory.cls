VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Memory_SMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public smsMemory As SMS3ASuiteLib.ISMSMemory

Public Property Get NumberOfParameterSets() As Long
    On Error GoTo ErrorHandler
    
    Dim tmpNumberOfParameterSets As Long
    
    Call smsMemory.GetCapasityOfParamSets(tmpNumberOfParameterSets)
    
    NumberOfParameterSets = tmpNumberOfParameterSets
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get SIMCapacity() As Long
    On Error GoTo ErrorHandler
    
    Dim pSIMSMSMemorySize As Long
    
    Call smsMemory.GetCapasityInSIM(pSIMSMSMemorySize)
    
    SIMCapacity = pSIMSMSMemorySize
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get PhoneCapacity() As Long
    On Error GoTo ErrorHandler
    
    Dim pMESMSMemorySize As Long
    
    Call smsMemory.GetCapasityInME(pMESMSMemorySize)
    
    PhoneCapacity = pMESMSMemorySize
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get Phone_Messages() As Long
    On Error GoTo ErrorHandler
    
    Dim pAmountOfMessagesInMESMSMemory As Long
    
    Call smsMemory.GetNumberOfMessagesInME(pAmountOfMessagesInMESMSMemory)
    
    Phone_Messages = pAmountOfMessagesInMESMSMemory
Exit Property
ErrorHandler:
    On Error Resume Next

    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get SIM_Messages() As Long
    On Error GoTo ErrorHandler
    
    Dim pAmountOfMessagesInSIMSMSMemory As Long
    
    Call smsMemory.GetNumberOfMessagesInSIM(pAmountOfMessagesInSIMSMSMemory)
    
    SIM_Messages = pAmountOfMessagesInSIMSMSMemory
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get PHONE_Unread_Messages() As Long
    On Error GoTo ErrorHandler
    
    Dim pAmountOfUnreadMessageInMESMSMemory As Long
    
    Call smsMemory.GetNumberOfUnreadInME(pAmountOfUnreadMessageInMESMSMemory)
    
    PHONE_Unread_Messages = pAmountOfUnreadMessageInMESMSMemory
Exit Property
ErrorHandler:
    On Error Resume Next

    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get SIM_Unread_Messages() As Long
    On Error GoTo ErrorHandler
    
    Dim pAmountOfUnreadMessagesInSIMSMSMemory As Long
    
    Call smsMemory.GetNumberOfUnreadInSIM(pAmountOfUnreadMessagesInSIMSMSMemory)
    
    SIM_Unread_Messages = pAmountOfUnreadMessagesInSIMSMSMemory
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get Memory_Configuration() As Long
    On Error GoTo ErrorHandler
    
    Dim pMemoryStatus As Long
    
    Call smsMemory.GetMemoryConfiguration(pMemoryStatus)
    
    Memory_Configuration = pMemoryStatus
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

