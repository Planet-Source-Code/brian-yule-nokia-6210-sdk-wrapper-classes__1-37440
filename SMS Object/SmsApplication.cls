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
Attribute VB_Ext_KEY = "Member0" ,"Memory_SMS"
Attribute VB_Ext_KEY = "Member1" ,"Message_Profiles"
Attribute VB_Ext_KEY = "Member2" ,"Messages"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
'                   SMS Messaging Class
'
'   Created by:         Brian Yule
'   Last Modified:      21 Feb 2002
'   Last Modified by:   Brian Yule
'
'
Option Explicit

Private WithEvents SMS As SMS3ASuiteLib.SMS_SuiteAdapter
Attribute SMS.VB_VarHelpID = -1

Private mvarMessageProfiles As Message_Profiles
Private mvarSIMMessages As Messages
Private mvarPhoneMessages As Messages
Private mvarMemoryInfo As Memory_SMS

Public Event Incomming()
Public Event Sent()
Public Event MaxCapacity()

Public Property Get MessageProfiles() As Message_Profiles
    If mvarMessageProfiles Is Nothing Then Set mvarMessageProfiles = New Message_Profiles
    Set MessageProfiles = mvarMessageProfiles
End Property

Public Property Get SIMMessages() As Messages
    If mvarSIMMessages Is Nothing Then Set mvarSIMMessages = New Messages
    Set SIMMessages = mvarSIMMessages
End Property

Public Property Get PhoneMessages() As Messages
    If mvarPhoneMessages Is Nothing Then Set mvarPhoneMessages = New Messages
    Set PhoneMessages = mvarPhoneMessages
End Property

Public Property Get MemoryInfo() As Memory_SMS
    If mvarMemoryInfo Is Nothing Then Set mvarMemoryInfo = New Memory_SMS
    Set MemoryInfo = mvarMemoryInfo
End Property

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    
    Set SMS = New SMS3ASuiteLib.SMS_SuiteAdapter
    
    Call SMS.StartListeningEvents
    
    Set mvarMemoryInfo = New Memory_SMS
    Set mvarMemoryInfo.smsMemory = SMS
    
    Dim i As Integer
    
    ' Message Profile
    Set mvarMessageProfiles = New Message_Profiles
    For i = 1 To mvarMemoryInfo.NumberOfParameterSets
        Call mvarMessageProfiles.Add(New Message_Profile, i)
        Set mvarMessageProfiles.Item(i).smsSettings = SMS
        mvarMessageProfiles.Item(i).ParameterSetIndex = i
    Next i
    ' Sim Messages
    Set mvarSIMMessages = New Messages
    If mvarMemoryInfo.SIM_Messages > 0 Then
        For i = 1 To mvarMemoryInfo.SIM_Messages
            Dim tmpSimMessage As New Message
            
            Set tmpSimMessage.smsMemory = SMS
            
            tmpSimMessage.MessageIndex = i
            tmpSimMessage.MemoryLocationIndex = SMS3ASuiteLib.SMS_MEMORY_LOCATION.SIM_MEMORY
            
            Call mvarSIMMessages.Add(tmpSimMessage, i)
            
            Set tmpSimMessage = Nothing
        Next i
    End If
    ' Phone Messages
    Set mvarPhoneMessages = New Messages
    If mvarMemoryInfo.Phone_Messages > 0 Then
        For i = 1 To mvarMemoryInfo.Phone_Messages
            Dim tmpPhoneMessage As New Message
            
            Set tmpPhoneMessage.smsMemory = SMS
            
            tmpPhoneMessage.MessageIndex = i
            tmpPhoneMessage.MemoryLocationIndex = SMS3ASuiteLib.SMS_MEMORY_LOCATION.PHONE_MEMORY
            
            Call mvarPhoneMessages.Add(tmpPhoneMessage, i)
            
            Set tmpPhoneMessage = Nothing
        Next i
    End If
Exit Sub
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print SMS.GetLastError, modPhone.getSMSError(SMS.GetLastError)
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrorHandler
    
    Call SMS.Terminate
    
    Set SMS = Nothing
Exit Sub
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print SMS.GetLastError, modPhone.getSMSError(SMS.GetLastError)
End Sub

Private Sub SMS_ShortMsgReceived(ByVal smsMemory As SMS3ASuiteLib.SMS_MEMORY_LOCATION, ByVal index As Long, ByVal pSMS As SMS3ASuiteLib.IShortMessage)
    On Error GoTo ErrorHandler
          
    Dim tmpMessage As New Message
    
    Set tmpMessage.smsMemory = SMS
    tmpMessage.MessageIndex = index
    tmpMessage.MemoryLocationIndex = smsMemory
       
    Select Case smsMemory
        Case SMS3ASuiteLib.SMS_MEMORY_LOCATION.PHONE_MEMORY: Call PhoneMessages.Add(tmpMessage, PhoneMessages.Count)
        Case SMS3ASuiteLib.SMS_MEMORY_LOCATION.SIM_MEMORY: Call SIMMessages.Add(tmpMessage, SIMMessages.Count)
    End Select
    RaiseEvent Incomming
Exit Sub
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print SMS.GetLastError, modPhone.getSMSError(SMS.GetLastError)
End Sub

Private Sub SMS_ShortMsgSent(ByVal pSMS As SMS3ASuiteLib.IShortMessage)
    On Error GoTo ErrorHandler
    
    RaiseEvent Sent
Exit Sub
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print SMS.GetLastError, modPhone.getSMSError(SMS.GetLastError)
End Sub

Private Sub SMS_SMSMemoryFull(ByVal smsMemory As SMS3ASuiteLib.SMS_MEMORY_LOCATION)
    On Error GoTo ErrorHandler
    
    RaiseEvent MaxCapacity
Exit Sub
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print SMS.GetLastError, modPhone.getSMSError(SMS.GetLastError)
End Sub

Public Function SendMessage(MessageNumber As String, MessageText As String, ParameterSetIndex As Integer)
    On Error GoTo ErrorHandler
    
    Dim smsSettings As SMS3ASuiteLib.ISMSSettings, smsMessage As SMS3ASuiteLib.ShortMessage
    Dim pParameterSetName As String, pProtocolIdentifier As Byte, pDataCodingScheme As Byte, pValidityPeriod As Byte, pSCAddress As String
    
    Set smsSettings = SMS
    
    Call smsSettings.GetSMSParametersSet(ParameterSetIndex, pParameterSetName, pProtocolIdentifier, pDataCodingScheme, pValidityPeriod, pSCAddress)
    
    Set smsSettings = Nothing

    Set smsMessage = SMS.CreateShortMsg
    
    Let smsMessage.SCAddress = pSCAddress
    Let smsMessage.ProtocolIdentifier = pProtocolIdentifier
    Let smsMessage.DataCodingScheme = pDataCodingScheme
    Let smsMessage.ValidityPeriodFormat = SMS3ASuiteLib.VALIDITY_PERIOD_FORMAT.RELATIVE_FORMAT
    Let smsMessage.UserDataFormat = SMS3ASuiteLib.USER_DATA_FORMAT.DEFAULT_ALPHABET_7_BIT  'user data format
    Let smsMessage.UserDataText = MessageText
    Let smsMessage.OtherEndAddress = MessageNumber

    Call SMS.Send(smsMessage)

    Set smsMessage = Nothing
Exit Function
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print SMS.GetLastError, modPhone.getSMSError(SMS.GetLastError)
End Function

Public Sub DeletePhoneMessage(index As Integer)
    On Error GoTo ErrorHandler
    
    Dim smsMemory As SMS3ASuiteLib.ISMSMemory
    
    Set smsMemory = SMS
    Call smsMemory.Delete(SMS3ASuiteLib.SMS_MEMORY_LOCATION.PHONE_MEMORY, index)   ' Delete the message
    Set smsMemory = Nothing
        
    Call PhoneMessages.Remove(index)
Exit Sub
ErrorHandler:
    On Error Resume Next

    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print SMS.GetLastError, modPhone.getSMSError(SMS.GetLastError)
End Sub

Public Sub DeleteSIMMessage(index As Integer)
    On Error GoTo ErrorHandler
    
    Dim smsMemory As SMS3ASuiteLib.ISMSMemory
    
    Set smsMemory = SMS
    Call smsMemory.Delete(SMS3ASuiteLib.SMS_MEMORY_LOCATION.SIM_MEMORY, index)    ' Delete the message
    Set smsMemory = Nothing
    
    Call SIMMessages.Remove(index)
Exit Sub
ErrorHandler:
    On Error Resume Next

    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print SMS.GetLastError, modPhone.getSMSError(SMS.GetLastError)
End Sub

Public Function IsValidNumber(ByVal Number As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tmpNumber As String
    
    tmpNumber = Number
    
    If tmpNumber = "" Then
        IsValidNumber = False
    Else
        IsValidNumber = True
        Dim i As Integer, Found As Boolean
        For i = 1 To Len(tmpNumber)
            Select Case Mid(tmpNumber, i, 1)
                Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "(", ")", "+", "-"
                Case Else:
                    Found = True
                    Exit For
            End Select
        Next i
        If Found Then IsValidNumber = False
    End If
Exit Function
ErrorHandler:
    On Error Resume Next

    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print SMS.GetLastError, modPhone.getSMSError(SMS.GetLastError)
End Function
