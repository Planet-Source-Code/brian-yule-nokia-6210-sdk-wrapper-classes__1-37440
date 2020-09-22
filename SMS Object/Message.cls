VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private Const SMS_Picture = 7

Public smsMemory As SMS3ASuiteLib.ISMSMemory
Public MessageIndex As Integer
Public MemoryLocationIndex As Integer

Public Property Get ReadMessageCommandMessageNumber() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
        
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)    'read the message
    
    ReadMessageCommandMessageNumber = smsMessage.CommandMessageNumber
Exit Property
ErrorHandler:
    On Error Resume Next

    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageText() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
        
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)    'read the message
    
    ReadMessageText = smsMessage.UserDataText
Exit Property
ErrorHandler:
    On Error Resume Next

    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageReplyPath() As Long
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
        
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
        
    ReadMessageReplyPath = smsMessage.ReplyPath
Exit Property
ErrorHandler:
    On Error Resume Next

    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageStatusReportRequest() As Long
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
        
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
        
    ReadMessageStatusReportRequest = smsMessage.StatusReportRequest
Exit Property
ErrorHandler:
    On Error Resume Next

    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageSMSCAddress() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
        
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
        
    ReadMessageSMSCAddress = CStr(smsMessage.SCAddress)
Exit Property
ErrorHandler:
    On Error Resume Next

    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageOtherEndAddress() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
       
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
    
    ReadMessageOtherEndAddress = CStr(smsMessage.OtherEndAddress)
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageREPORT_QUALIFIER() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
    
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
    
    Select Case smsMessage.StatusReportQualifier
        Case SMS3ASuiteLib.REPORT_QUALIFIER.COMMAND_REQUESTED: ReadMessageREPORT_QUALIFIER = "COMMAND REQUESTED"
        Case SMS3ASuiteLib.REPORT_QUALIFIER.SUBMIT_REQUESTED:  ReadMessageREPORT_QUALIFIER = "SUBMIT REQUESTED"
    End Select
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageValidityPeriodFormat() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
        
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
        
    Select Case smsMessage.ValidityPeriodFormat
        Case SMS3ASuiteLib.VALIDITY_PERIOD_FORMAT.ABSOLUTE_FORMAT: ReadMessageValidityPeriodFormat = "ABSOLUTE FORMAT"
        Case SMS3ASuiteLib.VALIDITY_PERIOD_FORMAT.ENHANCED_FORMAT: ReadMessageValidityPeriodFormat = "ENHANCED FORMAT"
        Case SMS3ASuiteLib.VALIDITY_PERIOD_FORMAT.NOT_PRESENT: ReadMessageValidityPeriodFormat = "NOT PRESENT"
        Case SMS3ASuiteLib.VALIDITY_PERIOD_FORMAT.RELATIVE_FORMAT: ReadMessageValidityPeriodFormat = "RELATIVE FORMAT"
    End Select
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageDateStamp() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
    
    Dim pYear&, pMonth&, pDay&, pHour&, pMinute&, pSecond&, pTimeZone&
    
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
    
    Call smsMessage.get_SCTimeStamp(pYear, pMonth, pDay, pHour, pMinute, pSecond, pTimeZone)
        
    ReadMessageDateStamp = pDay & "/" & pMonth & "/" & pYear
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageTimeStamp() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
    
    Dim pYear&, pMonth&, pDay&, pHour&, pMinute&, pSecond&, pTimeZone&
    
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
    
    Call smsMessage.get_SCTimeStamp(pYear, pMonth, pDay, pHour, pMinute, pSecond, pTimeZone)
        
    ReadMessageTimeStamp = pHour & ":" & pMinute & ":" & pSecond
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageTimeZone() As Long
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
    
    Dim pYear&, pMonth&, pDay&, pHour&, pMinute&, pSecond&, pTimeZone&
    
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
    
    Call smsMessage.get_SCTimeStamp(pYear, pMonth, pDay, pHour, pMinute, pSecond, pTimeZone)
        
    ReadMessageTimeZone = pTimeZone
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageTimeToGMT() As Long
    On Error GoTo ErrorHandler
    
     Dim smsMessage As SMS3ASuiteLib.ShortMessage
    
     Dim pYear&, pMonth&, pDay&, pHour&, pMinute&, pSecond&, pTimeZone&
     
     Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
     
     Call smsMessage.get_SCTimeStamp(pYear, pMonth, pDay, pHour, pMinute, pSecond, pTimeZone)
         
     ReadMessageTimeToGMT = 15 * pTimeZone / 60
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageMessageType() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
    
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)    'read the message
    
    Select Case smsMessage.MessageType
        Case SMS3ASuiteLib.SMS_MESSAGE_TYPE.COMMAND_MESSAGE: ReadMessageMessageType = "COMMAND MESSAGE"
        Case SMS3ASuiteLib.SMS_MESSAGE_TYPE.DELIVER_MESSAGE: ReadMessageMessageType = "DELIVER MESSAGE"
        Case SMS3ASuiteLib.SMS_MESSAGE_TYPE.MO_UNDEFINED_MESSAGE: ReadMessageMessageType = "MO UNDEFINED MESSAGE"
        Case SMS3ASuiteLib.SMS_MESSAGE_TYPE.MT_UNDEFINED_MESSAGE: ReadMessageMessageType = "MT UNDEFINED MESSAGE"
        Case SMS3ASuiteLib.SMS_MESSAGE_TYPE.STATUS_REPORT_MESSAGE: ReadMessageMessageType = "STATUS REPORT MESSAGE"
        Case SMS3ASuiteLib.SMS_MESSAGE_TYPE.SUBMIT_MESSAGE: ReadMessageMessageType = "SUBMIT MESSAGE"
        Case SMS_Picture: ReadMessageMessageType = "Picture Message"
    End Select
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageStorageStatus() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
    
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
    
    Select Case smsMessage.StorageStatus
        Case SMS3ASuiteLib.STORAGE_STATUS.DELIVERED: ReadMessageStorageStatus = "DELIVERED"
        Case SMS3ASuiteLib.STORAGE_STATUS.NOT_READ_FROM_PHONE: ReadMessageStorageStatus = "NOT READ FROM PHONE"
        Case SMS3ASuiteLib.STORAGE_STATUS.NOT_SENT_FROM_PHONE: ReadMessageStorageStatus = "NOT SENT FROM PHONE"
        Case SMS3ASuiteLib.STORAGE_STATUS.READ_FROM_PHONE: ReadMessageStorageStatus = "READ FROM PHONE"
        Case SMS3ASuiteLib.STORAGE_STATUS.SENT_FROM_PHONE: ReadMessageStorageStatus = "SENT FROM PHONE"
    End Select
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageUserDataFormat() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
    
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
    
    Select Case smsMessage.UserDataFormat
        Case SMS3ASuiteLib.USER_DATA_FORMAT.DATA_8_BIT: ReadMessageUserDataFormat = "DATA 8 BIT"
        Case SMS3ASuiteLib.USER_DATA_FORMAT.DATA_CODING_SCHEME_BASED: ReadMessageUserDataFormat = "DATA CODING_SCHEME BASED"
        Case SMS3ASuiteLib.USER_DATA_FORMAT.DEFAULT_ALPHABET_7_BIT: ReadMessageUserDataFormat = "DEFAULT ALPHABET 7 BIT"
        Case SMS3ASuiteLib.USER_DATA_FORMAT.UNICODE_16_BIT: ReadMessageUserDataFormat = "UNICODE 16 BIT"
    End Select
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageProtocol() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
    
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)   'read the message
    
    Select Case smsMessage.ProtocolIdentifier
        Case SMS_PROTOCOL.SMS_Email: ReadMessageProtocol = "Email"
        Case SMS_PROTOCOL.SMS_ERMES: ReadMessageProtocol = "ERMES"
        Case SMS_PROTOCOL.SMS_Fax: ReadMessageProtocol = "Fax"
        Case SMS_PROTOCOL.SMS_Paging: ReadMessageProtocol = "Paging"
        Case SMS_PROTOCOL.SMS_Text: ReadMessageProtocol = "Text"
        Case SMS_PROTOCOL.SMS_UCI: ReadMessageProtocol = "UCI"
        Case SMS_PROTOCOL.SMS_Voice: ReadMessageProtocol = "Voice"
        Case SMS_PROTOCOL.SMS_X400: ReadMessageProtocol = "X400"
    End Select
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property

Public Property Get ReadMessageUserDataHeader() As String
    On Error GoTo ErrorHandler
    
    Dim smsMessage As SMS3ASuiteLib.ShortMessage
    
    Call smsMemory.Read(MemoryLocationIndex, MessageIndex, smsMessage, MARK_MESSAGE.AS_UNREAD)      'read the message
    
    Select Case smsMessage.ProtocolIdentifier
        Case UserDataHeader.SMS_CallerIDLogo: ReadMessageUserDataHeader = "Caller ID Logo"
        Case UserDataHeader.SMS_ConcatenatedMessages: ReadMessageUserDataHeader = "Concatenated Messages"
        Case UserDataHeader.SMS_EmailMessage: ReadMessageUserDataHeader = "Email Message"
        Case UserDataHeader.SMS_FaxMessage: ReadMessageUserDataHeader = "Fax Message"
        Case UserDataHeader.SMS_MultipartMessage: ReadMessageUserDataHeader = "Multipart Message"
        Case UserDataHeader.SMS_NoUDH: ReadMessageUserDataHeader = "No UDH"
        Case UserDataHeader.SMS_OpLogo: ReadMessageUserDataHeader = "Op Logo"
        Case UserDataHeader.SMS_OtherMessage: ReadMessageUserDataHeader = "Other Message"
        Case UserDataHeader.SMS_Ringtone: ReadMessageUserDataHeader = "Ringtone"
        Case UserDataHeader.SMS_UnknownUDH: ReadMessageUserDataHeader = "Unknown UDH"
        Case UserDataHeader.SMS_VoiceMessage: ReadMessageUserDataHeader = "Voice Message"
        Case UserDataHeader.SMS_WAPvCalendar: ReadMessageUserDataHeader = "WAP v Calendar"
        Case UserDataHeader.SMS_WAPvCalendarSecure: ReadMessageUserDataHeader = "WAP v Calendar Secure"
        Case UserDataHeader.SMS_WAPvCard: ReadMessageUserDataHeader = "WAP v Card"
        Case UserDataHeader.SMS_WAPvCardSecure: ReadMessageUserDataHeader = "WAP v CardSecure"
    End Select
Exit Property
ErrorHandler:
    On Error Resume Next

    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsMemory.GetLastError, modPhone.getSMSError(smsMemory.GetLastError)
End Property
