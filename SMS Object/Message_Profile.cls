VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Message_Profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public smsSettings As SMS3ASuiteLib.ISMSSettings
Public ParameterSetIndex As Integer

Public Property Get SMSC_Number() As String
    On Error GoTo ErrorHandler
    
    Dim pParameterSetName As String, pProtocolIdentifier As Byte, pDataCodingScheme As Byte, pValidityPeriod As Byte, pSCAddress As String
            
    Call smsSettings.GetSMSParametersSet(ParameterSetIndex, pParameterSetName, pProtocolIdentifier, pDataCodingScheme, pValidityPeriod, pSCAddress)
        
    SMSC_Number = pSCAddress
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsSettings.GetLastError, modPhone.getSMSError(smsSettings.GetLastError)
End Property

Public Property Get SMSC_Name() As String
    On Error GoTo ErrorHandler
    
    Dim pParameterSetName As String, pProtocolIdentifier As Byte, pDataCodingScheme As Byte, pValidityPeriod As Byte, pSCAddress As String
            
    Call smsSettings.GetSMSParametersSet(ParameterSetIndex, pParameterSetName, pProtocolIdentifier, pDataCodingScheme, pValidityPeriod, pSCAddress)
        
    SMSC_Name = pParameterSetName
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsSettings.GetLastError, modPhone.getSMSError(smsSettings.GetLastError)
End Property

Public Property Get SMSC_Protocol() As String
    On Error GoTo ErrorHandler
    
    Dim pParameterSetName As String, pProtocolIdentifier As Byte, pDataCodingScheme As Byte, pValidityPeriod As Byte, pSCAddress As String
            
    Call smsSettings.GetSMSParametersSet(ParameterSetIndex, pParameterSetName, pProtocolIdentifier, pDataCodingScheme, pValidityPeriod, pSCAddress)
        
    Select Case pProtocolIdentifier
        Case SMS_PROTOCOL.SMS_Email: SMSC_Protocol = "Email"
        Case SMS_PROTOCOL.SMS_ERMES: SMSC_Protocol = "ERMES"
        Case SMS_PROTOCOL.SMS_Fax: SMSC_Protocol = "Fax"
        Case SMS_PROTOCOL.SMS_Paging: SMSC_Protocol = "Paging"
        Case SMS_PROTOCOL.SMS_Text: SMSC_Protocol = "Text"
        Case SMS_PROTOCOL.SMS_UCI: SMSC_Protocol = "UCI"
        Case SMS_PROTOCOL.SMS_Voice: SMSC_Protocol = "Voice"
        Case SMS_PROTOCOL.SMS_X400: SMSC_Protocol = "X400"
    End Select
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsSettings.GetLastError, modPhone.getSMSError(smsSettings.GetLastError)
End Property

Public Property Get SMSC_Validity() As String
    On Error GoTo ErrorHandler
    
    Dim pParameterSetName As String, pProtocolIdentifier As Byte, pDataCodingScheme As Byte, pValidityPeriod As Byte, pSCAddress As String
        
    Call smsSettings.GetSMSParametersSet(ParameterSetIndex, pParameterSetName, pProtocolIdentifier, pDataCodingScheme, pValidityPeriod, pSCAddress)
 
    Select Case pValidityPeriod
        Case SMS_ValidityLength.SMS_MAX_TIME: SMSC_Validity = "Max. time"
        Case SMS_ValidityLength.SMS_ONE_WEEK: SMSC_Validity = "One week"
        Case SMS_ValidityLength.SMS_THREE_DAYS: SMSC_Validity = "Three days"
        Case SMS_ValidityLength.SMS_24_HOURS: SMSC_Validity = "24 hours"
        Case SMS_ValidityLength.SMS_SIX_HOURS: SMSC_Validity = "Six hours"
        Case SMS_ValidityLength.SMS_ONE_HOUR: SMSC_Validity = "One hour"
    End Select
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsSettings.GetLastError, modPhone.getSMSError(smsSettings.GetLastError)
End Property

Public Property Get SMSC_Coding() As String
    On Error GoTo ErrorHandler
    
    Dim pParameterSetName As String, pProtocolIdentifier As Byte, pDataCodingScheme As Byte, pValidityPeriod As Byte, pSCAddress As String
        
    Call smsSettings.GetSMSParametersSet(ParameterSetIndex, pParameterSetName, pProtocolIdentifier, pDataCodingScheme, pValidityPeriod, pSCAddress)
        
    Select Case pDataCodingScheme
        Case SMS_DataType.SMS_BitmapData: SMSC_Coding = "Bitmap Data"
        Case SMS_DataType.SMS_NoData: SMSC_Coding = "No Data"
        Case SMS_DataType.SMS_OtherData: SMSC_Coding = "Other Data"
        Case SMS_DataType.SMS_PlainText: SMSC_Coding = "Plain Text"
        Case SMS_DataType.SMS_RingtoneData: SMSC_Coding = "Ringtone Data"
    End Select
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print smsSettings.GetLastError, modPhone.getSMSError(smsSettings.GetLastError)
End Property
