Attribute VB_Name = "modPhone"
Option Explicit

Public Enum UserDataHeader
    SMS_NoUDH = 0
    SMS_ConcatenatedMessages = 1
    SMS_Ringtone = 2
    SMS_OpLogo = 3
    SMS_CallerIDLogo = 4
    SMS_MultipartMessage = 5
    SMS_WAPvCard = 6
    SMS_WAPvCalendar = 7
    SMS_WAPvCardSecure = 8
    SMS_WAPvCalendarSecure = 9
    SMS_VoiceMessage = 10
    SMS_FaxMessage = 11
    SMS_EmailMessage = 12
    SMS_OtherMessage = 13
    SMS_UnknownUDH = 14
End Enum

Public Enum SMS_PROTOCOL
    SMS_Text = &H0
    SMS_Fax = &H22
    SMS_Voice = &H24
    SMS_ERMES = &H25
    SMS_Paging = &H26
    SMS_X400 = &H29
    SMS_UCI = &H2D
    SMS_Email = &H32
End Enum

Public Enum SMS_DataType
    SMS_NoData = 0
    SMS_PlainText = 1
    SMS_BitmapData = 2
    SMS_RingtoneData = 3
    SMS_OtherData = 4
End Enum

'SMS message validity length identifiers
Public Enum SMS_ValidityLength
    SMS_MAX_TIME = &HFF
    SMS_ONE_WEEK = &HAD
    SMS_THREE_DAYS = &HA9
    SMS_24_HOURS = &HA7
    SMS_SIX_HOURS = &H47
    SMS_ONE_HOUR = &HB
End Enum

'SMS message consts
Public Enum MARK_MESSAGE
    AS_READ = 1
    AS_UNREAD = 0
End Enum

Public Function getSMSError(index As Long) As String
    On Error Resume Next
    
    Select Case index
        Case SMS3ASuiteLib.NmpAdapterError.errBarred: getSMSError = "errBarred"
        Case SMS3ASuiteLib.NmpAdapterError.errCalendarCallEmpty: getSMSError = "errCalendarCallEmpty"
        Case SMS3ASuiteLib.NmpAdapterError.errCalendarComponentCreation: getSMSError = "errCalendarComponentCreation"
        Case SMS3ASuiteLib.NmpAdapterError.errCalendarEmpty: getSMSError = "errCalendarEmpty"
        Case SMS3ASuiteLib.NmpAdapterError.errCalendarItemDelete: getSMSError = "errCalendarItemDelete"
        Case SMS3ASuiteLib.NmpAdapterError.errCalendarItemRead: getSMSError = "errCalendarItemRead"
        Case SMS3ASuiteLib.NmpAdapterError.errCalendarItemWrite: getSMSError = "errCalendarItemWrite"
        Case SMS3ASuiteLib.NmpAdapterError.errCalendarNoDelete: getSMSError = "errCalendarNoDelete"
        Case SMS3ASuiteLib.NmpAdapterError.errCalendarNoMoreNotes: getSMSError = "errCalendarNoMoreNotes"
        Case SMS3ASuiteLib.NmpAdapterError.errCalendarNotSupported: getSMSError = "errCalendarNotSupported"
        Case SMS3ASuiteLib.NmpAdapterError.errCalendarUnknownItemType: getSMSError = "errCalendarUnknownItemType"
        Case SMS3ASuiteLib.NmpAdapterError.errCalendarUnknownNoteType: getSMSError = "errCalendarUnknownNoteType"
        Case SMS3ASuiteLib.NmpAdapterError.errCallAlreadyActive: getSMSError = "errCallAlreadyActive"
        Case SMS3ASuiteLib.NmpAdapterError.errCallCostLimitReached: getSMSError = "errCallCostLimitReached"
        Case SMS3ASuiteLib.NmpAdapterError.errCallInvalidMode: getSMSError = "errCallInvalidMode"
        Case SMS3ASuiteLib.NmpAdapterError.errCallNoActiveCall: getSMSError = "errCallNoActiveCall"
        Case SMS3ASuiteLib.NmpAdapterError.errCallNoDualModeCall: getSMSError = "errCallNoDualModeCall"
        Case SMS3ASuiteLib.NmpAdapterError.errCallSignallingFailure: getSMSError = "errCallSignallingFailure"
        Case SMS3ASuiteLib.NmpAdapterError.errCbInvalidLanguage: getSMSError = "errCbInvalidLanguage"
        Case SMS3ASuiteLib.NmpAdapterError.errCbInvalidTopic: getSMSError = "errCbInvalidTopic"
        Case SMS3ASuiteLib.NmpAdapterError.errCbSettingSetFailed: getSMSError = "errCbSettingSetFailed"
        Case SMS3ASuiteLib.NmpAdapterError.errCbToomanyLang: getSMSError = "errCbToomanyLang"
        Case SMS3ASuiteLib.NmpAdapterError.errCbToomanyTopic: getSMSError = "errCbToomanyTopic"
        Case SMS3ASuiteLib.NmpAdapterError.errCommunicationError: getSMSError = "errCommunicationError"
        Case SMS3ASuiteLib.NmpAdapterError.errDataNotAvail: getSMSError = "errDataNotAvail"
        Case SMS3ASuiteLib.NmpAdapterError.errDbXXXX: getSMSError = "errDbXXXX"
        Case SMS3ASuiteLib.NmpAdapterError.errDeviceFailure: getSMSError = "errDeviceFailure"
        Case SMS3ASuiteLib.NmpAdapterError.errEmptyLocation: getSMSError = "errEmptyLocation"
        Case SMS3ASuiteLib.NmpAdapterError.errInvalidLocation: getSMSError = "errInvalidLocation"
        Case SMS3ASuiteLib.NmpAdapterError.errInvalidNumber: getSMSError = "errInvalidNumber"
        Case SMS3ASuiteLib.NmpAdapterError.errInvalidParameter: getSMSError = "errInvalidParameter"
        Case SMS3ASuiteLib.NmpAdapterError.errMemoryFull: getSMSError = "errMemoryFull"
        Case SMS3ASuiteLib.NmpAdapterError.errMpApiNotAvail: getSMSError = "errMpApiNotAvail"
        Case SMS3ASuiteLib.NmpAdapterError.errNetAccessDenied: getSMSError = "errNetAccessDenied"
        Case SMS3ASuiteLib.NmpAdapterError.errNetCallActive: getSMSError = "errNetCallActive"
        Case SMS3ASuiteLib.NmpAdapterError.errNetNoMsgToQuit: getSMSError = "errNetNoMsgToQuit"
        Case SMS3ASuiteLib.NmpAdapterError.errNetUnableToQuit: getSMSError = "errNetUnableToQuit"
        Case SMS3ASuiteLib.NmpAdapterError.errNoError: getSMSError = "errNoError"
        Case SMS3ASuiteLib.NmpAdapterError.errNoSim: getSMSError = "errNoSim"
        Case SMS3ASuiteLib.NmpAdapterError.errNotSupported: getSMSError = "errNotSupported"
        Case SMS3ASuiteLib.NmpAdapterError.errPasswordNotRequired: getSMSError = "errPasswordNotRequired"
        Case SMS3ASuiteLib.NmpAdapterError.errPhoneNotConnected: getSMSError = "errPhoneNotConnected"
        Case SMS3ASuiteLib.NmpAdapterError.errPin2Required: getSMSError = "errPin2Required"
        Case SMS3ASuiteLib.NmpAdapterError.errPinRequired: getSMSError = "errPinRequired"
        Case SMS3ASuiteLib.NmpAdapterError.errPnCallergroupsNotsupported: getSMSError = "errPnCallergroupsNotsupported"
        Case SMS3ASuiteLib.NmpAdapterError.errPnEmpty: getSMSError = "errPnEmpty"
        Case SMS3ASuiteLib.NmpAdapterError.errPnEntryLocked: getSMSError = "errPnEntryLocked"
        Case SMS3ASuiteLib.NmpAdapterError.errPnInvalidCharacters: getSMSError = "errPnInvalidCharacters"
        Case SMS3ASuiteLib.NmpAdapterError.errPnInvalidIconFormat: getSMSError = "errPnInvalidIconFormat"
        Case SMS3ASuiteLib.NmpAdapterError.errPnInvalidMemory: getSMSError = "errPnInvalidMemory"
        Case SMS3ASuiteLib.NmpAdapterError.errPnMemoryFull: getSMSError = "errPnMemoryFull"
        Case SMS3ASuiteLib.NmpAdapterError.errPnNameTooLong: getSMSError = "errPnNameTooLong"
        Case SMS3ASuiteLib.NmpAdapterError.errPnNotAvail: getSMSError = "errPnNotAvail"
        Case SMS3ASuiteLib.NmpAdapterError.errPnNumberTooLong: getSMSError = "errPnNumberTooLong"
        Case SMS3ASuiteLib.NmpAdapterError.errPnSpeedkeyNotassigned: getSMSError = "errPnSpeedkeyNotassigned"
        Case SMS3ASuiteLib.NmpAdapterError.errPnTimestampNotavail: getSMSError = "errPnTimestampNotavail"
        Case SMS3ASuiteLib.NmpAdapterError.errProtocolError: getSMSError = "errProtocolError"
        Case SMS3ASuiteLib.NmpAdapterError.errPuk2Required: getSMSError = "errPuk2Required"
        Case SMS3ASuiteLib.NmpAdapterError.errPukRequired: getSMSError = "errPukRequired"
        Case SMS3ASuiteLib.NmpAdapterError.errReserved: getSMSError = "errReserved"
        Case SMS3ASuiteLib.NmpAdapterError.errRLP: getSMSError = "errRLP"
        Case SMS3ASuiteLib.NmpAdapterError.errSecurityCodeRequired: getSMSError = "errSecurityCodeRequired"
        Case SMS3ASuiteLib.NmpAdapterError.errSignallingFailure: getSMSError = "errSignallingFailure"
        Case SMS3ASuiteLib.NmpAdapterError.errSIMRejected: getSMSError = "errSIMRejected"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsCannotSendMTMessages: getSMSError = "errSmsCannotSendMTMessages"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsCreateFailed: getSMSError = "errSmsCreateFailed"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsDefaultSetNameUsed: getSMSError = "errSmsDefaultSetNameUsed"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsInvalidDataCodingScheme: getSMSError = "errSmsInvalidDataCodingScheme"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsInvalidParameterSetIndex: getSMSError = "errSmsInvalidParameterSetIndex"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsInvalidSCTimeStamp: getSMSError = "errSmsInvalidSCTimeStamp"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsInvalidType: getSMSError = "errSmsInvalidType"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsInvalidUserData: getSMSError = "errSmsInvalidUserData"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsInvalidUserDataFormat: getSMSError = "errSmsInvalidUserDataFormat"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsInvalidUserDataHeaderLength: getSMSError = "errSmsInvalidUserDataHeaderLength"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsInvalidUserDataLength: getSMSError = "errSmsInvalidUserDataLength"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsInvalidValidityPeriod: getSMSError = "errSmsInvalidValidityPeriod"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsNoSCAddress: getSMSError = "errSmsNoSCAddress"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsTooLongOtherEndAddress: getSMSError = "errSmsTooLongOtherEndAddress"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsTooLongSCAddress: getSMSError = "errSmsTooLongSCAddress"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsTypeCommand: getSMSError = "errSmsTypeCommand"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsTypeMOUndefined: getSMSError = "errSmsTypeMOUndefined"
        Case SMS3ASuiteLib.NmpAdapterError.errSmsTypeMTUndefined: getSMSError = "errSmsTypeMTUndefined"
        Case SMS3ASuiteLib.NmpAdapterError.errSsAbsentSubscriber: getSMSError = "errSsAbsentSubscriber"
        Case SMS3ASuiteLib.NmpAdapterError.errSsActivationDataLost: getSMSError = "errSsActivationDataLost"
        Case SMS3ASuiteLib.NmpAdapterError.errSsBearerServNotProvision: getSMSError = "errSsBearerServNotProvision"
        Case SMS3ASuiteLib.NmpAdapterError.errSsCallBarred: getSMSError = "errSsCallBarred"
        Case SMS3ASuiteLib.NmpAdapterError.errSsCUGReject: getSMSError = "errSsCUGReject"
        Case SMS3ASuiteLib.NmpAdapterError.errSsDataError: getSMSError = "errSsDataError"
        Case SMS3ASuiteLib.NmpAdapterError.errSsDataMissing: getSMSError = "errSsDataMissing"
        Case SMS3ASuiteLib.NmpAdapterError.errSsErrorStatus: getSMSError = "errSsErrorStatus"
        Case SMS3ASuiteLib.NmpAdapterError.errSsFacilityNotSupported: getSMSError = "errSsFacilityNotSupported"
        Case SMS3ASuiteLib.NmpAdapterError.errSsIllegalSsOperation: getSMSError = "errSsIllegalSsOperation"
        Case SMS3ASuiteLib.NmpAdapterError.errSsIncompatibility: getSMSError = "errSsIncompatibility"
        Case SMS3ASuiteLib.NmpAdapterError.errSsMaxnumOfMptyPartExceed: getSMSError = "errSsMaxnumOfMptyPartExceed"
        Case SMS3ASuiteLib.NmpAdapterError.errSsMaxnumOfPwAttViolation: getSMSError = "errSsMaxnumOfPwAttViolation"
        Case SMS3ASuiteLib.NmpAdapterError.errSsMMError: getSMSError = "errSsMMError"
        Case SMS3ASuiteLib.NmpAdapterError.errSsMMRelease: getSMSError = "errSsMMRelease"
        Case SMS3ASuiteLib.NmpAdapterError.errSsMsgError: getSMSError = "errSsMsgError"
        Case SMS3ASuiteLib.NmpAdapterError.errSsNegativePasswordCheck: getSMSError = "errSsNegativePasswordCheck"
        Case SMS3ASuiteLib.NmpAdapterError.errSsNotAvailable: getSMSError = "errSsNotAvailable"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPasswordRegisFailure: getSMSError = "errSsPasswordRegisFailure"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPBadlyStructuredComp: getSMSError = "errSsPBadlyStructuredComp"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPDuplicateInvokeID: getSMSError = "errSsPDuplicateInvokeID"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPDuplicateInvokeID: getSMSError = "errSsPDuplicateInvokeID"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPInitiatingRelease: getSMSError = "errSsPInitiatingRelease"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPLinkedRespUnexpected: getSMSError = "errSsPLinkedRespUnexpected"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPMistypedComp: getSMSError = "errSsPMistypedComp"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPMistypedErrParameter: getSMSError = "errSsPMistypedErrParameter"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPMistypedInvParameter: getSMSError = "errSsPMistypedInvParameter"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPMistypedResParameter: getSMSError = "errSsPMistypedResParameter"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPResourceLimitation: getSMSError = "errSsPResourceLimitation"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPReturnErrorProblem: getSMSError = "errSsPReturnErrorProblem"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPReturnErrorUnexpected: getSMSError = "errSsPReturnErrorUnexpected"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPReturnResultUnexpected: getSMSError = "errSsPReturnResultUnexpected"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPUnexpectedError: getSMSError = "errSsPUnexpectedError"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPUnexpectedLinkedOper: getSMSError = "errSsPUnexpectedLinkedOper"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPUnrecognizedComp: getSMSError = "errSsPUnrecognizedComp"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPUnrecognizedError: getSMSError = "errSsPUnrecognizedError"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPUnrecognizedInvokeID: getSMSError = "errSsPUnrecognizedInvokeID"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPUnrecognizedLinkedID: getSMSError = "errSsPUnrecognizedLinkedID"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPUnrecognizedOperation: getSMSError = "errSsPUnrecognizedOperation"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPWErrorBadPassword: getSMSError = "errSsPWErrorBadPassword"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPWErrorBadPasswordFormat: getSMSError = "errSsPWErrorBadPasswordFormat"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPWErrorEnterNewPassword: getSMSError = "errSsPWErrorEnterNewPassword"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPWErrorEnterNewPasswordAgain: getSMSError = "errSsPWErrorEnterNewPasswordAgain"
        Case SMS3ASuiteLib.NmpAdapterError.errSsPWErrorEnterPassword: getSMSError = "errSsPWErrorEnterPassword"
        Case SMS3ASuiteLib.NmpAdapterError.errSsResourcesNotAvailable: getSMSError = "errSsResourcesNotAvailable"
        Case SMS3ASuiteLib.NmpAdapterError.errSsServiceBusy: getSMSError = "errSsServiceBusy"
        Case SMS3ASuiteLib.NmpAdapterError.errSsSpecificError: getSMSError = "errSsSpecificError"
        Case SMS3ASuiteLib.NmpAdapterError.errSsSubscriptionViolation: getSMSError = "errSsSubscriptionViolation"
        Case SMS3ASuiteLib.NmpAdapterError.errSsSystemFailure: getSMSError = "errSsSystemFailure"
        Case SMS3ASuiteLib.NmpAdapterError.errSsTeleServNotProvision: getSMSError = "errSsTeleServNotProvision"
        Case SMS3ASuiteLib.NmpAdapterError.errSsTimerExpired: getSMSError = "errSsTimerExpired"
        Case SMS3ASuiteLib.NmpAdapterError.errSsUnexpectedDataValue: getSMSError = "errSsUnexpectedDataValue"
        Case SMS3ASuiteLib.NmpAdapterError.errSsUnknownAlphabet: getSMSError = "errSsUnknownAlphabet"
        Case SMS3ASuiteLib.NmpAdapterError.errSsUnknownSubscriber: getSMSError = "errSsUnknownSubscriber"
        Case SMS3ASuiteLib.NmpAdapterError.errSsUSSDBusy: getSMSError = "errSsUSSDBusy"
        Case SMS3ASuiteLib.NmpAdapterError.errTerminalNotReady: getSMSError = "errTerminalNotReady"
        Case SMS3ASuiteLib.NmpAdapterError.errUnknown: getSMSError = "errUnknown"
        Case SMS3ASuiteLib.NmpAdapterError.errUpdateImpossible: getSMSError = "errUpdateImpossible"
        Case SMS3ASuiteLib.NmpAdapterError.errWrongPassword: getSMSError = "errWrongPassword"
    End Select
End Function
