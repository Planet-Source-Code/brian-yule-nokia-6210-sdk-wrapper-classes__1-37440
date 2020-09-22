VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phone_Capability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Phone_Capability"
Option Explicit

Public Phone_Capability As STTNGS3A_SLib.IPhoneCapability

Private Sub Class_Terminate()
    Set Phone_Capability = Nothing
End Sub

'Public Property Get ALARM() As Boolean
'    ALARM = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_ALARM)
'End Property
'Public Property Get ALTERNATING_CALL() As Boolean
'    ALTERNATING_CALL = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_ALTERNATING_CALL)
'End Property
'Public Property Get ANALOG_DATA_CALL() As Boolean
'    ANALOG_DATA_CALL = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_ANALOG_DATA_CALL)
'End Property
'Public Property Get AOC() As Boolean
'    AOC = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_AOC)
'End Property
'Public Property Get CAPAB_CALENDAR() As Boolean
'    CAPAB_CALENDAR = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_CALENDAR)
'End Property
'Public Property Get CALL_SS() As Boolean
'    CALL_SS = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_CALL_SS)
'End Property
'Public Property Get CB_MSG_ROUTING() As Boolean
'    CB_MSG_ROUTING = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_CB_MSG_ROUTING)
'End Property
'Public Property Get CIRCUIT_SWITCH_CALL() As Boolean
'    CIRCUIT_SWITCH_CALL = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_CIRCUIT_SWITCH_CALL)
'End Property
'Public Property Get CLOCK() As Boolean
'    CLOCK = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_CLOCK)
'End Property
'Public Property Get DIG_AUDIO() As Boolean
'    DIG_AUDIO = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_DIG_AUDIO)
'End Property
'Public Property Get DTMF_RECEIV() As Boolean
'    DTMF_RECEIV = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_DTMF_RECEIV)
'End Property
'Public Property Get DTMF_SEND() As Boolean
'    DTMF_SEND = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_DTMF_SEND)
'End Property
'Public Property Get FAST_DATA() As Boolean
'    FAST_DATA = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_FAST_DATA)
'End Property
'Public Property Get FDNSIM() As Boolean
'    FDNSIM = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_FDNSIM)
'End Property
'Public Property Get HALF_RATE_DATA() As Boolean
'    HALF_RATE_DATA = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_HALF_RATE_DATA)
'End Property
'Public Property Get IN_CALL_SS() As Boolean
'    IN_CALL_SS = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_IN_CALL_SS)
'End Property
'Public Property Get KEYBOARD_EVENT() As Boolean
'    KEYBOARD_EVENT = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_KEYBOARD_EVENT)
'End Property
'Public Property Get KEYBOARD_SIM() As Boolean
'    KEYBOARD_SIM = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_KEYBOARD_SIM)
'End Property
'Public Property Get NETWORK_SEL() As Boolean
'    NETWORK_SEL = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_NETWORK_SEL)
'End Property
'Public Property Get PACK_DATA_CALL() As Boolean
'    PACK_DATA_CALL = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_PACK_DATA_CALL)
'End Property
'Public Property Get PB_UPDATE() As Boolean
'    PB_UPDATE = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_PB_UPDATE)
'End Property
'Public Property Get REG_STATUS_EVENTS() As Boolean
'    REG_STATUS_EVENTS = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_REG_STATUS_EVENTS)
'End Property
'Public Property Get SCM_SPEED_UPDAT() As Boolean
'    SCM_SPEED_UPDAT = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_SCM_SPEED_UPDAT)
'End Property
'Public Property Get SECURITY() As Boolean
'    SECURITY = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_SECURITY)
'End Property
'Public Property Get SHOW_INDIGATOR_EVENT() As Boolean
'    SHOW_INDIGATOR_EVENT = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_SHOW_INDIGATOR_EVENT)
'End Property
'Public Property Get SHOW_TEXT_EVENT() As Boolean
'    SHOW_TEXT_EVENT = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_SHOW_TEXT_EVENT)
'End Property
'Public Property Get SIM() As Boolean
'    SIM = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_SIM)
'End Property
'Public Property Get SIM_ACCESS() As Boolean
'    SIM_ACCESS = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_SIM_ACCESS)
'End Property
'Public Property Get SIMLSB() As Boolean
'    SIMLSB = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_SIMLSB)
'End Property
'Public Property Get SIMMSB() As Boolean
'    SIMMSB = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_SIMMSB)
'End Property
'Public Property Get SM_ROUTING() As Boolean
'    SM_ROUTING = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_SM_ROUTING)
'End Property
'Public Property Get SM_SENDING() As Boolean
'    SM_SENDING = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_SM_SENDING)
'End Property
'Public Property Get SMT_MSG_SERVER() As Boolean
'    SMT_MSG_SERVER = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_SMT_MSG_SERVER)
'End Property
'Public Property Get UI_SETTING() As Boolean
'    UI_SETTING = Phone_Capability.IsCapability(STTNGS3A_SLib.CapabilityOpt.CAPAB_UI_SETTING)
'End Property
