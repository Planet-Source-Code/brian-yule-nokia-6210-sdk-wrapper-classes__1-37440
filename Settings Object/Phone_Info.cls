VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phone_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Phone_Capability"
Attribute VB_Ext_KEY = "Member1" ,"Phone_Status"
Option Explicit

Private WithEvents Settings As STTNGS3A_SLib.PhoneInfo_Suite3
Attribute Settings.VB_VarHelpID = -1

Private mvarPhoneCapability As Phone_Capability
Private mvarPhoneStatus As Phone_Status

Public Event Battery(ByVal Level As Long, Charging As Boolean)
Public Event Device(ByVal State As String)

Public Property Get PhoneStatus() As Phone_Status
    If mvarPhoneStatus Is Nothing Then
        Set mvarPhoneStatus = New Phone_Status
        Set mvarPhoneStatus.Phone_Status = Settings
    End If
    Set PhoneStatus = mvarPhoneStatus
End Property

Public Property Get PhoneCapability() As Phone_Capability
    If mvarPhoneCapability Is Nothing Then
        Set mvarPhoneCapability = New Phone_Capability
        Set mvarPhoneCapability.Phone_Capability = Settings
    End If
    Set PhoneCapability = mvarPhoneCapability
End Property

Public Property Get DeviceStatus() As String
    Dim pVal As STTNGS3A_SLib.DevNotifyOpt
     
    Call Settings.get_DeviceStatus(pVal)

    Select Case pVal
        Case STTNGS3A_SLib.DevNotifyOpt.ATTACHED: DeviceStatus = "Attached"
        Case STTNGS3A_SLib.DevNotifyOpt.DISCONNECTED: DeviceStatus = "Disconnected"
        Case STTNGS3A_SLib.DevNotifyOpt.REMOVED: DeviceStatus = "Removed"
        Case STTNGS3A_SLib.DevNotifyOpt.UNKNOWN: DeviceStatus = "Unknown"
    End Select
End Property

Public Property Get HardwareVersion() As String
    HardwareVersion = Settings.HwVersion
End Property

Public Property Get ProductCode() As String
    ProductCode = Settings.ProductCode
End Property

Public Property Get ProductType() As String
    ProductType = Settings.ProductType
End Property

Public Property Get SorfwareVersion() As String
    SorfwareVersion = Settings.SwVersion
End Property

Private Sub Class_Initialize()
    Set Settings = New STTNGS3A_SLib.PhoneInfo_Suite3
    
    Call Settings.StartListeningPhoneInfoEvents
End Sub

Private Sub Class_Terminate()
    Set mvarPhoneStatus = Nothing
    Set mvarPhoneCapability = Nothing
    
    Call Settings.Terminate
    
    Set Settings = Nothing
End Sub

Private Sub Settings_BatteryNotify(ByVal bLevel As Long, ByVal bCharging As Boolean)
    RaiseEvent Battery(bLevel, bCharging)
End Sub

Private Sub Settings_DeviceNotify(ByVal val As STTNGS3A_SLib.DevNotifyOpt)
    Select Case val
        Case STTNGS3A_SLib.DevNotifyOpt.ATTACHED: RaiseEvent Device("Attached")
        Case STTNGS3A_SLib.DevNotifyOpt.DISCONNECTED: RaiseEvent Device("Disconnected")
        Case STTNGS3A_SLib.DevNotifyOpt.REMOVED: RaiseEvent Device("Removed")
        Case STTNGS3A_SLib.DevNotifyOpt.UNKNOWN: RaiseEvent Device("Unknown")
    End Select
End Sub
