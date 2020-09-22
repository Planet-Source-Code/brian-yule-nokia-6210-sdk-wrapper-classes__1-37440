VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phone_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Phone_Status"
Option Explicit

Public Phone_Status As STTNGS3A_SLib.IPhoneStatus
Attribute Phone_Status.VB_VarHelpID = -1

Private Sub Class_Terminate()
    Set Phone_Status = Nothing
End Sub

Public Property Get ChargingBattery() As Boolean
    ChargingBattery = Phone_Status.BatteryCharging
End Property

Public Property Get LevelBattery() As Long
    LevelBattery = Phone_Status.BatteryLevel
End Property

Public Property Get IMEI() As String
    IMEI = Phone_Status.IMEICode
End Property

Public Property Get IMSI() As String
    IMSI = Phone_Status.IMSICode
End Property

Public Property Get status() As String
    Select Case Phone_Status.PhoneStatus
        Case 1: status = "In Call"
        Case 0: status = "Idle"
        Case 2: status = "No Service"
        Case 3: status = "Off"
    End Select
End Property

Public Property Get RF_Level() As Long
    RF_Level = Phone_Status.RFLevel
End Property

Public Property Get SelectedMemory() As String
    Select Case Phone_Status.SelectedMemory
        Case MemTypeOpt.FDN: SelectedMemory = "FDN"
        Case MemTypeOpt.PHONE: SelectedMemory = "Phone"
        Case MemTypeOpt.SIM: SelectedMemory = "SIM"
    End Select
End Property

Public Property Get ServiceState() As Boolean
    ServiceState = Phone_Status.ServState
End Property
