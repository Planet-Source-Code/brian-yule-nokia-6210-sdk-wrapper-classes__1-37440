VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Network_Info"
Attribute VB_Ext_KEY = "Member1" ,"Phone_Control"
Attribute VB_Ext_KEY = "Member2" ,"Phone_Info"
Attribute VB_Ext_KEY = "Member3" ,"Security_Info"
Option Explicit

Public Event SignalLevelChange(ByVal signalLevel As Long)
Public Event BatteryNotify(ByVal Level As Long, Charging As Boolean)
Public Event DeviceNotify(ByVal State As String)

Private WithEvents mvarNetwork_Info As Network_Info
Attribute mvarNetwork_Info.VB_VarHelpID = -1
Private WithEvents mvarPhone_Info As Phone_Info
Attribute mvarPhone_Info.VB_VarHelpID = -1
Private mvarPhone_Control As Phone_Control
Private mvarSecurity_Info As Security_Info

Public Property Get Security_Info() As Security_Info
    If mvarSecurity_Info Is Nothing Then Set mvarSecurity_Info = New Security_Info
    Set Security_Info = mvarSecurity_Info
End Property

Public Property Get Phone_Control() As Phone_Control
    If mvarPhone_Control Is Nothing Then Set mvarPhone_Control = New Phone_Control
    Set Phone_Control = mvarPhone_Control
End Property

Public Property Get Phone_Info() As Phone_Info
    If mvarPhone_Info Is Nothing Then Set mvarPhone_Info = New Phone_Info
    Set Phone_Info = mvarPhone_Info
End Property

Public Property Get Network_Info() As Network_Info
    If mvarNetwork_Info Is Nothing Then Set mvarNetwork_Info = New Network_Info
    Set Network_Info = mvarNetwork_Info
End Property

Private Sub Class_Terminate()
  Set mvarSecurity_Info = Nothing
    Set mvarPhone_Control = Nothing
    Set mvarPhone_Info = Nothing
    Set mvarNetwork_Info = Nothing
End Sub

Private Sub mvarNetwork_Info_SignalChange(newSignal As Long)
    RaiseEvent SignalLevelChange(newSignal)
End Sub

Private Sub mvarPhone_Info_Battery(ByVal Level As Long, Charging As Boolean)
    RaiseEvent BatteryNotify(Level, Charging)
End Sub

Private Sub mvarPhone_Info_Device(ByVal State As String)
    RaiseEvent DeviceNotify(State)
End Sub
