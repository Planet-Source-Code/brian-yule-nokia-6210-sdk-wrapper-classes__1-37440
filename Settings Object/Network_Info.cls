VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Network_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"AOC_Info"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private WithEvents Settings As STTNGS3A_SLib.Network_Suite3
Attribute Settings.VB_VarHelpID = -1

Public Event SignalChange(newSignal As Long)
Private mvarAOC_Info As AOC_Info

Public Property Get AOC_Info() As AOC_Info
    If mvarAOC_Info Is Nothing Then
        Set mvarAOC_Info = New AOC_Info
        Set mvarAOC_Info.NetAOC = Settings
    End If
    Set AOC_Info = mvarAOC_Info
End Property

Private Sub Class_Initialize()
    Set Settings = New STTNGS3A_SLib.Network_Suite3
    
    Call Settings.StartListeningAOCEvents
    Call Settings.StartListeningNetworkEvents
End Sub

Private Sub Class_Terminate()
    Set mvarAOC_Info = Nothing
    Call Settings.Terminate
    Set Settings = Nothing
End Sub

Public Property Get operatorName() As String
    Dim State As STTNGS3A_SLib.NETWORK_REG_STATUS, sel As STTNGS3A_SLib.NETWORK_SELECTION_MODE, _
    cellID As Integer, lac As Integer, opercode As Long, operName As String
    
    Call Settings.GetRegistrationStatus(State, sel, cellID, lac, opercode, operName)
    
    operatorName = operName
End Property

Public Property Get operatorCode() As String
    Dim State As STTNGS3A_SLib.NETWORK_REG_STATUS, sel As STTNGS3A_SLib.NETWORK_SELECTION_MODE, _
    cellID As Integer, lac As Integer, opercode As Long, operName As String
    
    Call Settings.GetRegistrationStatus(State, sel, cellID, lac, opercode, operName)
    
    operatorCode = opercode
End Property

Public Property Get cellIdentifier() As Integer
    Dim State As STTNGS3A_SLib.NETWORK_REG_STATUS, sel As STTNGS3A_SLib.NETWORK_SELECTION_MODE, _
    cellID As Integer, lac As Integer, opercode As Long, operName As String
    
    Call Settings.GetRegistrationStatus(State, sel, cellID, lac, opercode, operName)
    
    cellIdentifier = cellID
End Property

Public Property Get selectionMode() As String
    Dim State As STTNGS3A_SLib.NETWORK_REG_STATUS, sel As STTNGS3A_SLib.NETWORK_SELECTION_MODE, _
    cellID As Integer, lac As Integer, opercode As Long, operName As String
    
    Call Settings.GetRegistrationStatus(State, sel, cellID, lac, opercode, operName)
    
    Select Case State
        Case 0: selectionMode = "Manual"
        Case 1: selectionMode = "Automatic"
        Case 255: selectionMode = "Unknown"
    End Select
End Property

Public Property Get regState() As String
    Dim State As STTNGS3A_SLib.NETWORK_REG_STATUS, sel As STTNGS3A_SLib.NETWORK_SELECTION_MODE, _
    cellID As Integer, lac As Integer, opercode As Long, operName As String
    
    Call Settings.GetRegistrationStatus(State, sel, cellID, lac, opercode, operName)
    
    Select Case State
        Case 0: regState = "Home"
        Case 1: regState = "No Service, No SIM"
        Case 2: regState = "No Service, Searching"
        Case 3: regState = "No Service, NOT Searching"
        Case 4: regState = "Roam"
    End Select
End Property

Public Property Get IMSI() As String
    IMSI = Settings.IMSI
End Property

Private Sub Settings_SignalStrengthChanges(ByVal nSignal As Long)
    RaiseEvent SignalChange(nSignal)
End Sub


