VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"UserInterface"
Option Explicit

Public UserInterface As STTNGS3A_SLib.IPhoneUI

Private Sub Class_Terminate()
    UserInterface.Terminate
    
    Set UserInterface = Nothing
End Sub

Public Property Get AllowReplyCharge() As Boolean
    AllowReplyCharge = UserInterface.AllowReplyCost
End Property

Public Property Get AnswerAnyKey() As Boolean
    AnswerAnyKey = UserInterface.AnykeyAnswer
End Property

Public Property Get AutoPowerOff() As Boolean
    AutoPowerOff = UserInterface.AutomaticPowerOff
End Property

Public Property Get AutoPowerOffTime() As Integer
    AutoPowerOffTime = UserInterface.AutomaticPowerOffTime
End Property

Public Property Get AutoRedial() As Boolean
    AutoRedial = UserInterface.AutomaticRedial
End Property

Public Property Get CellInformationDisplay() As Boolean
    CellInformationDisplay = UserInterface.CellInfoDisplay
End Property

Public Property Get DeliveryReport() As Boolean
    DeliveryReport = UserInterface.DeliveryReports
End Property

Public Property Get ModeSilent() As Boolean
    ModeSilent = UserInterface.SilentMode
End Property

Public Property Get KeySpeedDialing() As Boolean
    KeySpeedDialing = UserInterface.SpeedDialing
End Property
Public Property Get WelcomingText() As String
    WelcomingText = UserInterface.WelcomeText
End Property

Public Property Get WelcomingLogo() As Integer
    WelcomingLogo = UserInterface.WelcomeLogo
End Property

