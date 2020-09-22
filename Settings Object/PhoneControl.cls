VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phone_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Phone_Control_2"
Attribute VB_Ext_KEY = "Member1" ,"Phone_Control_Clock"
Attribute VB_Ext_KEY = "Member2" ,"Phone_Control_Media_Pool"
Attribute VB_Ext_KEY = "Member3" ,"UserInterface"
Option Explicit

Private Settings As STTNGS3A_SLib.PhoneControl_Suite3
Attribute Settings.VB_VarHelpID = -1
Private mvarUserInterface As UserInterface
Private mvarPhone_Control_2 As Phone_Control_2
Private mvarPhone_Control_Clock As Phone_Control_Clock
Private mvarPhone_Control_Media_Pool As Phone_Control_Media_Pool


Public Property Get Phone_Control_Media_Pool() As Phone_Control_Media_Pool
    If mvarPhone_Control_Media_Pool Is Nothing Then
        Set mvarPhone_Control_Media_Pool = New Phone_Control_Media_Pool
        Set mvarPhone_Control_Media_Pool.MediaPool = Settings
    End If
    Set Phone_Control_Media_Pool = mvarPhone_Control_Media_Pool
End Property

Public Property Get Phone_Control_Clock() As Phone_Control_Clock
    If mvarPhone_Control_Clock Is Nothing Then
        Set mvarPhone_Control_Clock = New Phone_Control_Clock
        Set mvarPhone_Control_Clock.Clock = Settings
    End If
    Set Phone_Control_Clock = mvarPhone_Control_Clock
End Property

Public Property Get Phone_Control_2() As Phone_Control_2
    If mvarPhone_Control_2 Is Nothing Then
        Set mvarPhone_Control_2 = New Phone_Control_2
        Set mvarPhone_Control_2.PhoneControl2 = Settings
    End If
    Set Phone_Control_2 = mvarPhone_Control_2
End Property

Public Property Get UserInterface() As UserInterface
    If mvarUserInterface Is Nothing Then
        Set mvarUserInterface = New UserInterface
        Set mvarUserInterface.UserInterface = Settings
    End If
    Set UserInterface = mvarUserInterface
End Property

Private Sub Class_Initialize()
    Set Settings = New STTNGS3A_SLib.PhoneControl_Suite3
End Sub

Private Sub Class_Terminate()
  Set mvarPhone_Control_Media_Pool = Nothing
    Set mvarPhone_Control_Clock = Nothing
    Set mvarPhone_Control_2 = Nothing
    Set mvarUserInterface = Nothing
    Settings.Terminate
    Set Settings = Nothing
End Sub

Public Sub PhoneRestoreFactorySetting()
    Call Settings.RestoreFactorySettings
End Sub

Public Sub PhoneClose()
    Call Settings.ClosePhone
End Sub

