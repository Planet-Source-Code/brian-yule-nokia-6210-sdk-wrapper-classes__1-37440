VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Security_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private Settings As STTNGS3A_SLib.Security_Suite3

Private Sub Class_Initialize()
    Set Settings = New STTNGS3A_SLib.Security_Suite3
End Sub

Private Sub Class_Terminate()
    Settings.Terminate
End Sub

Public Property Let ChangeLock(ByVal OldPassword As String, ByVal NewPassword As String)
    Call Settings.ChangePw(7, OldPassword, NewPassword)
End Property
Public Property Let ChangeMaster(ByVal OldPassword As String, ByVal NewPassword As String)
    Call Settings.ChangePw(6, OldPassword, NewPassword)
End Property
Public Property Let ChangeNone(ByVal OldPassword As String, ByVal NewPassword As String)
    Call Settings.ChangePw(0, OldPassword, NewPassword)
End Property
Public Property Let ChangePIN(ByVal OldPassword As String, ByVal NewPassword As String)
    Call Settings.ChangePw(1, OldPassword, NewPassword)
End Property
Public Property Let ChangePIN2(ByVal OldPassword As String, ByVal NewPassword As String)
    Call Settings.ChangePw(2, OldPassword, NewPassword)
End Property
Public Property Let ChangePUK(ByVal OldPassword As String, ByVal NewPassword As String)
    Call Settings.ChangePw(3, OldPassword, NewPassword)
End Property
Public Property Let ChangePUK2(ByVal OldPassword As String, ByVal NewPassword As String)
    Call Settings.ChangePw(4, OldPassword, NewPassword)
End Property
Public Property Let ChangeSecurityCode(ByVal OldPassword As String, ByVal NewPassword As String)
    Call Settings.ChangePw(5, OldPassword, NewPassword)
End Property

Public Property Get PasswordStatus() As String
    Select Case Settings.GetPwStatus
        Case 7: PasswordStatus = "Lock"
        Case 6: PasswordStatus = "Master"
        Case 0: PasswordStatus = "None"
        Case 1: PasswordStatus = "PIN"
        Case 2: PasswordStatus = "PIN2"
        Case 3: PasswordStatus = "PUK"
        Case 4: PasswordStatus = "PUK2"
        Case 5: PasswordStatus = "SecurityCode"
    End Select
End Property
