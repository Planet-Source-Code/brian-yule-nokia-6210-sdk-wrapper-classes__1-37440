VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phone_Control_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public PhoneControl2 As STTNGS3A_SLib.IPhoneControl2

Private Sub Class_Terminate()
    PhoneControl2.Terminate
    Set PhoneControl2 = Nothing
End Sub

Public Sub ChangeMedia_BLUETOOTH()
    Call PhoneControl2.ChangeMedia(STTNGS3A_SLib.MediaType.MEDIA_BLUETOOTH)
End Sub

Public Sub ChangeMedia_DRIVER()
    Call PhoneControl2.ChangeMedia(STTNGS3A_SLib.MediaType.MEDIA_DRIVER)
End Sub

Public Sub ChangeMedia_IRDA()
    Call PhoneControl2.ChangeMedia(STTNGS3A_SLib.MediaType.MEDIA_IRDA)
End Sub

Public Sub ChangeMedia_RS232()
    Call PhoneControl2.ChangeMedia(STTNGS3A_SLib.MediaType.MEDIA_RS232)
End Sub

Public Sub ChangeMedia_UNKNOWN()
    Call PhoneControl2.ChangeMedia(STTNGS3A_SLib.MediaType.MEDIA_UNKNOWN)
End Sub

Public Property Let PortNumber(ByVal IPortNumber As Long)
    Call PhoneControl2.ChangePort(IPortNumber)
End Property

Public Sub PhoneClose()
    Call PhoneControl2.ClosePhone
End Sub

Public Property Get CurrentMedia() As String
    Dim mMedia As STTNGS3A_SLib.MediaType, pIPortNumber As Long
    
    Call PhoneControl2.GetCurrentMedia(mMedia, pIPortNumber)
    
    Select Case mMedia
        Case STTNGS3A_SLib.MediaType.MEDIA_BLUETOOTH: CurrentMedia = "Blue Tooth"
        Case STTNGS3A_SLib.MediaType.MEDIA_DRIVER: CurrentMedia = "Driver"
        Case STTNGS3A_SLib.MediaType.MEDIA_IRDA: CurrentMedia = "IRDA"
        Case STTNGS3A_SLib.MediaType.MEDIA_RS232: CurrentMedia = "RS-232"
        Case STTNGS3A_SLib.MediaType.MEDIA_UNKNOWN: CurrentMedia = "Unknown"
    End Select
End Property

Public Property Get CurrentPort() As Long
    Dim mMedia As STTNGS3A_SLib.MediaType, pIPortNumber As Long
    
    Call PhoneControl2.GetCurrentMedia(mMedia, pIPortNumber)
    
    CurrentPort = pIPortNumber
End Property

Public Property Get SupportedMedia() As String
    Dim mMedia As STTNGS3A_SLib.MediaType
    
    Call PhoneControl2.GetSupportedMedia(mMedia)
    
    Select Case mMedia
        Case STTNGS3A_SLib.MediaType.MEDIA_BLUETOOTH: SupportedMedia = "Blue Tooth"
        Case STTNGS3A_SLib.MediaType.MEDIA_DRIVER: SupportedMedia = "Driver"
        Case STTNGS3A_SLib.MediaType.MEDIA_IRDA: SupportedMedia = "IRDA"
        Case STTNGS3A_SLib.MediaType.MEDIA_RS232: SupportedMedia = "RS-232"
        Case STTNGS3A_SLib.MediaType.MEDIA_UNKNOWN: SupportedMedia = "Unknown"
    End Select
End Property
