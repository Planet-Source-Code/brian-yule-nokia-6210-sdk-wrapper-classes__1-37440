VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phone_Control_Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Clock As STTNGS3A_SLib.IClock

Private Sub Class_Terminate()
    Clock.Terminate
    Set Clock = Nothing
End Sub

Public Property Get AlarmActive() As Boolean
    Dim pVal As Date
    Call Clock.GetAlarm(AlarmActive, pVal)
End Property

Public Property Get AlarmDate() As Date
    Dim Active As Boolean
    Call Clock.GetAlarm(Active, AlarmDate)
End Property

Public Property Let AlarmDate(ByVal newDate As Date)
    Call Clock.SetAlarm(AlarmActive, newDate)
End Property

Public Property Let AlarmActivate(ByVal State As Boolean)
    Call Clock.SetAlarm(State, AlarmDate)
End Property

