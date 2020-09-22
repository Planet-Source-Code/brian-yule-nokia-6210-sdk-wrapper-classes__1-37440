VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phone_Control_Media_Pool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public MediaPool As STTNGS3A_SLib.IMediaPool

Private Sub Class_Terminate()
    Set MediaPool = Nothing
End Sub

Public Function ToneCreate() As STTNGS3A_SLib.Tone
    Set ToneCreate = MediaPool.CreateTone
End Function

Public Sub ToneSave(pITone As STTNGS3A_SLib.Tone)
    Call MediaPool.SaveTone(pITone)
End Sub
