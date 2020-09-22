VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AOC_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public NetAOC As STTNGS3A_SLib.INetAOC

Public Property Get ACM() As Long
    Dim pICCM As Long, pIACM As Long, pIACMMax As Long
    
    Call NetAOC.GetAOCParameters(pICCM, pIACM, pIACMMax)
    
    ACM = pIACM
End Property

Public Property Get ACMMax() As Long
    Dim pICCM As Long, pIACM As Long, pIACMMax As Long
    
    Call NetAOC.GetAOCParameters(pICCM, pIACM, pIACMMax)
    
    ACMMax = pIACMMax
End Property

Public Property Get CCM() As Long
    Dim pICCM As Long, pIACM As Long, pIACMMax As Long
    
    Call NetAOC.GetAOCParameters(pICCM, pIACM, pIACMMax)
    
    CCM = pICCM
End Property

