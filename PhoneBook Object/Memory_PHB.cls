VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Memory_PHB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public phb2 As PhonebookAdapterDS3.IPhonebook2
Public MemoryLocationIndex As Long

Public Property Get Free() As Long
    On Error GoTo ErrorHandler
    
    Dim pTotal As Long, pFull As Long, pFree As Long
    
    Call phb2.GetMemoryCaps(MemoryLocationIndex, pTotal, pFull, pFree)
        
    Free = pFree
Exit Property
ErrorHandler:
    On Error Resume Next
        
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print phb2.GetLastError, getPHBError(phb2.GetLastError)
End Property

Public Property Get Full() As Long
    On Error GoTo ErrorHandler
    
    Dim pTotal As Long, pFull As Long, pFree As Long
    
    Call phb2.GetMemoryCaps(MemoryLocationIndex, pTotal, pFull, pFree)
    
    Full = pFull
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print phb2.GetLastError, getPHBError(phb2.GetLastError)
End Property

Public Property Get Total() As Long
    On Error GoTo ErrorHandler
    
    Dim pTotal As Long, pFull As Long, pFree As Long
    
    Call phb2.GetMemoryCaps(MemoryLocationIndex, pTotal, pFull, pFree)
    
    Total = pTotal
Exit Property
ErrorHandler:
    On Error Resume Next
    
    If Err.Number <> 0 Then Debug.Print Err.Number, Err.Description Else Debug.Print phb2.GetLastError, getPHBError(phb2.GetLastError)
End Property

