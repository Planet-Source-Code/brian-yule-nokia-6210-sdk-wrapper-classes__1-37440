VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Version"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public phb2 As PhonebookAdapterDS3.IPhonebook2

Public Property Get Software() As String
    On Error GoTo ErrorHandler
    
    Dim ppsIMEI As String, ppsType As String, ppsVerSW As String, ppsVerHW As String
    
    Call phb2.GetDevIdentificationInfo(ppsIMEI, ppsType, ppsVerSW, ppsVerHW)
    
    Software = ppsVerSW
Exit Property
ErrorHandler:
    On Error Resume Next
    Debug.Print Err.Number, Err.Description,
    Debug.Print phb2.GetLastError, getPHBError(phb2.GetLastError)
End Property

Public Property Get Hardware() As String
    On Error GoTo ErrorHandler
    
    Dim ppsIMEI As String, ppsType As String, ppsVerSW As String, ppsVerHW As String
    
    Call phb2.GetDevIdentificationInfo(ppsIMEI, ppsType, ppsVerSW, ppsVerHW)
    
    Hardware = ppsVerHW
Exit Property
ErrorHandler:
    On Error Resume Next
    Debug.Print Err.Number, Err.Description,
    Debug.Print phb2.GetLastError, getPHBError(phb2.GetLastError)
End Property
