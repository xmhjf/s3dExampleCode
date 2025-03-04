VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrayPartSeqGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 2009, Intergraph Corporation.  All Rights Reserved.
'
' File:  CTrayPartSeqGen.cls
'
' Author: Sreelekha
'
' Abstract: The file contains the naming rule implementation for Cabletray Stock Parts
'           using the Sequence Id. This name rule will be used by the sequence command
'           and will have to be bulkloaded prior to running the sequence command.
' History:
'      20-Feb-09 Sreelekha  CR144602  Support Sequencing of Cable Tray and Conduit Parts With Sequence Command
'*************************************************************************************************************
Option Explicit

Implements IJNameRule
Private m_oErrors As IJEditErrors
Private Const E_FAIL = -2147467259
Private Const MODULE = "CTrayPartSeqGen: "

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
    If m_oErrors Is Nothing Then
        Set m_oErrors = CreateObject("IMSErrorLog.Errors")
    End If
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''IJNameRule''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''/////////////////////////////////////////////////////////////////////////////
'' METHOD:         IJNameRule_ComputeName
' Description:
'  Creates a name for the object passed in. The Naming Parents are added in AddNamingParents()
'  of the same interface. Both these methods are called from NamingRule Semantic.
'
' Arguments:
'  oObject       - Input.  Child object that needs to have the NamingRule naming.
'  oParents      - Input.  Naming parents collection.
'  oActiveEntity - Input.  Naming rules active entity on which the NamingParentsString is stored.
'/////////////////////////////////////////////////////////////////////////////
Private Sub IJNameRule_ComputeName(ByVal oEntity As Object, ByVal oParents As IJElements, ByVal oActiveEntity As Object)
Const METHOD = "IJNameRule_ComputeName"
On Error GoTo ErrHndler

    Dim oNamedItem As IJNamedItem
    Dim oParent As IJNamedItem
    Dim strPartName As String
    Dim strRunName As String
    Dim strSequenceNo As String
    
    If oEntity Is Nothing Then Exit Sub
    
   'To get the SystemParentName
    Set oParent = oParents.Item(1)
    strRunName = oParent.Name
    Set oParent = Nothing
    
    'Append the SequenceNumber to the name
    strSequenceNo = GetSequenceId(oEntity)
    strPartName = strRunName + "-" + "TRAY" + "-" + strSequenceNo
    Set oNamedItem = oEntity
    If (oNamedItem.Name <> strPartName) Then
        oNamedItem.Name = strPartName
    End If
    
    If m_oErrors Is Nothing Then
        Set m_oErrors = CreateObject("IMSErrorLog.Errors")
    End If
 
    Set oNamedItem = Nothing
    
Exit Sub
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'///////////////////////////////////////////////////////////////////////////////////////
' METHOD:         IJNameRule_GetNamingParents
' Description:
'  All the Naming Parents that need to Participate in an Objects Naming are added here to the
' IJElements collection. The Parents added here are used in Computing the Name of the Object in
' ComputeName() of the same Interface. Both these methods are called from Naming Rule Semantic.
'
' Arguments:
'  oEntity        - Input.  Child object that needs to have the NamingRule naming.
'
'/////////////////////////////////////////////////////////////////////////////
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
Const METHOD = "IJNameRule_GetNamingParents"
On Error GoTo ErrHndler
    
    Dim oParent As IJDesignParent
    Dim oChild As IJDesignChild
    
   'Getting the Parent
    Set oChild = oEntity
    Set oParent = oChild.GetParent
    
    On Error GoTo ErrHndler
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    If Not (oParent Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oParent)
    End If
    
    If m_oErrors Is Nothing Then
        Set m_oErrors = CreateObject("IMSErrorLog.Errors")
    End If
     
    Set oParent = Nothing
    Set oChild = Nothing
    
Exit Function
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Function
