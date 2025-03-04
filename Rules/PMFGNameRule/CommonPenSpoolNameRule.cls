VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonPenSpoolNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2006 Intergraph Corporation.  All rights reserved.
'
'  Project: PMFGNameRule
'  Class:   CommonPenSpoolNameRule
'
'  Abstract: The file contains the Implementation for naming rule interface for Pen Spools
'
'  Author:
'
'  History:
'       07/24/2006  Sreelekha   TR#101967  Creating penetration spool has given an error
'******************************************************************
Option Explicit

Implements IJNameRule
Const vbInvalidArg = &H80070057

Private Const MODULE = "CommonPenSpoolNameRule: "

Private Const E_FAIL = -2147467259
Dim m_oErrors As IJEditErrors  ' Middle tier errors

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub
Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
Const METHOD = "IJNameRule_ComputeName "
    Dim oSpoolNamedItem As IJNamedItem
    Dim oSpool As IJDPenSpool
    Dim strName As String
    Dim strParentName As String
    Dim oPlate As Object
    On Error GoTo ErrorHandler
    
    If pEntity Is Nothing Then
        Err.Raise vbInvalidArg, MODULE, METHOD
    End If
    
    Set oSpoolNamedItem = pEntity
    Set oSpool = pEntity
     
    Set oPlate = pParents.Item(1)
     
    strName = m_GetSpoolName(oPlate, "PEN")  'oSpoolNamedItem.TypeString)
    
    oSpoolNamedItem.Name = strName
    
    Set oSpool = Nothing
    Set oPlate = Nothing
    Set oSpoolNamedItem = Nothing
        
    Exit Sub
    
ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IJElements
Const METHOD = "IJNameRule_GetNamingParents"
On Error GoTo ErrorHandler
    
    Dim oChild As Object
    Dim oPenSpool As IJMfgParent
    Dim oChildren As IJDObjectCollection
    Dim oNewCol As IJDObjectCollection
    Dim oJDSpool As IJDSpool
    
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
   
    Set oPenSpool = pEntity
    
    oPenSpool.GetChildren oChildren, "IsChildOf"             ' AssemblyHierarchy
    oPenSpool.GetChildren oNewCol, "IsSpoolChildOf"         ' Spool Hierarchy
    
    If (oNewCol.Count > 0) Then _
        oChildren.SetAdd oNewCol
    
    On Error Resume Next
    
    ' Get the penetration plate (naming parent) from the collection
    For Each oChild In oChildren
        Set oJDSpool = oChild
        If oJDSpool Is Nothing Then
            Exit For
        End If
        Set oJDSpool = Nothing
        Set oChild = Nothing
    Next oChild
    
    IJNameRule_GetNamingParents.Add oChild
         
    oChildren.Clear
    Set oChildren = Nothing
    Set oJDSpool = Nothing
    Set oPenSpool = Nothing
    Set oChild = Nothing
    
    Exit Function
ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Function
Private Function m_GetSpoolName(oPipeLine As Object, strUserName As String) As String

Const METHOD = "m_GetSpoolName "
Dim strSpoolName As String
Dim oParentNamedItem As IJNamedItem
Dim ParentName As String
On Error GoTo ErrorHandler
    
    Set oParentNamedItem = oPipeLine
    
    ParentName = oParentNamedItem.Name

   
  '  strUserName = "Spool"
    
    
    strSpoolName = ParentName & "_" & strUserName


    m_GetSpoolName = strSpoolName
    
   
    Set oParentNamedItem = Nothing

Exit Function
ErrorHandler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Function

