VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHoleCutNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2003, Intergraph Corporation.  All rights reserved.
'
'  Project: GSCADHoleNamingRules
'
'  Abstract: The file contains a implementation of a HoleFeatureNameRule
'    for Hole Management
'
'  Author: C.R.Swager (CRS)
'
' Changes:
'    Date    By    Reason
'  05/28/03  CRS  Original issue.
'
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Private Const Module = "CHoleCutNameRule: "
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "00"       ' define fixed-width number field for name

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'***********************************************************************
' Description:
' Creates a name for the object passed in. The name is based on the
' parent's name.  It is something like this: "Base Name"
' + "-Cut" + Index. "Base Name" is derived from the Naming
' Parents and the Index is unique for the "Base Name".  It is assumed
' that all Naming Parents and the Object implement IJNamedItem.  The
' Naming Parents are added in AddNamingParents() of the same interface.
' Both these methods are called from naming rule semantic.
'***********************************************************************
Private Sub IJNameRule_ComputeName _
       (ByVal oHoleFeature As Object, _
        ByVal oNamingParents As IJElements, _
        ByVal oActiveEntity As Object)

    Const Method = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
    Dim strError As String
    
    If oHoleFeature Is Nothing Then
        strError = "HoleFeature is nothing"
        GoTo ErrorHandler
    End If

    If oNamingParents.Count = 0 Then
        strError = "HoleFeature has no parents"
        GoTo ErrorHandler
    End If
    
    Dim oFeature As IJHoleFeature
    Set oFeature = oHoleFeature
    
    'is this a cut for the hole or a cut on the center flange
    Dim oHoleTrace As IJHoleTraceAE
    Set oHoleTrace = oFeature.HoleTrace
    
    Dim strFeatureName As String
    If Not oHoleTrace Is Nothing Then
        Dim oParent As Object
        Set oParent = oNamingParents.Item(1)
    
        Dim oParentNamedItem As IJNamedItem
        Set oParentNamedItem = oParent
        Dim strHoleName As String
        strHoleName = oParentNamedItem.Name
        strFeatureName = strHoleName & "-Cut"
    
        Set oParent = Nothing
        Set oParentNamedItem = Nothing
    Else
        strFeatureName = "CenterFlange-Cut"
    End If
    
    ' Use the name generator service to compute a unique number to append
    ' to the name generated.
    Dim oNameCounter As IJNameCounter
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    
    ' Get the middle context (to set up the name counter).
    Dim jContext As IJContext
    Set jContext = GetJContext()
    
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    
    Dim oConnectMiddle As IJDAccessMiddle
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    Dim strModelDBID As String
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    
    Dim oModelResourceMgr As IUnknown
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)

    ' Finally everything is set up to obtain a unique number to append to
    ' the Feature name.
    Dim nCount As Long
    Dim strLocationID As String
    nCount = oNameCounter.GetCountEx(oModelResourceMgr, strFeatureName, strLocationID)

    Dim oHoleFeatureNamedItem As IJNamedItem
    Set oHoleFeatureNamedItem = oHoleFeature
    
    oHoleFeatureNamedItem.Name = strFeatureName & strLocationID & Format(nCount, strCountFormat)
    
    GoTo Cleanup

ErrorHandler:
'    ReportAndRaiseUnanticipatedError Method, Module
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName: " * strError, Err.Description
    
Cleanup:
    Set oHoleFeatureNamedItem = Nothing
    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oModelResourceMgr = Nothing
    Set oNameCounter = Nothing
    Set oHoleTrace = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' All the Naming Parents that need to participate in an objects naming
' are added here to the IJElements collection. The parents added here
' are used in computing the name of the object in ComputeName() of the
' same interface. Both these methods are called from naming rule
' semantic.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oHoleFeature As Object) As IJElements

    Const Method = "IJNameRule_GetNamingParents"

    On Error Resume Next
    Dim oChild As IJDesignChild
    Set oChild = oHoleFeature
    Dim oParent As IJDesignParent
    Set oParent = oChild.GetParent

    On Error GoTo ErrorHandler
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oParent Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oParent)
    End If

    GoTo Cleanup
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError Method, Module
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    
Cleanup:
    Set oParent = Nothing
    Set oChild = Nothing

End Function

 