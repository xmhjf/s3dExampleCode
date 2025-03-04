VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project: PWeldNameRule
'  Class:   WeldNameRule
'
'  Abstract: The file contains the Implementation for naming rule interface for Welds
'
'  Author: Prasad
'
'  History:
'
'  25/02/2002   Prasad  RI#16673: Implemented IJNameRule interface for welds.
'  23/05/2002   Prasad  Updated with code review comments.
'  11/06/2002   Mohon   Made sure GetCount was not passed "" for Pipeline name.
'  18/06/2002   Mohon   Now middle tier error handling is used.
'******************************************************************

Option Explicit

Implements IJNameRule

Const vbInvalidArg = &H80070057
Private Const MODULE = "WeldNameRule: "
Private Const E_FAIL = -2147467259
Private Const MODELDATABASE = "Model"
Dim m_oErrors As IJEditErrors  ' Middle tier errors

Private Sub Class_Initialize()
    'DI-CP-164764    Prevent DLL unloading for high usage DLLs
    If m_ForceDllInProcess = 0 Then
        m_ForceDllInProcess = ForceDllInProcess(AddressOf ForceDllInProcess)
    End If
    
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
'
' Creates a name for the object passed in. The name is based on the Pipeline
' to which the Weld belongs.  It is something like this: "Pipeline Name" + Index.
' The Pipeline is found in AddNamingParents() and is returned as the first
' item in an elements collection.  Then this sub is called and is passed that
' elements collection.  The naming rule semantic first calls AddNamingParents()
' and then this sub, ComputeName().
'
' Other naming rules may base the name on other parents and may construct a
' different form for the name.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)
    On Error GoTo ERR_SUB
    
    Const METHOD = "IJNameRule_ComputeName"
    Dim oWeldNamedItem As IJNamedItem
    Dim oWeldObj As IJRteWeld
    Dim oPipelineNamedItem As IJNamedItem
    Dim oNameCounter As IJNameCounter
    Dim ojContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    Dim strWeldName As String
    Dim lCount As Long
    Dim strPipelineName As String

    'Get the Model Resource manager
    'This manager will help us to build a unique name
    Set ojContext = GetJContext()
    If ojContext Is Nothing Then Exit Sub
    
    Set oDBTypeConfig = ojContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = ojContext.GetService("ConnectMiddle")
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    If oModelResourceMgr Is Nothing Then Exit Sub
    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, MODULE, METHOD
    End If
    On Error Resume Next
    
    'GetNamingParents returns an elements collection.
    'GetNamingParents placed this Weld's Pipeline as the first item in the collection.
    'The elements argument that is passed to this sub is that elements collection.
    
    Set oPipelineNamedItem = elements.Item(1) 'Get the IJNamedItem inteface for the Pipeline
    Set oWeldNamedItem = oObject 'This is the interface on which we change the Weld's name
    Set oWeldObj = oWeldNamedItem
    On Error GoTo ERR_SUB
    
    If oNameCounter Is Nothing Then
        Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    End If
    
    'Make sure that the name of the Pipeline is not NULL.  This will cause problems
    'for GetCount.
    If Not oPipelineNamedItem Is Nothing Then
        strPipelineName = oPipelineNamedItem.Name
    End If
    
    If strPipelineName = "" Then
        strPipelineName = "Pipeline"
    End If
    
    'If base part of name hasn't changed, don't compute a new name.
    If oWeldNamedItem.Name = (strPipelineName & CStr(oWeldObj.SequenceNumber)) Then Exit Sub
    
    'Instead of using sequence number it has been decided to use GetCount on IJNameCounter for naming.
    'The counter that we get will be a unique number for objects whose names begin with the name
    'of the Pipeline.  The number will also be used as a unique sequence number for the Weld.
    If Not oNameCounter Is Nothing Then
        lCount = oNameCounter.GetCount(oModelResourceMgr, strPipelineName)
        oWeldObj.SequenceNumber = lCount
        Set oWeldObj = Nothing
    End If
        
    'Build the Weld's name and set it.
    strWeldName = strPipelineName + CStr(lCount)
    oWeldNamedItem.Name = strWeldName

    'Release object references.
    Set oWeldNamedItem = Nothing
    Set oPipelineNamedItem = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oNameCounter = Nothing
    
    
    Exit Sub
ERR_SUB:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
    
End Sub
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
'
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parents added here are used in computing the name of the object in
' ComputeName() of the same interface. Both these methods are called from naming rule semantic.
' In the case of Welds, we are using the Pipeline as our naming parent. Pipeline->Run->Part->Weld relation
' is used here
'
' The naming rule semantic will take the parents returned in the elements collection and
' create a relationship so that whenever their name changes, the rule is again fired and
' the name is updated by this rule.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    On Error GoTo ERR_SUB
    
    Const METHOD = "IJNameRule_GetNamingParents"
    
    'Create elements collection in which to return parents on which the name
    'of the object are to be based.  In this case, only the Pipeline is needed.
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    If oEntity Is Nothing Then Exit Function 'No Weld
    
    Dim oDesignParent As IJDesignParent
    Dim oDesignChild As IJDesignChild
    
    'Get Part, which is parent of Weld
    Set oDesignChild = oEntity
    Set oDesignParent = oDesignChild.GetParent
    
    If Not oDesignParent Is Nothing Then
        'Get Run, which is parent of Part
        Set oDesignChild = oDesignParent
        Set oDesignParent = oDesignChild.GetParent
            
        If Not oDesignParent Is Nothing Then
            'Get Pipeline, which is parent of Run
            Set oDesignChild = oDesignParent
            Set oDesignParent = oDesignChild.GetParent
            
            If Not oDesignParent Is Nothing Then
                'Add Pipeline to elements collection
                Call IJNameRule_GetNamingParents.Add(oDesignParent)
            End If
        End If
    End If
    
    'Release object references
    Set oDesignChild = Nothing
    Set oDesignParent = Nothing
    
    Exit Function
ERR_SUB:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
    
End Function
