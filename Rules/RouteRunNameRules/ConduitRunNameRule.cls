VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConduitRunNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************************
'  Copyright (C) 2004-2005, Intergraph Corporation.  All rights reserved.
'
'  Project  : K:\CommonRoute\Rules\RouteRunNameRules
'
'  Class    : ConduitRunNameRule
'
'  Abstract : The file contains implementation of the naming rules for ConduitRuns.
'
'
'      25-Oct-04        Harini       Author
'      22-Feb-2005      Bharath      Removed the dependancy on the client tier component
'                               i.e. K:\CommonRoute\Client\Bin\RteErrorUtils.dll.
'                               included the reference to X:\Shared\Bin\ErrorLog.dll.
'      10-Mar-2005      Pavan   TR-CP·72719  Changing Electrical System name deletes Seq # from Cableway,
'                                            the same thing is observed in Conduit too
'***************************************************************************************

Option Explicit

Implements IJNameRule

Private m_oErrors As IJEditErrors ' Middle tier errors
Private Const E_FAIL = -2147467259
Private Const MODULE = "CRunNamingRule"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"
Private Const CONDUIT = "EC"
Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub
Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
'  Creates a name for the object passed in. The name is based on the parents
'  name and object name.It is assumed that all Naming Parents and the Object
'  implement IJPRodModelItem.The Naming Parents are added in AddNamingParents()
'  of the same interface. Both these methods are called from naming rule semantic.
'
' Arguments:
'  oObject       - Input.  Child object that needs to have the naming rule naming.
'  oParents      - Input.  Naming parents collection.
'  oActiveEntity - Input.  Naming rules active entity on which the NamingParentsString is stored.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal oParents As IJElements, ByVal oActiveEntity As Object)
Const METHOD = "IJNameRule_ComputeName"
On Error GoTo ErrHndler
     Dim oModelResourceMgr     As IUnknown
    Dim oNameCounter          As IJNameCounter
    Dim oNamedItem            As IJNamedItem
    Dim strName               As String
    Dim strNameBasis          As String
    Dim strLocation           As String
    Dim strSeqNo              As String
    Dim strSystemParent       As String
    Dim lCount                As Long

    If oObject Is Nothing Then GoTo ErrHndler
    
    'To Get the System Parent Name
    Set oNamedItem = oParents.Item(1)
    strSystemParent = oNamedItem.Name
    Set oNamedItem = Nothing
    
    strNameBasis = oActiveEntity.NamingParentsString
    
    If strSystemParent <> strNameBasis Then
        'Get the middle context
        GetMiddleContext oModelResourceMgr
    
        'Check if the new parent name string constructed and the old parent name string
        'existing are the same.  If they are, we do not need to generate a new name
        oActiveEntity.NamingParentsString = strSystemParent
        'Increment and append a counter.
        Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
        strLocation = vbNullString
        lCount = oNameCounter.GetCountEx(oModelResourceMgr, strSystemParent, strLocation)
        strSeqNo = Format(lCount, strCountFormat)
        
        'Need to Include Service Level
        strName = strSystemParent + "-" + CONDUIT + "-" + "-" + strSeqNo
        
        'Setting the Name on the Object
        Set oNamedItem = oObject
        oNamedItem.Name = strName
    End If
    'Clean up
    Set oNameCounter = Nothing
    Set oNamedItem = Nothing
    
Exit Sub
    
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description:
' Initial design:
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parents added here are used in computing the name of the object in
' ComputeName() of the same interface. Both these methods are called from naming rule semantic.
'
' Arguments:
'  oEntity - Input.  Child object that needs to have the naming rule naming.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
Const METHOD = "IJNameRule_GetNamingParents"
On Error GoTo ErrHndler

    Dim oParent As IJSystem
    Dim oChild As IJSystemChild

    On Error Resume Next
    'Getting the Parent of the Object
    Set oChild = oEntity
    Set oParent = oChild.GetParent
    
    On Error GoTo ErrHndler

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oParent Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oParent)
    End If

    'Set to Nothing
    Set oParent = Nothing
    Set oChild = Nothing

Exit Function
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Private Sub GetMiddleContext(oModelResourceMgr As IUnknown)
    Dim jContext              As IJContext
    Dim oDBTypeConfig         As IJDBTypeConfiguration
    Dim oConnectMiddle        As IJDAccessMiddle
    Dim strModelDBID          As String
    
    Const METHOD = "GetMiddleContext"
    On Error GoTo ErrHndler
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    'Clean up
    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
Exit Sub
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Sub



