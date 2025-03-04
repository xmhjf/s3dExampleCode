VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PipeCompNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'********************************************************************
' Copyright (C) 2004-2006, Intergraph Corporation.  All Rights Reserved.
'
' File: PipeCompNameRule.cls
'
' Author: Pavan
'
' Abstract:
'      The file contains Implementation of a NamingRule for PipeComponents.
'      the Name for a PipeComponent is set as : ShortCode + "-" + SequenceNumber
'      if ShortCode is not there then TypeString is used to generate the name.
'      There were no naming parents to participate while generating the name for piping components.

'
' History:
'      18-Nov-2004   Pavan                 Created
'      22-Feb-2005   Bharath      Removed the dependancy on the client tier component
'                               i.e. K:\CommonRoute\Client\Bin\RteErrorUtils.dll.
'                               Included the reference to X:\Shared\Bin\ErrorLog.dll.
'       09-Aug-2006 Furno       TR 103424 Remove reference to client tier component Elements
'*************************************************************************************************************
Option Explicit

Implements IJNameRule
Private m_oErrors As IJEditErrors
Private Const E_FAIL = -2147467259
Private Const MODULE = "CPComponentNameRule: "
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name
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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''IJNameRule''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''/////////////////////////////////////////////////////////////////////////////
'' METHOD:         IJNameRule_ComputeName
' Description:
'  Creates a name for the object passed in. The Naming Parents are added in AddNamingParents()
'  of the same interface. Both these methods are called from naming rule semantic.
'
' Arguments:
'  oObject       - Input.  Child object that needs to have the NamingRule naming.
'  oParents      - Input.  Naming parents collection.
'  oActiveEntity - Input.  Naming rules active entity on which the NamingParentsString is stored.
'/////////////////////////////////////////////////////////////////////////////
Private Sub IJNameRule_ComputeName(ByVal oEntity As Object, ByVal oParents As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrHndler
    Dim oModelResourceMgr As IUnknown
    Dim oNameCounter As IJNameCounter
    Dim oNamedItem As IJNamedItem

    Dim strName As String
    Dim strNameBasis As String
    
    Dim lcount As Long
    Dim strLocation As String
    Dim strSequenceNumber As String

    Dim strNameSuffix As String
    Dim oDistribpart  As IJDistribPartOccur
    Dim oPathGenPart As IJRtePathGenPart
    Dim oFeatCollect As IJDTargetObjectCol
    Dim oPathfeat As IJRtePipePathFeat
    Dim oPartData As IJRtePartData
    
    If oEntity Is Nothing Then Exit Sub
       
    'To get the Engineering Tag
    Set oDistribpart = oEntity
    Set oPathGenPart = oDistribpart
    Set oNamedItem = oEntity
    If oPathGenPart.IsBasePart > 0 Then
        Set oFeatCollect = oPathGenPart.GetPathFeatures()
        Set oPathfeat = oFeatCollect.Item(1)
        strNameSuffix = oPathfeat.Tag
        'If No Engineering Tag then it is Replaced By ShortCode
        If Len(Trim(strNameSuffix)) <= 0 Then
            strNameSuffix = oPathfeat.GetShortCode
        End If
    Else
        Set oPartData = oPathGenPart
        If Not oPartData Is Nothing Then
            strNameSuffix = oPartData.ShortCode
        End If
    End If
    If Len(Trim(strNameSuffix)) <= 0 Then
        strNameSuffix = oNamedItem.TypeString
    End If
    strNameBasis = oActiveEntity.NamingParentsString
    Dim strOldNameBasis As String
    Dim arr() As String
    strOldNameBasis = strNameBasis
    If oNamedItem.Name <> vbNullString Then
        strOldNameBasis = vbNullString
        arr() = Split(oNamedItem.Name, "-", , vbTextCompare)
        strOldNameBasis = arr(0)
    End If
    If strNameSuffix <> strNameBasis Or strOldNameBasis <> strNameSuffix Then
        'Get the middle context
        GetMiddleContext oModelResourceMgr
        oActiveEntity.NamingParentsString = strNameSuffix
        'Increment and append a counter.
        strLocation = vbNullString
        'To get the SequenceNumber
        Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
        lcount = oNameCounter.GetCountEx(oModelResourceMgr, strNameSuffix, strLocation)
        strSequenceNumber = Format(lcount, strCountFormat)
        
        'Name is Formed in the Required Sequence
        strName = strNameSuffix + "-" + strSequenceNumber
        'Setting Name on the Object
        oNamedItem.Name = strName
    End If
    
    'Clean up
    Set oModelResourceMgr = Nothing
    Set oNameCounter = Nothing
    Set oNamedItem = Nothing
    Set oEntity = Nothing
Exit Sub
   
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.number, MODULE & METHOD, Err.Description
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
  Const METHOD = "IJNameRuleComp_GetNamingParents"
    On Error GoTo ErrHndler
'   get feature from part
    Dim oPathGenPart As IJRtePathGenPart
'   Dim lIsBasePart
    Dim oPathfeat As IJRtePathFeat
    Dim oPathFeatCol As IJDTargetObjectCol
    Dim oElements As IJElements
    Set oPathGenPart = oEntity
    If Not oPathGenPart Is Nothing Then
        Set oPathFeatCol = oPathGenPart.GetPathFeatures
        If oPathFeatCol.Count >= 1 Then
            Set oPathfeat = oPathFeatCol.Item(1)
        End If
    End If
    Set oElements = New IMSCoreCollections.JObjectCollection
    
    If Not oPathfeat Is Nothing Then
        oElements.Add oPathfeat
    End If
    Set IJNameRule_GetNamingParents = oElements
    Exit Function
    
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.number, MODULE & METHOD, Err.Description
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
    m_oErrors.Add Err.number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Sub
