VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PipeRunNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************************
'  Copyright (C) 2004-2005, Intergraph Corporation.  All rights reserved.
'
'  Project  : K:\CommonRoute\Rules\RouteRunNameRules
'
'  Class    : PipeRunNameRule
'
'  Abstract : The file contains implementation of the naming rules for PipeRuns.
'
'      25-Oct-04        Harini       Author
'      22-Feb-2005      Bharath      Removed the dependancy on the client tier component
'                               i.e. K:\CommonRoute\Client\Bin\RteErrorUtils.dll.
'                               included the reference to X:\Shared\Bin\ErrorLog.dll.
'      10-Mar-2005      Pavan   TR-CP·72719  Changing Electrical System name deletes Seq # from Cableway,
'                                            the same thing is observed in Piping too
'      19-Apr-2005      Pavan   TR-CP·72713  The Process Unit name only get added to a Pipe Run in certain situation.
'      27-May-2005      Furno   TR 77654    To get the fluid code, use the metadata method directly instead of
'                                           the method on the utility interface since the metadata method will return
'                                           a valid "<undefined>" string instead of an error message that cannot be used
'***************************************************************************************

Option Explicit

Implements IJNameRule
Private m_oErrors As IJEditErrors ' Middle tier errors
Private Const E_FAIL = -2147467259
Private Const MODULE = "CRunNamingRule"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"
Private Const CodeList_FluidCode = "FluidCode"

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
    Dim oPipeline             As IJPipelineSystem
    Dim oParentSystem         As IJSystem
    Dim oPipeRun              As IJRtePipeRun
    Dim oNameCounter          As IJNameCounter
    Dim oNamedItem            As IJNamedItem
    Dim oDesignChild          As IJDesignChild
    Dim oPipeSpec             As IJDPipeSpec
    Dim oSpec As IJDSpec
    Dim strName               As String
    Dim strUnitSystem         As String
    Dim strPipeSpec           As String
    Dim strFluidCode          As String
    Dim strNPD                As String
    Dim strSystemName         As String
    Dim strRunName            As String
    Dim strLocation           As String
    Dim strSeqNo              As String
    Dim strSystemParent       As String
    Dim lCount                As Long
    Dim strValidateName As String
    Dim oMetaData As IJDCodeListMetaData
    'Dim strNameBasis          As String
    If oObject Is Nothing Then GoTo ErrHndler
    'Get the Parent of the PipeRun
    Set oParentSystem = oParents.Item(1)
    
    Set oNamedItem = oParentSystem
    strSystemName = oNamedItem.Name
    Set oNamedItem = Nothing

    Set oNamedItem = oObject
    strRunName = oNamedItem.Name
    Set oNamedItem = Nothing

    'No parents - use just the strSystemName as the base
    'strNameBasis = oActiveEntity.NamingParentsString
    'Check if the new parent name string constructed and the old parent name string
    'existing are the same.  If they are, we do not need to generate a new name
    'Get the Parent of the PipeLine
    Set oPipeline = oParentSystem

    'Get NPD from PipeRun
    Set oPipeRun = oObject
    strNPD = CStr(oPipeRun.NPD)
    
    'Get the Fluid Code from the PipeLine
    Set oMetaData = oPipeRun
    strFluidCode = oMetaData.ShortStringValue(CodeList_FluidCode, oPipeline.FluidCode)
    Set oMetaData = Nothing
    Set oPipeline = Nothing
    
    'Get Pipe Spec from PipeRun
    Set oPipeSpec = oPipeRun.GetUsedSpec
    Set oSpec = oPipeSpec
    strPipeSpec = oSpec.DisplayName
    Set oSpec = Nothing
    Set oPipeRun = Nothing
    'To get the UnitSystem
    Do While TypeOf oParentSystem Is IJDesignChild
        If TypeOf oParentSystem Is IJUnitSystem Then
            'Get the Unit system string.
            Set oNamedItem = oParentSystem
            strUnitSystem = oNamedItem.Name
            Set oNamedItem = Nothing
            ' Presently we are Displaying the Name of the Unitsystem but if we want to change
            ' it to display UnitCode of the Unit system then use the below code
            ' Set oUnitsystem = oParentSystem
            ' strUnitSystem = oUnitsystem.UnitCode
            Exit Do
        Else
            Set oDesignChild = Nothing
            Set oDesignChild = oParentSystem
            Set oParentSystem = Nothing
            Set oParentSystem = oDesignChild.GetParent
        End If
    Loop
    
    Set oDesignChild = Nothing
    Set oParentSystem = Nothing
    
    On Error GoTo ErrHndler
    
    If Len(strUnitSystem) > 0 Then
        strValidateName = strUnitSystem + "-" + strNPD + "-" + strFluidCode + "-"
    Else
        strValidateName = strNPD + "-" + strFluidCode + "-"
    End If
    Dim arr() As String
    Dim intUpperBound As Integer
    Dim intLowerBound As Integer
    Dim strOldPipeSpec As String
    Dim strOldSeqNo As String
    Dim RunNameLength As Integer
    Dim Checklength As Integer
    Dim strOldName As String
    arr() = Split(strRunName, "-", , vbTextCompare)
    intUpperBound = UBound(arr)
    If intUpperBound > 0 Then
        strOldPipeSpec = arr(intUpperBound)
        intLowerBound = intUpperBound - 1
        strOldSeqNo = arr(intLowerBound)
        strOldName = strRunName
        RunNameLength = Len(Trim(strOldName))
        Checklength = Len(Trim(strOldPipeSpec)) + Len(Trim(strOldSeqNo))
        'Deleting "1" more from RunNameLength to Include "-" also .
        Checklength = RunNameLength - Checklength - 1
        strOldName = Left$(strOldName, Checklength)
    End If
    'We Compare the NewString (strUnitSystem + "-" + strNPD + "-" + strFluidCode + "-" )and Oldstring from the RunName and
    'if they are Different we generate a New Name.we also Compare whethere the PipeRunSpec has Changed or Not.
    If Not (StrComp(strOldName, strValidateName, vbTextCompare) = 0) Or Not StrComp(strOldPipeSpec, strPipeSpec, vbTextCompare) = 0 Then
         'Get the middle context
        GetMiddleContext oModelResourceMgr
        
        'Increment and append a counter.
        Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
        strLocation = vbNullString
        lCount = oNameCounter.GetCountEx(oModelResourceMgr, strSystemName, strLocation)
        strSeqNo = Format(lCount, strCountFormat)
    
        If Len(strUnitSystem) > 0 Then
            strName = strUnitSystem + "-" + strNPD + "-" + strFluidCode + "-" + strSeqNo + "-" + strPipeSpec
        Else
            strName = strNPD + "-" + strFluidCode + "-" + strSeqNo + "-" + strPipeSpec
        End If
        oActiveEntity.NamingParentsString = strName
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



