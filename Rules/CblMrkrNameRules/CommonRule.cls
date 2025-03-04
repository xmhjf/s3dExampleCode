VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'Copyright (C) 2004, Intergraph Corporation. All rights reserved.                            '
'                                                                                            '
'Abstract:                                                                                   '
'    Common Naming Rule for Cable Marker                                                     '
'                                                                                            '
'Description:                                                                                '
'    This naming rule will name the marker after the parent to which the feature holding     '
'    the marker is related.  This will be either a conduit run or a cableway.                '
'                                                                                            '
'           Marker Name = RunName & TypeName & Counter                                       '
'                   where,                                                                   '
'                                                                                            '
'                       RunName is the name of the cableway or conduit run                   '
'                       TypeName is the user name for the cable marker object class          '
'                       Counter is a numberic counter with format '0000'                     '
'                                                                                            '
'Notes:                                                                                      '
'                                                                                            '
'History                                                                                     '
'                                                                                            '
'    lew            01/22/2004      Created                                                  '
'                                                                                            '
'*********************************************************************************************
Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "CommonRule: "
Private Const DEFAULTBASE = "Defaultname"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "000"

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    Sub:           IJNameRule_GetNamingParents                                        '
'                                                                                      '
'    Type:          Function                                                           '
'                                                                                      '
'    Description:   Part of IJNameRule implementation.  This function will return the  '
'                     parent object of the cable marker such that the NamedEntity      '
'                     relationship can be created between the marker and it's naming   '
'                     parent.  The relationship will, subsequently, be used to main-   '
'                     the proper name on the marker in the event the name of the       '
'                     run is modified.                                                 '
'                                                                                      '
'    Inputs:        oEntity     The item being named (expected to be a cable marker).  '
'                                                                                      '
'    OutPuts:       IJElements  Collection containing parent object                    '
'                                                                                      '
'                                                                                      '
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Dim oPathFeat As IJRtePathFeat
    Dim oPathMarker As IJRteMarker
    Dim oPathRun As IJRtePathRun
    
    On Error Resume Next
    
'--- Access the Features Parent, the Cableway or Conduit Run

    Set oPathMarker = oEntity
    Set oPathFeat = oPathMarker.GetMarkFeature
    Set oPathRun = oPathFeat.GetPathRun
    
    On Error GoTo ErrorHandler

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

    If Not (oPathRun Is Nothing) Then
        Call IJNameRule_GetNamingParents.Add(oPathRun)
    End If

    Set oPathRun = Nothing
    Set oPathFeat = Nothing
    Set oPathMarker = Nothing
    
Exit Function

ErrorHandler:
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    Sub:           IJNameRule_ComputeName                                             '
'                                                                                      '
'    Type:          Sub                                                                '
'                                                                                      '
'    Description:   Part of IJNameRule implementation.  This function will return the  '
'                     parent object of the cable marker such that the NamedEntity      '
'                     relationship can be created between the marker and it's naming   '
'                     parent.  The relationship will, subsequently, be used to main-   '
'                     the proper name on the marker in the event the name of the       '
'                     run is modified.                                                 '
'                                                                                      '
'    Inputs:        oObject         The object for which the name is to be updated.    '
'                                                                                      '
'                   elements        The collection containing a pointer to the parent  '
'                                    object.                                           '
'                                                                                      '
'                   oActiveEntity   The naming rule entity.  It is used to determine   '
'                                    if the name of an object's naming parent has      '
'                                    changed.                                          '
'    OutPuts:       None                                                               '
'                                                                                      '
'                                                                                      '
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
    
    Dim jContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    Dim strLocation As String
    
'--- Secure the Middle Context

    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)

    Dim oNameCounter As IJNameCounter
    Dim oChildNamedItem As IJNamedItem
    Dim oParent As Object
    Dim oParentNamedItem As IJNamedItem

    Dim strChildName As String
    Dim strParentName As String
    Dim strNamedParentsString As String
    Dim nCount As Long
    Dim strBaseName As String

    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService

'    Set oChildNamedItem = oObject
'
'    strBaseName = oChildNamedItem.TypeString
'
'    If strBaseName = "" Then
        strBaseName = "CableMarker"  ' <== Please Localize
'    End If
    
'--- Using the Type Name for the Base Portion of the Name
    
    strBaseName = Join(Split(strBaseName), "")

    If elements.Count > 0 Then
        For Each oParent In elements
            On Error Resume Next
            Set oParentNamedItem = oParent
            On Error GoTo ErrorHandler

            strParentName = oParentNamedItem.Name

            If (Len(strChildName) = 0) Then
                strChildName = strParentName
            Else
                strChildName = strChildName + "-" + strParentName
            End If

            Set oParentNamedItem = Nothing
        Next oParent

        strNamedParentsString = oActiveEntity.NamingParentsString


    '--- Comparison of the Current Parent Name String with That of the Old Parent
        
        If (strChildName & "-" & strBaseName) <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strChildName & "-" & strBaseName
        
            nCount = oNameCounter.GetCountEx(oModelResourceMgr, strChildName & "-" & strBaseName, strLocation)
            
            strChildName = strChildName + "-" + strBaseName + "-" + Format(nCount, strCountFormat)
            Set oChildNamedItem = oObject
            oChildNamedItem.Name = strChildName
        End If
    Else
        
    '--- When no Parents are Found, the Base Portion of the Name is Used By Itself
        
        strNamedParentsString = oActiveEntity.NamingParentsString
        
    '--- Comparison of the Current Parent Name String with That of the Old Parent
        
        If strBaseName <> strNamedParentsString Then
            oActiveEntity.NamingParentsString = strBaseName

            nCount = oNameCounter.GetCount(oModelResourceMgr, strBaseName)

            strChildName = strBaseName + "-" + Format(nCount, strCountFormat)
            oChildNamedItem.Name = strChildName
        End If
    End If
   
    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    
    Set oNameCounter = Nothing
    Set oChildNamedItem = Nothing
    Set oParent = Nothing
    Set oParentNamedItem = Nothing
    
Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
End Sub
