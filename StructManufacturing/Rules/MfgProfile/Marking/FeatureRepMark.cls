VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FeatureRepMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2010 Intergraph Corporation  All rights reserved.
'
'  Project: MfgProfileMarking
'
'  Abstract:    Rule for creating the Feature Representation Markings in the MfgProfile command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      StructMfg     December1st 2010    created
'
'******************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule
Const MODULE = "MfgProfileMarking.FeatureRepMark"
Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    Const METHOD = "FeatureRepMark: IJDMfgProcessRule2_ExecuteBeforeUnfoldRule"
    On Error GoTo ErrorHandler
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2050, , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "FeatureRepMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Initialize
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oGeomCol3dOut As IJMfgGeomCol3d
    Set oGeomCol3dOut = m_oGeomCol3dFactory.Create(oResourceManager)
    
    CreateAPSProfileMarkings STRMFG_FEATURE_LABEL_MARK, ReferenceObjColl, oGeomCol3dOut
    
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3dOut
    
    
    '*** Create the Profile Wrapper Object ***'
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Dim oMfgProfileWrapper As MfgRuleHelpers.MfgProfilePartHlpr
    Dim oMfgPart As Object
    
    If TypeOf Part Is IJProfilePart Then
        Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
        Set oProfileWrapper.object = Part
        
        '*** Get the Manufactured Profile Part ***'
        If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
            Set oMfgProfileWrapper = New MfgRuleHelpers.MfgProfilePartHlpr
            Set oMfgProfileWrapper.object = oMfgPart
        Else
            Exit Function
        End If
    End If
    
    Dim oMfgProfilePart As IJMfgProfilePart
    Set oMfgProfilePart = oMfgPart
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = oMfgProfilePart.GeometriesBeforeUnfold

    If oGeomCol3d Is Nothing Then
        'Since there is nothing to be marked you can exit the function after cleanup
        GoTo CleanUp
    End If
    
    Dim oObject As IJDObject
    Set oObject = Part
    
    Dim oPOM As IJDPOM
    Set oPOM = oObject.ResourceManager
    
    Dim lCount As Long
    lCount = 1
    
    Dim ii As Long
    For ii = 1 To oGeomCol3d.Getcount
        Dim oGeom3d As IJMfgGeom3d
        Set oGeom3d = oGeomCol3d.GetGeometry(ii)
        
        Dim eGeomType As StrMfgGeometryType
        eGeomType = oGeom3d.GetGeometryType
        
        If eGeomType = STRMFG_FEATURE_MARK Then
            Dim oMoniker As IMoniker
            Set oMoniker = oGeom3d.GetMoniker
            
            Dim oBoundMoniker As Object
            Set oBoundMoniker = oPOM.GetObject(oMoniker)
            
            Dim oFeature As IJStructFeature
            If TypeOf oBoundMoniker Is IJStructFeature Then
                Set oFeature = oBoundMoniker
            Else
                Set oMoniker = Nothing
                Set oBoundMoniker = Nothing
                GoTo NextII
            End If
            
            Dim eNewMarkGeomType As StrMfgGeometryType
            Dim eFeatureType As StructFeatureTypes

            eFeatureType = oFeature.get_StructFeatureType
            If eFeatureType = SF_WebCut Or eFeatureType = SF_FlangeCut Then
                eNewMarkGeomType = STRMFG_ENDCUT_MARK
            Else
                eNewMarkGeomType = STRMFG_FEATURE_LABEL_MARK
            End If
            
            Dim oFeatureHelper As IJMfgFeatureSymbolHelper
            Set oFeatureHelper = New MfgFeatSymbolServ
            Set oFeatureHelper.FeatureObject = oFeature
            
            Dim strRepName As String
            
            strRepName = ""
            If eFeatureType = SF_WebCut Then
                 If oGeom3d.GetProfileCell = Web Then
                    strRepName = Get_CustomAttributeInfo(oFeature, "IJProductionMarking", "MarkingRep")
                End If
            ElseIf eFeatureType = SF_FlangeCut Then
                If oGeom3d.GetProfileCell = TopFlange Or oGeom3d.GetProfileCell = BottomFlange Then
                    strRepName = Get_CustomAttributeInfo(oFeature, "IJProductionMarking", "MarkingRep")
                End If
            Else
                On Error Resume Next
                strRepName = Get_CustomAttributeInfo(oFeature, "IJProductionMarking", "MarkingRep")
                On Error GoTo ErrorHandler
            End If
            
            If Not strRepName = "" Then
                Dim strSymbolPath As String
                Dim oTransMat As IJDT4x4
                Dim oFeatureInputNamesCol As JCmnShp_CollectionAlias
                Dim oFeatureInputValuesCol As JCmnShp_CollectionAlias
                Dim oFeatureOutputNamesCol As JCmnShp_CollectionAlias
                Dim oFeatureOutputValuesCol As JCmnShp_CollectionAlias
                Set oFeatureInputNamesCol = Nothing
                Set oFeatureInputValuesCol = Nothing
                Set oFeatureOutputNamesCol = Nothing
                Set oFeatureOutputValuesCol = Nothing
            
                strSymbolPath = oFeatureHelper.FeatureSymbolRelativePath
                oFeatureHelper.FeatureSymbolInputs oFeatureInputNamesCol, oFeatureInputValuesCol
                oFeatureHelper.EvaluateSymbol strSymbolPath, oFeatureInputNamesCol, oFeatureInputValuesCol
                oFeatureHelper.GetRepresentationOutputs strRepName, oTransMat, oFeatureOutputNamesCol, oFeatureOutputValuesCol
                
                If Not oFeatureOutputNamesCol Is Nothing Then
                    Dim lOutCount As Long
                    Dim i As Long
                    Dim varValue As Variant
                    Dim oGeom As Object
                    Dim oWB As IJWireBody
                    Dim oCS As IJComplexString
                    lOutCount = oFeatureOutputNamesCol.Count
                    Dim oMarkingInfo As MarkingInfo
        
                    For i = 1 To lOutCount
                        Set varValue = Nothing
                        On Error Resume Next
                        Set varValue = oFeatureOutputValuesCol.Item(i)
                        On Error GoTo ErrorHandler
                        Set oGeom = Nothing
                        Set oGeom = varValue
                        If Not oGeom Is Nothing Then
                            If Not TypeOf oGeom Is IJPoint Then
                                Set oCS = New ComplexString3d
                                oCS.AddCurve oGeom, True
    
                                If Not oCS Is Nothing Then
                                    oGeomCol3dOut.AddGeometry lCount, PrMrkHelpers.CreateGeom3dObject(oCS, eNewMarkGeomType, oGeom3d.FaceId, oMarkingInfo)
                                    If oFeature.get_StructFeatureType = SF_Slot Then
                                        oMarkingInfo.name = Get_CustomAttributeInfo(oFeature, "IJProductionName", "ProductionName")
                                    End If
                                    lCount = lCount + 1
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
        End If
NextII:
    Next ii

CleanUp:
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3dOut

    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2050, , "RULES")
    GoTo CleanUp
End Function

'********************************************************************
' Routine: Set_CustomAttributeInfoObject
'
' Description:
'       Sets a the Attribute collection for a specific Virtual interface
'       and then returns the attribute.
' Inputs:
'       oObject - Object to retrieve Custom Attributes from
'       strUserInterfaceName As String - The interface that has the attribute
'       strAttributeName As String - the name of the attribute
'       setValue As String - The value to be written to the attribute name
' Outputs:
'       None
'********************************************************************
Private Function Get_CustomAttributeInfo(oObject As IJDObject, _
                                         strUserInterfaceName As String, _
                                         strAttributeName As String) As String
On Error GoTo ErrorHandler
Const MT = "Get_CustomAttributeInfoObject"

Dim InterfaceID As Variant
Dim oAttributes As IJDAttributes

Dim iCount As Long
Dim nInterface As Long
Dim oAttributeColl As CollectionProxy
Dim oInterfaceInfo As IJDInterfaceInfo
Dim oAttributeInfo As IJDAttributeInfo
Dim oInterfaceAttributeCollection As IJDInfosCol
Dim oInterfaceSqlAttributeCollection As IJDInfosCol
    iCount = 0
    nInterface = 0
    Set oAttributes = oObject
    Dim bOK As Boolean
    bOK = True
    Get_CustomAttributeInfo = ""

    ' Verify that the Attributes contain at least one
    On Error Resume Next
    If oAttributes.Count > 0 Then
        ' Loop thru each Interface in the Attributes Collection
        nInterface = oAttributes.Count
        For Each InterfaceID In oAttributes
            iCount = iCount + 1
            Set oAttributeColl = oAttributes.CollectionOfAttributes(InterfaceID)

            ' verify the current interface Collection is valid
            If Not oAttributeColl Is Nothing Then
                
                If oAttributeColl.InterfaceInfo.IsHardCoded = False Then

                ' verify that the current Attribute Interface collection
                ' represents a "User" Attribute interface not System Attribute(??)
                    
                    If oAttributeColl.InterfaceInfo.UserName = strUserInterfaceName Then
                    Dim oAttribute As IJDAttribute
                    For Each oAttribute In oAttributeColl
                        Set oAttributeInfo = oAttribute.AttributeInfo
                        If oAttributeInfo.name = strAttributeName Then
                            Get_CustomAttributeInfo = oAttribute.Value
                            Exit Function
                        End If
                        Set oAttributeInfo = Nothing
                        Set oAttribute = Nothing
                     Next
                    End If
                
                End If 'is hard coded

                Set oAttributeColl = Nothing
            End If ' if oAttributeColl
            Set InterfaceID = Nothing
        Next
    End If 'oAttributes

    Set oAttributes = Nothing

Exit Function
    
ErrorHandler:
   Err.Raise LogError(Err, MODULE, MT).Number
End Function


