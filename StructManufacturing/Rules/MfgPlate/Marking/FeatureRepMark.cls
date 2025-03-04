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
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Feature representation Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      StructMfg     December1st 2010    created
'
'******************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule
Const MODULE = "MfgPlateMarking.FeatureRepMark"
Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    Const METHOD = "FeatureRepMark: IJDMfgProcessRule2_ExecuteBeforeUnfoldRule"
    On Error GoTo ErrorHandler
    
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1069, , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "FeatureRepMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Initialize
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim ogeomcol3dout As IJMfgGeomCol3d
    Set ogeomcol3dout = m_oGeom3dColFactory.Create(oResourceManager)
    
    CreateAPSMarkings STRMFG_FEATURE_LABEL_MARK, ReferenceObjColl, ogeomcol3dout
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = ogeomcol3dout
    
    If bSelectiveRecompute Then
        Exit Function
    End If
    
    '*** Create the Profile Wrapper Object ***'
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oMfgPart As Object
    Dim oMfgPlatePart As IJMfgPlatePart
    
    If TypeOf Part Is IJPlatePart Then
        Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
        Set oPlateWrapper.object = Part
        
        '*** Get the Manufactured Plate Part ***'
        If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
            Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
            Set oMfgPlateWrapper.object = oMfgPart
        Else
            Exit Function
        End If
    End If
    
    Dim oGeomCol3d   As IJMfgGeomCol3d
    Dim oMfgParent   As IJMfgGeomParent
    Dim oChildColl   As IJDTargetObjectCol
    Dim oMfgMGHelper As IJMfgMGHelper
    Dim oMoniker     As IMoniker
    
    Set oMfgParent = oMfgPart
    Set oMfgPlatePart = oMfgPart
        Dim oMfgPlateCreation_AE As IJMfgPlateCreation_AE
        Set oMfgPlateCreation_AE = oMfgPart.ActiveEntity
    Set oGeomCol3d = oMfgPlateCreation_AE.GeometriesBeforeUnfold

    If oGeomCol3d Is Nothing Then
        'Since there is nothing to be marked you can exit the function after cleanup
        GoTo CleanUp
    End If
    
    Dim oPOM As IJDPOM
    Dim oObject As IJDObject
    Set oObject = Part
    Set oPOM = oObject.ResourceManager
    
    Dim oGeom3d As IJMfgGeom3D
    Dim eGeomType As StrMfgGeometryType
    Dim ii As Long
    Dim lCount As Long
    lCount = 1
    
    For ii = 1 To oGeomCol3d.GetCount
        Set oGeom3d = oGeomCol3d.GetGeometry(ii)
        eGeomType = oGeom3d.GetGeometryType
        
        If eGeomType = STRMFG_FEATURE_MARK Then
            Set oMoniker = oGeom3d.GetMoniker
            Dim oFeatureHelper As IJMfgFeatureSymbolHelper
            Dim oFeature As IJStructFeature
            Set oFeature = oPOM.GetObject(oMoniker)
            'oFeature.
            Dim strSymbolPath As String
            Dim oTransMat As IJDT4x4
            Set oFeatureHelper = New MfgFeatSymbolServ
            Set oFeatureHelper.FeatureObject = oFeature
            Dim oFeatureInputNamesCol As JCmnShp_CollectionAlias
            Dim oFeatureInputValuesCol As JCmnShp_CollectionAlias
            Dim oFeatureOutputNamesCol As JCmnShp_CollectionAlias
            Dim oFeatureOutputValuesCol As JCmnShp_CollectionAlias
            Set oFeatureInputNamesCol = Nothing
            Set oFeatureInputValuesCol = Nothing
            Set oFeatureOutputNamesCol = Nothing
            Set oFeatureOutputValuesCol = Nothing
            Dim strRepName As String
            strRepName = ""
            If strRepName = "" Then
                On Error Resume Next
                strRepName = Get_CustomAttributeInfo(oFeature, "IJProductionMarking", "MarkingRep")
    On Error GoTo ErrorHandler
            End If
            
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
'                                ogeomcol3dout.AddGeometry lCount, PlMrkHelpers.Create3DLineComplexString(oCS, STRMFG_FEATURE_LABEL_MARK, oGeom3d.FaceId, oMarkingInfo)
                                Dim oGeom3dLine As IJMfgGeom3D
                                Dim oSystemMark As IJMfgSystemMark
                                Set oGeom3dLine = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                                 
                                oGeom3dLine.PutGeometry oCS
                                'Create a SystemMark object to store additional information
                                Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                                'Set the marking side
                                oSystemMark.SetMarkingSide UpSide
                                Set oMarkingInfo = oSystemMark
                                oMarkingInfo.Name = ""
                        
                                oGeom3dLine.PutGeometrytype STRMFG_FEATURE_LABEL_MARK
                        
                                oSystemMark.Set3dGeometry oGeom3dLine
                                
                                oGeomCol3d.AddGeometry 1, oGeom3dLine
                                lCount = lCount + 1
                                
                                If oFeature.get_StructFeatureType = SF_Slot Then
                                    oMarkingInfo.Name = Get_CustomAttributeInfo(oFeature, "IJProductionName", "ProductionName")
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    Next ii
    
CleanUp:
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = ogeomcol3dout

    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1069, , "RULES")
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

Dim icount As Long
Dim nInterface As Long
Dim oAttributeColl As CollectionProxy
Dim oInterfaceInfo As IJDInterfaceInfo
Dim oAttributeInfo As IJDAttributeInfo
Dim oInterfaceAttributeCollection As IJDInfosCol
Dim oInterfaceSqlAttributeCollection As IJDInfosCol
    icount = 0
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
            icount = icount + 1
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
                        If oAttributeInfo.Name = strAttributeName Then
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


