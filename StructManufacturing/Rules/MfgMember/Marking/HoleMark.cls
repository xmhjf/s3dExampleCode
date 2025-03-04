VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileHole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberMarking
' Module: Marking Rules
'
' Description:  This Class Creates hole reference marks(cross or '+') on Members which have Sketch Features or struct openings on it
'
' Author:
'
' Comments:
'*******************************************************************************
Option Explicit
'To be desired after cycle 3 fixes

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgMemberMarking.ProfileHole"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    Helpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    Helpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "EndConnectionLogical: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    
    On Error GoTo ErrorHandler
    
    'Create the SD profile Wrapper and initialize it
    Dim oSDMemberWrapper As New StructDetailObjects.MemberPart
    'Create the SD profile Wrapper and initialize it
    Set oSDMemberWrapper = New StructDetailObjects.MemberPart
    Set oSDMemberWrapper.object = Part
    
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    Dim oMfgProfilePart As IJMfgProfilePart
    Dim oMfgProfileOutput As IJMfgProfileOutput
        
    If Not oProfileWrapper.ProfileHasMfgPart(oMfgProfilePart) Then Exit Function
    Set oMfgProfileOutput = oMfgProfilePart

    'Getting the Active Faces of the Member to be Manufactured
    Dim oFaces() As Variant
    Dim varFace As Variant
    
    oMfgProfileOutput.GetPartFaces StructMfgAnySide, StructMfgAsAny, oFaces
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    If bSelectiveRecompute Then
        Exit Function
    End If
       
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Set oSDPartSupport.Part = Part
    
    Dim iCount As Integer
    Dim oClosestPosition1 As IJDPosition, oClosestPosition2 As IJDPosition
    Dim dMinDistance As Double
    Dim oFeatureModelBody As IJDModelBody
    Dim oFacesModelBody As IJDModelBody
    Dim oGeom3dCustom As IJMfgGeom3d 'Initialized in the case of APS Marks
    
    'Logic for IJDStructCutoutTool
    Dim oRelationHelper As IMSRelation.DRelationHelper
    Dim oCollection As IMSRelation.DCollectionHelper
       
    'Getting the Struct Opening Via Relation Helper
    Set oRelationHelper = Part
    Set oCollection = oRelationHelper.CollectionRelations("IJDesignParent", "FromCustomToDesignChildRln_ORIG")
        
    Dim eSubPort As JXSEC_CODE
    Dim bMarkOnBothFlanges As Boolean
    
    Dim oProcessedFeature As IJElements
    Set oProcessedFeature = New JObjectCollection
                
    If oCollection.Count > 0 Then
        If UBound(oFaces) >= LBound(oFaces) Then
            For Each varFace In oFaces
               For iCount = 1 To oCollection.Count
                    Dim oStructCutOutTool As IJDStructCutoutTool
                    Set oStructCutOutTool = oCollection.Item(iCount)
                   
                    'Getting the Wirebody from the Opening
                    Dim oContourWireBody As IJWireBody
                    Set oContourWireBody = oStructCutOutTool.InputContour
                
                    Dim oContourModelBody As IJDModelBody
                                    
                    Set oContourModelBody = oContourWireBody
                    Set oFacesModelBody = oSDMemberWrapper.SubPort(varFace).Geometry
                    
                    'Finding the Minimum Distance between the Active Face and Feature Geomentry
                    oContourModelBody.GetMinimumDistance oFacesModelBody, oClosestPosition1, oClosestPosition2, dMinDistance
                    
                    If dMinDistance < 0.001 Then
                        
                        If oProcessedFeature.Contains(oCollection.Item(iCount)) Then
                            GoTo NextOpening
                        Else
                            oProcessedFeature.Add oCollection.Item(iCount)
                        End If
                        
                        'Checking Whether the Opening is on the Flanges
                        eSubPort = varFace
                        If Not (eSubPort = JXSEC_WEB_LEFT) And Not (eSubPort = JXSEC_WEB_RIGHT) Then
                                                                                            
                            'For Getting the Modeling type of Opening 'ThroughAll','ThroughNext','UserDefienedDepth'
                            Dim oAttributes As IJDAttributes
                            Dim oAttibute As IJDAttribute
                            Dim oAssocRelation As IJDAssocRelation
                            Dim oColOfRelations As Object
                            Dim oTargetObjectCol As IJDTargetObjectCol
    
                            Set oAssocRelation = oCollection.Item(iCount)
                            Set oColOfRelations = oAssocRelation.CollectionRelations("{A4794C84-AF14-4AE3-93BC-8B3A91A2B934}", "OpeningTypeRln_ORIG")
                            Set oTargetObjectCol = oColOfRelations
                            Set oAttributes = oTargetObjectCol.Item(1)
                            Set oAttibute = oAttributes.CollectionOfAttributes("{8BACE11F-2B12-42A2-BC06-B9CBAC29DAB6}").Item(2)  'Attribute Value = 2 Cutting Limit
                            
                            Dim iCuttingType As Integer
                            iCuttingType = oAttibute.Value 'Attribut Value '1=UserDefined'  '2=ThroughAll' '3=ThroughNext'
                                                            
                            If iCuttingType = 2 Then
                                bMarkOnBothFlanges = True
                            End If
                        End If
                        
                        'Create Cross Marks and add to oGeomCol3d
                        CreateMemberHoleMarks Part, oSDMemberWrapper.SubPort(varFace), CLng(varFace), oContourWireBody, UpSide, oCollection.Item(iCount), bMarkOnBothFlanges, oGeom3dCustom, oGeomCol3d
                        
                    End If
NextOpening:
                Next iCount
            Next varFace
        End If
    End If
                       
                   
    'Logic for IJSkectchFeatures
                       
    'Getting the All Sketch Features as a Collection
    Dim oSketchFeatureCollection As Collection
    oSDPartSupport.GetFeatures oSketchFeatureCollection
        
    If oSketchFeatureCollection.Count <= 0 Then Exit Function
        
    If UBound(oFaces) >= LBound(oFaces) Then
        For Each varFace In oFaces
            For iCount = 1 To oSketchFeatureCollection.Count
                If Not TypeOf oSketchFeatureCollection.Item(iCount) Is IJSketchFeature Then
                    GoTo NextFeature
                End If
                             
                Set oFeatureModelBody = oSketchFeatureCollection.Item(iCount)
                Set oFacesModelBody = oSDMemberWrapper.SubPort(varFace).Geometry
                                                
                'Finding the Minimum Distance between the Sketch Feature and Active Faces
                oFeatureModelBody.GetMinimumDistance oFacesModelBody, oClosestPosition1, oClosestPosition2, dMinDistance
                
                If dMinDistance < 0.001 Then
                    
                    'Checking Whether the Feature is on the Flanges
                    eSubPort = varFace
                    If Not (eSubPort = JXSEC_WEB_LEFT) Or (eSubPort = JXSEC_WEB_RIGHT) Then
                       bMarkOnBothFlanges = True
                    End If
                    
                    Dim oFeatureWireBody As IJWireBody
                    Set oFeatureWireBody = oSketchFeatureCollection.Item(iCount)
                    
                    'Create Cross Marks and add to oGeomCol3d
                    CreateMemberHoleMarks Part, oSDMemberWrapper.SubPort(varFace), CLng(varFace), oFeatureWireBody, UpSide, oSketchFeatureCollection.Item(iCount), bMarkOnBothFlanges, oGeom3dCustom, oGeomCol3d
                    
                End If
NextFeature:
            Next iCount
        Next
    End If
    
    'Return collection of ML's geometry
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oSDMemberWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oMfgProfileOutput = Nothing
    'Set oFaces() = Nothing
    Set oSDPartSupport = Nothing
    Set oRelationHelper = Nothing
    Set oCollection = Nothing
    Set oContourModelBody = Nothing
    Set oFacesModelBody = Nothing
    Set oFeatureModelBody = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2000, , "RULES")
    
End Function

 


