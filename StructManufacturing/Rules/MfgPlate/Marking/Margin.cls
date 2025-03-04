VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Margin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Bending Control Line Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      KONI     June 12 2002    created
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.Margin"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    Const METHOD = "Margin: IJDMfgSystemMarkingRule_CreateAfterUnfold"
    On Error Resume Next
    
    Dim oGeomCol2d As IJMfgGeomCol2d
    Dim pResMgr As IUnknown
    Dim oWB As IJWireBody
    Dim oSystemMark As IJMfgSystemMark
    Dim oObjSystemMark As IUnknown
    Dim oMarkingInfo As MarkingInfo
    Dim oMfgPart As IJMfgPlatePart
    
    Set pResMgr = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part
    
    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    Set oGeomCol2d = oMfgPlateWrapper.GetGeometriesAfterUnfold
    
    If oGeomCol2d Is Nothing Then
        'Since there is nothing to be marked we can exit the function
        GoTo CleanUp
    End If
    
    Dim i As Long, j As Long
    Dim oGeom2d As IJMfgGeom2d
    Dim oNewGeom2d As IJMfgGeom2d
    Dim eGeomType As StrMfgGeometryType
    Dim oMoniker As IMoniker
    Dim oObject As IUnknown
    Dim oPort As IJPort
    Dim oStructPort As IJStructPortEx
    Dim oPrimaryStructPort As IJStructPort
    Dim oMfgDefCol As Collection
    Dim oConstMargin As IJConstMargin
    Dim oObliqueMargin As IJObliqueMargin
    
    'Dim oCustomMargin as ??? not implemented
    Dim lIndexCount As Long
    lIndexCount = 0
            
    Dim oGeomCol2dOut As IJMfgGeomCol2d
    Set oGeomCol2dOut = m_oGeom2dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    For i = 1 To oGeomCol2d.GetCount
        Set oGeom2d = oGeomCol2d.GetGeometry(i)
        eGeomType = oGeom2d.GetGeometryType
        If eGeomType = STRMFG_OUTER_CONTOUR Or eGeomType = STRMFG_INNER_CONTOUR Then
            Set oMoniker = oGeom2d.GetMoniker
            If oMoniker Is Nothing Then GoTo NextMoniker
            Set oObject = m_oMfgRuleHelper.BindToObject(pResMgr, oMoniker)
            
            If Not TypeOf oObject Is IJPort Then GoTo NextMoniker
            
            Set oPort = oObject
			
            ' If the port doesn't belong to the input Plate, continue with next contour edge
            If (Not (oPort.Connectable Is Part)) Then GoTo NextMoniker
			
            Set oMfgDefCol = m_oMfgRuleHelper.GetMfgDefinitions(oPort)
            
            ' Search the Primary Port also.
            If oMfgDefCol.Count = 0 Then
                Set oStructPort = oPort
                Set oPrimaryStructPort = oStructPort.PrimaryPort
                Set oPort = oPrimaryStructPort
                Set oMfgDefCol = m_oMfgRuleHelper.GetMfgDefinitions(oPort)
            End If
            
            If oMfgDefCol.Count > 0 Then
                Dim lFabMargin As Double, lAssyMargin As Double, lCustomMargin As Double
                lFabMargin = 0
                lAssyMargin = 0
                lCustomMargin = 0
                
                For j = 1 To oMfgDefCol.Count
                    If TypeOf oMfgDefCol.Item(j) Is IJDFabMargin Then
                        Dim oFabMargin As IJDFabMargin
                        Set oFabMargin = oMfgDefCol.Item(j)
                    
                        If oFabMargin.GeometryChange = AsMargin Then ''1 = As Margin, 2 = As Shrinkage, 3 = As Reference
                            If oFabMargin.MarginGeometryMode = RegionalFabricationMargin Then
                                Dim indLocal As Long
                                Dim oOuterCntrCurve As IJCurve
                                Dim bRegMarginBelongToThisEdge As Boolean
                                Dim dOuterCntrStartX As Double, dOuterCntrStartY As Double, dOuterCntrStartZ As Double
                                Dim dOuterCntrEndX As Double, dOuterCntrEndY As Double, dOuterCntrEndZ As Double
                                
                                bRegMarginBelongToThisEdge = False
                                Set oOuterCntrCurve = oGeom2d
                                oOuterCntrCurve.EndPoints dOuterCntrStartX, dOuterCntrStartY, dOuterCntrStartZ, dOuterCntrEndX, dOuterCntrEndY, dOuterCntrEndZ
                                
                                For indLocal = 1 To oGeomCol2d.GetCount
                                    Dim oLocalGeom2d As IJMfgGeom2d
                                    Dim eLocalGeomType As StrMfgGeometryType
                                    
                                    Set oLocalGeom2d = oGeomCol2d.GetGeometry(indLocal)
                                    eLocalGeomType = oLocalGeom2d.GetGeometryType
                                    If eLocalGeomType = STRMFG_REGIONAL_MARGIN_REF_MARK Then
                                        Dim dDummy As Double, dStartMinDist As Double, dEndMinDist As Double
                                        Dim oThisMarginRefCurve As IJCurve
                                        Dim oTestEndPos As New DPosition
                                        
                                        Set oThisMarginRefCurve = oLocalGeom2d
                                        
                                        oTestEndPos.Set dOuterCntrStartX, dOuterCntrStartY, dOuterCntrStartZ
                                        oThisMarginRefCurve.DistanceBetween oTestEndPos, dStartMinDist, dDummy, dDummy, dDummy, dDummy, dDummy, dDummy
                                        
                                        oTestEndPos.Set dOuterCntrEndX, dOuterCntrEndY, dOuterCntrEndZ
                                        oThisMarginRefCurve.DistanceBetween oTestEndPos, dEndMinDist, dDummy, dDummy, dDummy, dDummy, dDummy, dDummy
                                        
                                        If ((dStartMinDist < 0.001) And (dEndMinDist < 0.001)) Then
                                            bRegMarginBelongToThisEdge = True
                                        End If
                                        Set oThisMarginRefCurve = Nothing
                                    End If
                                    Set oLocalGeom2d = Nothing
                                Next
                                Set oOuterCntrCurve = Nothing
                                If bRegMarginBelongToThisEdge Then
                                    If TypeOf oMfgDefCol.Item(j) Is IJObliqueMargin Then
                                        Set oObliqueMargin = oMfgDefCol.Item(j)
                                        If oObliqueMargin.EndValue > oObliqueMargin.StartValue Then
                                            lFabMargin = lFabMargin + oObliqueMargin.EndValue
                                        Else
                                            lFabMargin = lFabMargin + oObliqueMargin.StartValue
                                        End If
                                    ElseIf TypeOf oMfgDefCol.Item(j) Is IJConstMargin Then
                                        Set oConstMargin = oMfgDefCol.Item(j)
                                        lFabMargin = lFabMargin + oConstMargin.Value
                                    End If
                                End If
                            ElseIf TypeOf oMfgDefCol.Item(j) Is IJAssyMarginChild Then
                                Set oConstMargin = oMfgDefCol.Item(j)
                                lAssyMargin = lAssyMargin + oConstMargin.Value
                            ElseIf TypeOf oMfgDefCol.Item(j) Is IJObliqueMargin Then
                                Set oObliqueMargin = oMfgDefCol.Item(j)
                                If oObliqueMargin.EndValue > oObliqueMargin.StartValue Then
                                    lFabMargin = lFabMargin + oObliqueMargin.EndValue
                                Else
                                    lFabMargin = lFabMargin + oObliqueMargin.StartValue
                                End If
                            ElseIf TypeOf oMfgDefCol.Item(j) Is IJConstMargin Then
                                Set oConstMargin = oMfgDefCol.Item(j)
                                lFabMargin = lFabMargin + oConstMargin.Value
                            'ElseIf TypeOf oMfgDefCol.Item(j) Is ??? Then
                            End If
                        End If
                    End If
                Next j
                
                'Tr#44569 (Negative value for margin, is not reflected in part monitor)
                'when applied negative margin,this mark hangs in the air ,so don't
                'place margin mark in case of negative margin.
    
                If lAssyMargin > 0 Or lFabMargin > 0 Or lCustomMargin > 0 Then
                     'Create a SystemMark object to store additional information
                     Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        
                     'Set the marking side
                     oSystemMark.SetMarkingSide UpSide
        
                     'QI for the MarkingInfo object on the SystemMark
                     Set oMarkingInfo = oSystemMark
                
                     oMarkingInfo.Name = "MARGIN"
                     oMarkingInfo.MaxAssyMarginValue = lAssyMargin
                     oMarkingInfo.MaxCustomMarginValue = lCustomMargin
                     oMarkingInfo.MaxFabMarginValue = lFabMargin
        
                     Set oNewGeom2d = oGeom2d.Clone
                     oNewGeom2d.PutGeometrytype STRMFG_MARGIN_MARK
                     Set oObjSystemMark = oSystemMark
                     
                     oSystemMark.Set2dGeometry oNewGeom2d
                     
                     lIndexCount = lIndexCount + 1
                     oGeomCol2dOut.AddGeometry lIndexCount, oNewGeom2d
                End If
            End If
        End If
        
NextMoniker:
        
        Set oNewGeom2d = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom2d = Nothing
        Set oMoniker = Nothing
        Set oObject = Nothing
        Set oPort = Nothing
        Set oMfgDefCol = Nothing
        Set oConstMargin = Nothing
        Set oObliqueMargin = Nothing
        Set oStructPort = Nothing
        Set oPrimaryStructPort = Nothing
    Next i
    
    Set IJDMfgSystemMarkingRule_CreateAfterUnfold = oGeomCol2dOut
    
CleanUp:
    Set oGeomCol2d = Nothing
    Set pResMgr = Nothing
    Set oWB = Nothing

    Set oObjSystemMark = Nothing
    Set oMfgPart = Nothing
    Set oPlateWrapper = Nothing
    Set oMfgPlateWrapper = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1013, , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d

End Function


 
