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

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgMemberMarking.Margin"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    Helpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    Helpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    Const METHOD = "Margin: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim oGeomCol2d As IJMfgGeomCol2d
    Dim pResMgr As IUnknown
    Dim oWB As IJWireBody
    Dim oCS As IJComplexString
    Dim oSystemMark As IJMfgSystemMark
    Dim oObjSystemMark As IUnknown
    Dim oMarkingInfo As MarkingInfo
    Dim oMfgPart As IJMfgProfilePart
    
    Set pResMgr = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    Dim oMfgProfileWrapper As MfgRuleHelpers.MfgProfilePartHlpr
    If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
        Set oMfgProfileWrapper = New MfgRuleHelpers.MfgProfilePartHlpr
        Set oMfgProfileWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    Set oGeomCol2d = oMfgProfileWrapper.GetGeometriesAfterUnfold
    
    Dim i As Long, j As Long
    Dim oGeom2d As IJMfgGeom2d
    Dim oNewGeom2d As IJMfgGeom2d
    Dim eGeomType As StrMfgGeometryType
    Dim oMoniker As IMoniker, oPortMoniker As IMoniker
    Dim oObject As IUnknown
    Dim oPort As IJPort
    Dim oMfgDefCol As Collection
    Dim oConstMargin As IJConstMargin
    Dim lGeomFaceId As Long
            
    Dim oGeomCol2dOut As IJMfgGeomCol2d
    Set oGeomCol2dOut = m_oGeomCol2dFactory.Create(pResMgr)
    
    If Not oGeomCol2d Is Nothing Then
        'Since there is a manufacturing object related we should investigate
        For i = 1 To oGeomCol2d.Getcount
            Set oGeom2d = oGeomCol2d.GetGeometry(i)
            eGeomType = oGeom2d.GetGeometryType
            lGeomFaceId = oGeom2d.FaceId
            If eGeomType = STRMFG_OUTER_CONTOUR Then
                Set oMoniker = oGeom2d.GetMoniker
                
                ' Make sure the rule does not crash even though
                ' there are no monikers !!
                If oMoniker Is Nothing Then GoTo NextMoniker
                
                Set oObject = m_oMfgRuleHelper.BindToObject(pResMgr, oMoniker)
                
                If oObject Is Nothing Then GoTo NextMoniker
                
                If Not TypeOf oObject Is IJPort Then GoTo NextMoniker
                            
                Set oPort = oObject
                Set oMfgDefCol = m_oMfgRuleHelper.GetMfgDefinitions(oPort)
                
                If oMfgDefCol.Count > 0 Then
                    Dim lFabMargin As Double, lAssyMargin As Double, lCustomMargin As Double, lClampMargin As Double, lCoamingMargin As Double
                    lFabMargin = 0
                    lAssyMargin = 0
                    lCustomMargin = 0
                    lClampMargin = 0
                    lCoamingMargin = 0
                    
                    For j = 1 To oMfgDefCol.Count
                        If TypeOf oMfgDefCol.Item(j) Is IJDFabMargin Then
                            Dim oFabMargin As IJDFabMargin
                            Set oFabMargin = oMfgDefCol.Item(j)
                        
                            If oFabMargin.GeometryChange = AsMargin Then ''1 = As Margin, 2 = As Shrinkage, 3 = As Reference
                                If oFabMargin.MarginGeometryMode = RegionalFabricationMargin Then
                                    Dim indLocal As Long
                                    Dim oMarginMoniker As New JDMoniker
                                    Dim bRegMarginBelongToThisEdge As Boolean
                                    bRegMarginBelongToThisEdge = False
                                    
                                    oMarginMoniker.JMoniker = m_oMfgRuleHelper.GetMoniker(oFabMargin)
                                    
                                    For indLocal = 1 To oGeomCol2d.Getcount
                                        Dim oLocalGeom2d As IJMfgGeom2d
                                        Dim eLocalGeomType As StrMfgGeometryType
                                        
                                        Set oLocalGeom2d = oGeomCol2d.GetGeometry(indLocal)
                                        eLocalGeomType = oLocalGeom2d.GetGeometryType
                                        If eLocalGeomType = STRMFG_REGIONAL_MARGIN_REF_MARK Then
                                            Dim oStoredMarginMoniker As New JDMoniker
                                            oStoredMarginMoniker.JMoniker = oLocalGeom2d.GetMoniker
                                            If (oMarginMoniker.IsEqual(oStoredMarginMoniker)) Then
                                                If (oLocalGeom2d.GetProfileCell = oGeom2d.GetProfileCell) Then
                                                    bRegMarginBelongToThisEdge = True
                                                End If
                                            End If
                                            Set oStoredMarginMoniker = Nothing
                                        End If
                                        Set oLocalGeom2d = Nothing
                                    Next
                                    Set oMarginMoniker = Nothing
                                    If bRegMarginBelongToThisEdge Then
                                        Set oConstMargin = oMfgDefCol.Item(j)
                                        lFabMargin = lFabMargin + oConstMargin.Value
                                    End If
                                ElseIf TypeOf oMfgDefCol.Item(j) Is IJAssyMarginChild Then
                                    Set oConstMargin = oMfgDefCol.Item(j)
                                    
                                    If oConstMargin.Type = 7 Then ' ClampMargin
                                        lClampMargin = lClampMargin + oConstMargin.Value
                                    Else
                                        If oConstMargin.Type = 8 Then 'CoamingMargin
                                            lCoamingMargin = lCoamingMargin + oConstMargin.Value
                                        Else
                                            lAssyMargin = lAssyMargin + oConstMargin.Value
                                        End If
                                    End If
                                    
                                ElseIf TypeOf oMfgDefCol.Item(j) Is IJConstMargin Then
                                    Set oConstMargin = oMfgDefCol.Item(j)
                                    lFabMargin = lFabMargin + oConstMargin.Value
                                'ElseIf TypeOf oMfgDefCol.Item(j) Is ??? Then
                                End If
                            End If
                        End If
                    Next j
                    
                    lAssyMargin = lAssyMargin + lClampMargin + lCoamingMargin
        
                    If Abs(lAssyMargin) > 0 Or Abs(lFabMargin) > 0 Or Abs(lCustomMargin) > 0 Then
                    
                        'Create a SystemMark object to store additional information
                        Set oSystemMark = m_oSystemMarkFactory.Create(pResMgr)
            
                        'Set the marking side
                        oSystemMark.SetMarkingSide lGeomFaceId
            
                        'QI for the MarkingInfo object on the SystemMark
                        Set oMarkingInfo = oSystemMark
                    
                        oMarkingInfo.name = "MARGIN"
                        oMarkingInfo.MaxAssyMarginValue = lAssyMargin
                        oMarkingInfo.MaxCustomMarginValue = lCustomMargin
                        oMarkingInfo.MaxFabMarginValue = lFabMargin
            
                        Set oNewGeom2d = oGeom2d.Clone
                        oNewGeom2d.PutGeometrytype STRMFG_MARGIN_MARK
                        oNewGeom2d.FaceId = lGeomFaceId
                        
                        oNewGeom2d.PutMoniker oMoniker
                        
                        oSystemMark.Set2dGeometry oNewGeom2d
                        oGeomCol2dOut.AddGeometry 1, oNewGeom2d
                    End If
                End If
            End If
            
NextMoniker:
            
            Set oNewGeom2d = Nothing
            Set oSystemMark = Nothing
            Set oMarkingInfo = Nothing
            Set oGeom2d = Nothing
            Set oMoniker = Nothing
            Set oPortMoniker = Nothing
            Set oObject = Nothing
            Set oPort = Nothing
            Set oMfgDefCol = Nothing
            Set oConstMargin = Nothing
        Next i
    End If
    
    Set IJDMfgSystemMarkingRule_CreateAfterUnfold = oGeomCol2dOut
    
CleanUp:
    Set oGeomCol2d = Nothing
    Set pResMgr = Nothing
    Set oWB = Nothing
    Set oCS = Nothing
    Set oObjSystemMark = Nothing
    Set oMfgPart = Nothing
    Set oProfileWrapper = Nothing
    Set oMfgProfileWrapper = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2009, , "RULES")
    GoTo CleanUp

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d

End Function

 
