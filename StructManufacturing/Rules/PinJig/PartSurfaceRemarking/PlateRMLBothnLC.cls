VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateRMLBothnLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigPartSurfLCRule
' Module: PlateRMLBothnLC
'
' Description:  Example to show the remarking lines due to plates
'               that are connected to the supported plates, on the pinjig
'
' Author: Himaja
'
'*******************************************************************************
Option Explicit
Private Const MODULE As String = "MfgPinJigPartSurfLCRule.LCPlateRMLBothAndLC"

Implements IJDMfgPinJigRulePerRemarkingType

Private Function IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking(ByVal PartialPinJig As Object) As IJElements
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking"
    On Error GoTo ErrorHandler

    Dim oReturnCollection As IJElements
    Set oReturnCollection = New JObjectCollection
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig

    Dim oPlateSysColl As IJElements
    Set oPlateSysColl = GetPinJigSupportedPlateSystems(oPinJig)

    Dim RemarkingSideOfPlate() As eUSER_CTX_FLAGS
    RemarkingSideOfPlate = GetSidesOfPlatesFacingPlane(oPinJig, ConnectionLogical)

    Dim Iter As Integer
    ' For each supported plate ...
    For Iter = 1 To oPlateSysColl.Count
        Dim ConnectedObjColl As Collection
        Dim ConnectionsColl As Collection
        GetConnectedObjects oPlateSysColl.Item(Iter), _
                            PLATE_TYPE, _
                            ConnectionLogical, _
                            PARTSUPPORT_CONNTYPE_TEE, _
                            RemarkingSideOfPlate(Iter), _
                            ConnectedObjColl, ConnectionsColl

        Dim i As Long
        ' For each connected object ...
        For i = 1 To ConnectedObjColl.Count
            oReturnCollection.Add ConnectedObjColl.Item(i)
        Next
    Next
    
    GetAPSMarkingLines oPinJig, STRMFG_PLATELOCATION_MARK, oReturnCollection
   
    Set IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking = oReturnCollection
    
    Set oReturnCollection = Nothing
    Set oPlateSysColl = Nothing
    Set oPinJig = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5018, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As String
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry"
    On Error GoTo ErrorHandler

    IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities = "[MfgPinJigRule.RemarkingFilter,IsConnectedPlatePart]"
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5018, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As IJMfgGeomCol3d
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry"
    On Error GoTo ErrorHandler
    
    Dim ReturnColl As IJMfgGeomCol3d
    Set ReturnColl = CreateMfgGeomCol3dObject

    Dim oMfgGeomHelper As New MfgGeomHelper
    Dim oSDConPlateWrapper As New StructDetailObjects.PlatePart
    Dim oPlate As IJDPlateMoldedConventions
    Dim oMfgMGHelper As New MfgMGHelper
       
    Dim oParent As IJSystem
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig

    Dim oPlateSysColl As IJElements
    Set oPlateSysColl = GetPinJigSupportedPlateSystems(oPinJig)
    
    Dim RemarkingSideOfPlate() As eUSER_CTX_FLAGS
    RemarkingSideOfPlate = GetSidesOfPlatesFacingPlane(oPinJig, ConnectionLogical)

    Dim oRemarkingSurface As IJSurfaceBody
    Set oRemarkingSurface = oPinJig.RemarkingSurface
    

    Dim oSDPlatePart As New StructDetailObjects.PlatePart
     
    ' ******************************* IMPORTANT NOTE ******************************* '
    ' The Pin jig semantic expects Connection Geometry (with proper attribution).    '
    ' So create a MfgGeom3d object (that has both geometry and attribution).         '
    ' The semantic will use these objects to create connection remarking lines and   '
    ' merge them based on MONIKERS and NAMES of these objects.                       '
    ' ******************************* IMPORTANT NOTE ******************************* '

    Dim i As Long
    Dim Iter As Integer
    ' For each supported plate ...
    For Iter = 1 To oPlateSysColl.Count

        Dim ConnectedObjColl As Collection
        Dim ConnectionsColl As Collection
        GetConnectedObjects oPlateSysColl.Item(Iter), _
                            PLATE_TYPE, ConnectionLogical, _
                            PARTSUPPORT_CONNTYPE_TEE, _
                            RemarkingSideOfPlate(Iter), _
                            ConnectedObjColl, ConnectionsColl
                        
        ' For each connected object ...
        For i = 1 To ConnectedObjColl.Count
        
            ' Retrieving detailed parts
            Dim oChildColl As IJDTargetObjectCol
            Dim iChild As Integer
            Dim oPlatePart As IJPlatePart
            Set oParent = ConnectedObjColl.Item(i)
            Set oChildColl = oParent.GetChildren
            
            For iChild = 1 To oChildColl.Count
                If TypeOf oChildColl.Item(iChild) Is IJPlatePart Then
                    Set oPlatePart = oChildColl.Item(iChild)
                    Set oSDPlatePart.object = oPlatePart
                End If
            Next
            
            Dim oWireBody As IJWireBody
              
            If ElemsToRemark.Contains(ConnectedObjColl.Item(i)) And _
                   LogConnIsTeeTypeConn(ConnectionsColl.Item(i), _
                                   ConnectedObjColl.Item(i), _
                                   oPlateSysColl.Item(Iter)) And _
              TypeOf ConnectionsColl.Item(i) Is IJWireBody _
            Then
               Set oWireBody = ConnectionsColl.Item(i)
            End If
            
            If Not oWireBody Is Nothing Then
                
                Dim oNI As IJNamedItem
                Set oNI = ConnectedObjColl.Item(i)
                Dim RemarkingLineName As String
                
                If Not oNI Is Nothing Then
                    ' Decorate your name here (e.g., prefix with "Deck" etc.)
                    RemarkingLineName = oNI.Name
                    ' Decorate your name here (e.g., suffix with "SideA" etc.)
                End If
                
                Set oPlate = ConnectedObjColl.Item(i)
                
                Dim eMoldedDir As StructMoldedDirection
                eMoldedDir = oPlate.plateThicknessDirection

                Set oSDConPlateWrapper.object = oPlatePart
                
                Dim sMoldedSide As String
                sMoldedSide = oSDConPlateWrapper.MoldedSide
                
                Dim oPartSupp As IJPartSupport
                Dim oPlatePartSupp As IJPlatePartSupport
                Set oPartSupp = New PlatePartSupport
                Set oPlatePartSupp = oPartSupp
                Set oPartSupp.Part = oPlatePart
                
                Dim dThickness As Double
                Dim dOffsetVal As Double
                oPlatePartSupp.GetThickness dThickness
                 
                Dim oSurfaceBody As IJSurfaceBody
                Set oSurfaceBody = ConnectedObjColl.Item(i)
                                          
                Dim oCSColl As IJElements
                Dim oCS As IJComplexString
                Set oCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWireBody)

                For Each oCS In oCSColl
               
                    Dim oDirOnToPlate   As IJDVector
                    Dim oResultCS       As IJComplexString
                    Dim oResultWire     As IJWireBody
                     
                    Set oDirOnToPlate = GetPlateOrientationVector(oCS, ConnectedObjColl.Item(i))
                    
                    ' Remarking surface is obtained only when the offset is > 0.001
                    If m_dRemarkingSurfaceOffset > 0.001 Then
                        ' Need to use the plate orientation for computing the projection vector
                        On Error Resume Next
                        oMfgMGHelper.ProjectComplexStringToSurface oCS, oRemarkingSurface, oDirOnToPlate, oResultCS
                        On Error GoTo ErrorHandler
                        
                        If oResultCS Is Nothing Then
                            StrMfgLogError Err, MODULE, METHOD, "Failed to project complexstring onto remarking surface", , , , "RULES"
                            Set oResultCS = oCS
                        End If
                
                    Else
                        Set oResultCS = oCS
                    End If
                    
                    '************************************************
                    'RemarkingLine from LC
                    '**************************************************
                    CreateMfgGeom3dObject oResultCS, STRMFG_PinJigRemarkingLine3D, _
                            ConnectedObjColl.Item(i), ReturnColl, _
                            RemarkingLineName, STRMFG_PinJig_Remarking_Plate
                            
                    ' If the Connected Plate is DECK, both remarking lines are created.
                    ' Else single remarking line representing LC is created.
                    ' Comment this "if" loop to create remarking lines for all connected plates irrespective of "PlateType".
                    If oSDPlatePart.plateType <> DeckPlate Then
                        GoTo NextConnectedPlate
                    End If
                    
                    oMfgMGHelper.ComplexStringToWireBody oResultCS, oResultWire
                         
                    'Check for the existence of offset
                    If eMoldedDir = Centered And oPlate.plateThicknessOffset = 0 Then ' When only Centered thickness
                    
                        '************************************************
                            'First RemarkingLine
                        '************************************************
                        Set oResultCS = Nothing
                        RemarkingLineName = RemarkingLineName & "_OtherSide1"
                                                                    
                        If sMoldedSide = "Base" Then
                            dOffsetVal = (dThickness / 2)
                        ElseIf sMoldedSide = "Offset" Then
                            dOffsetVal = -1 * (dThickness / 2)
                        End If
                        
                        Set oResultCS = GetOffsetCurve(oSurfaceBody, oResultWire, dOffsetVal, oRemarkingSurface)
                        
                        If Not oResultCS Is Nothing Then
                                                
                            CreateMfgGeom3dObject oResultCS, STRMFG_PinJigRemarkingLine3D, _
                                            ConnectedObjColl.Item(i), ReturnColl, _
                                            RemarkingLineName, STRMFG_PinJig_Remarking_Plate
                        End If
                                                                  
                        '************************************************
                            'Second RemarkingLine
                        '**************************************************
                        Set oResultCS = Nothing
                        RemarkingLineName = RemarkingLineName & "_OtherSide2"
                        
                        dOffsetVal = -1 * dOffsetVal
                        
                        Set oResultCS = GetOffsetCurve(oSurfaceBody, oResultWire, dOffsetVal, oRemarkingSurface)
                        
                        If Not oResultCS Is Nothing Then
                                                
                            CreateMfgGeom3dObject oResultCS, STRMFG_PinJigRemarkingLine3D, _
                                            ConnectedObjColl.Item(i), ReturnColl, _
                                            RemarkingLineName, STRMFG_PinJig_Remarking_Plate
                        End If
                         
                         

                    ElseIf oPlate.plateThicknessOffset <> 0 Then ' When there is Offset
                        
                        '************************************************
                            'First RemarkingLine
                        '************************************************
                        Set oResultCS = Nothing
                        RemarkingLineName = RemarkingLineName & "_OtherSide1"
                        
                        dOffsetVal = GetOffsetValue(oPlatePart, oResultWire, "Base")
                        
                        Set oResultCS = GetOffsetCurve(oSurfaceBody, oResultWire, dOffsetVal, oRemarkingSurface)
                        
                        If Not oResultCS Is Nothing Then
                                                
                            CreateMfgGeom3dObject oResultCS, STRMFG_PinJigRemarkingLine3D, _
                                            ConnectedObjColl.Item(i), ReturnColl, _
                                            RemarkingLineName, STRMFG_PinJig_Remarking_Plate
                        End If
                                                                  
                        '************************************************
                            'Second RemarkingLine
                        '**************************************************
                        Set oResultCS = Nothing
                        RemarkingLineName = RemarkingLineName & "_OtherSide2"
                        
                        dOffsetVal = GetOffsetValue(oPlatePart, oResultWire, "Offset")
                                                
                        Set oResultCS = GetOffsetCurve(oSurfaceBody, oResultWire, dOffsetVal, oRemarkingSurface)
                        
                        If Not oResultCS Is Nothing Then
                                                
                            CreateMfgGeom3dObject oResultCS, STRMFG_PinJigRemarkingLine3D, _
                                            ConnectedObjColl.Item(i), ReturnColl, _
                                            RemarkingLineName, STRMFG_PinJig_Remarking_Plate
                        End If
                    
                    
                    Else ' When NO Offset or non-Centered thickness
                                
                        '************************************************
                        'Other RemarkingLine
                        '**************************************************
                        Set oResultCS = Nothing
                        RemarkingLineName = RemarkingLineName & "_OtherSide"
                                                                    
                        If sMoldedSide = "Base" Then
                            dOffsetVal = dThickness
                        ElseIf sMoldedSide = "Offset" Then
                            dOffsetVal = -1 * dThickness
                        End If
                        
                        Set oResultCS = GetOffsetCurve(oSurfaceBody, oResultWire, dOffsetVal, oRemarkingSurface)
                        
                        If Not oResultCS Is Nothing Then
                                                
                            CreateMfgGeom3dObject oResultCS, STRMFG_PinJigRemarkingLine3D, _
                                            ConnectedObjColl.Item(i), ReturnColl, _
                                            RemarkingLineName, STRMFG_PinJig_Remarking_Plate
                        End If
                    
                        
                    End If
                                          
                 Next
                 
                 
                Set oMfgGeomHelper = Nothing
                Set oMfgMGHelper = Nothing
                Set oResultCS = Nothing
                Set oDirOnToPlate = Nothing
                Set oNI = Nothing
                Set oWireBody = Nothing
                Set oCS = Nothing
                Set oCSColl = Nothing
                
            End If ' end check for valid geometry
NextConnectedPlate:
        Next ' end looping around connected objects
    Next ' end looping around supported plates
    
    Dim PlateLocationMarks As IJElements
    Set PlateLocationMarks = New JObjectCollection
     
    GetAPSMarkingLines oPinJig, STRMFG_PLATELOCATION_MARK, PlateLocationMarks

    If PlateLocationMarks.Count > 0 Then
        CreateGeom3dFromAPSMarkingLines oPinJig, ElemsToRemark, PlateLocationMarks, STRMFG_PinJig_Remarking_Plate, ReturnColl
    End If
    Set PlateLocationMarks = Nothing
    
    Set IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry = ReturnColl

    Set oMfgGeomHelper = Nothing
    Set oSDConPlateWrapper = Nothing
    Set oPlateSysColl = Nothing
    Set ReturnColl = Nothing
    Set oPinJig = Nothing
    Set oParent = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5019, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingTypesSetByRule(ByVal PartialPinJig As Object) As Long()

End Function




