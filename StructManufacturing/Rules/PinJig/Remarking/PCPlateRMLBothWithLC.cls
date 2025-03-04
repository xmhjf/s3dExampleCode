VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCPlateRMLBothWithLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
' Module: ConnectedRemarking
'
' Description:  Example to show the remarking lines due to plates
'               that are connected to the supported plates, on the pinjig
'
' Author: Himaja
'
'*******************************************************************************
Option Explicit
Private Const MODULE As String = "MfgPinJigRule.PhyConnPlateRemarking"

Implements IJDMfgPinJigRulePerRemarkingType



Private Function IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking(ByVal PartialPinJig As Object) As IJElements
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetEntitiesForRemarking"
    On Error GoTo ErrorHandler

    Dim oReturnCollection As IJElements
    Set oReturnCollection = New JObjectCollection
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig

    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates

    Dim PlateSide() As eUSER_CTX_FLAGS
    PlateSide = GetSidesOfPlatesFacingPlane(oPinJig, ConnectionPhysical, oPlateColl)
    
    Dim Iter As Integer
    ' For each supported plate ...
    For Iter = 1 To oPlateColl.Count
        Dim ConnectedObjColl As Collection
        Dim ConnectionsColl As Collection
        GetConnectedObjects oPlateColl.Item(Iter), _
                            PLATE_TYPE, ConnectionPhysical, _
                            PARTSUPPORT_CONNTYPE_TEE, _
                            PlateSide(Iter), _
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
    Set oPlateColl = Nothing
    Set oPinJig = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5018, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As String
    IJDMfgPinJigRulePerRemarkingType_GetLocateFilterForRemarkingEntities = "[MfgPinJigRule.RemarkingFilter,IsConnectedPlatePart]"
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry(ByVal PartialPinJig As Object, ByVal ElemsToRemark As IJElements) As IJMfgGeomCol3d
    Const METHOD = "IJDMfgPinJigRulePerRemarkingType_GetRemarkingGeometry"
    On Error GoTo ErrorHandler

    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PartSupport
       
    Dim ReturnColl As IJMfgGeomCol3d
    Set ReturnColl = CreateMfgGeomCol3dObject

    Dim oMfgGeomHelper As New MfgGeomHelper
    Dim oSDConPlateWrapper As New StructDetailObjects.PlatePart

    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig

    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates
    
    
    Dim RemarkingSideOfPlate() As eUSER_CTX_FLAGS
    RemarkingSideOfPlate = GetSidesOfPlatesFacingPlane(oPinJig, ConnectionPhysical, oPlateColl)

    ' ******************************* IMPORTANT NOTE ******************************* '
    ' The Pin jig semantic expects Connection Geometry (with proper attribution).    '
    ' So create a MfgGeom3d object (that has both geometry and attribution).         '
    ' The semantic will use these objects to create connection remarking lines and   '
    ' merge them based on MONIKERS and NAMES of these objects.                       '
    ' ******************************* IMPORTANT NOTE ******************************* '

    Dim Iter As Integer
    Dim i As Long
    ' For each supported plate ...
    For Iter = 1 To oPlateColl.Count

        Dim ConnectedObjColl As Collection
        Dim ConnectionsColl As Collection
        GetConnectedObjects oPlateColl.Item(Iter), _
                            PLATE_TYPE, _
                            ConnectionPhysical, _
                            PARTSUPPORT_CONNTYPE_TEE, _
                            RemarkingSideOfPlate(Iter), _
                            ConnectedObjColl, ConnectionsColl

       
        ' For each connected object ...
        For i = 1 To ConnectedObjColl.Count
            Dim oWireBody As IJWireBody
        ' ... check if this connected object is another supported plate
            If ElemsToRemark.Contains(ConnectedObjColl.Item(i)) And _
               Not oPlateColl.Contains(ConnectedObjColl.Item(i)) Then
               
        ' ... get the desired portion of the geometry (SideA/SideB)
                Set oSDConPlateWrapper.object = ConnectedObjColl.Item(i)
                
                Dim oPartInfo As IJDPartInfo
                Set oPartInfo = New PartInfo
                Dim eMoldedDir As StructMoldedDirection
                eMoldedDir = oPartInfo.GetPlatePartThicknessDirection(ConnectedObjColl.Item(i))
                               
                Dim sMoldedSide As String
                sMoldedSide = oSDConPlateWrapper.MoldedSide
                
                Dim eSideOfConnectedObjectToBeMarked As GSCADSDPartSupport.thicknessSide
                
               ' If eMoldedDir = Centered Then
               '   eSideOfConnectedObjectToBeMarked = SideUnspecified
              '  Else
                  If sMoldedSide = "Base" Then
                      eSideOfConnectedObjectToBeMarked = SideA
                  ElseIf sMoldedSide = "Offset" Then
                      eSideOfConnectedObjectToBeMarked = SideB
                  Else
                      eSideOfConnectedObjectToBeMarked = SideUnspecified
                  End If
                  
              '  End If
                
                Set oPartSupport.Part = oPlateColl.Item(Iter)
                Dim oThickDir As IJDVector
                On Error Resume Next
                oPartSupport.GetConnectionContour_Tee ConnectionsColl.Item(i), oWireBody, oThickDir, eSideOfConnectedObjectToBeMarked
                On Error GoTo ErrorHandler
                Set oThickDir = Nothing

            End If

            If Not oWireBody Is Nothing Then
                '************************************************
                'First RemarkingLine
                '**************************************************
                Dim oNI As IJNamedItem
                Dim RemarkingLineName As String
                Set oNI = ConnectedObjColl.Item(i)
                
                If Not oNI Is Nothing Then
                    RemarkingLineName = oNI.Name
                End If
                RemarkingLineName = RemarkingLineName & " _OtherSide1"
                
                Dim oCSColl As IJElements
                Set oCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWireBody)
                Dim oCS As IJComplexString
                
                ' RemarkingLine
                For Each oCS In oCSColl
                    CreateMfgGeom3dObject oCS, STRMFG_PinJigRemarkingLine3D, _
                                      ConnectedObjColl.Item(i), ReturnColl, _
                                      RemarkingLineName, STRMFG_PinJig_Remarking_Plate
                Next
                
                '*************************************************
                'Second RemarkingLine
                '*************************************************
                If eSideOfConnectedObjectToBeMarked = SideA Then
                    eSideOfConnectedObjectToBeMarked = SideB
                Else
                    eSideOfConnectedObjectToBeMarked = SideA
                End If
                
                Set oWireBody = Nothing
                Set oThickDir = Nothing
                On Error Resume Next
                oPartSupport.GetConnectionContour_Tee ConnectionsColl.Item(i), oWireBody, oThickDir, eSideOfConnectedObjectToBeMarked
                On Error GoTo ErrorHandler
                Set oThickDir = Nothing

                ' IMPORTANT: Give a name DIFFERENT from above, here (e.g., suffix with "OtherSide" etc.)
                If Not oNI Is Nothing Then
                    RemarkingLineName = oNI.Name
                End If
                RemarkingLineName = RemarkingLineName & " _OtherSide2"
                If Not oWireBody Is Nothing Then
                    Set oCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWireBody)
                    ' RemarkingLine
                    For Each oCS In oCSColl
                        CreateMfgGeom3dObject oCS, STRMFG_PinJigRemarkingLine3D, _
                                      ConnectedObjColl.Item(i), ReturnColl, _
                                      RemarkingLineName, STRMFG_PinJig_Remarking_Plate
                    Next
                End If
                Set oNI = Nothing
                Set oWireBody = Nothing
                Set oCS = Nothing
                Set oCSColl = Nothing
                    
            End If ' end check for valid geometry
            
                '*************************************************
                    'Center RemarkingLine
                '*************************************************
            Dim oPlate As IJDPlateMoldedConventions
            Dim ochild As IJSystemChild
            Dim oParent As IJSystemChild
            
            Set ochild = ConnectedObjColl.Item(i)
            If Not TypeOf ConnectedObjColl.Item(i) Is IJCollarPart Then 'Exclude collar plates
                Set oPlate = ochild.GetParent ' Leaf Plate System
                
                If eMoldedDir = Centered Or oPlate.plateThicknessOffset <> 0 Then
                    Dim ConnectedObj As Object
    
                    Dim oSuperParent As Object
                    Set ochild = ConnectionsColl.Item(i) ' Physical Connection
                    Set oParent = ochild.GetParent 'Assembly Connection
                    Set oSuperParent = oParent.GetParent ' Logical Connection
                    
                    Set oWireBody = oSuperParent
                    Set ConnectedObj = ConnectedObjColl.Item(i)
                    
                    If Not oWireBody Is Nothing Then
                      
                        Set oNI = ConnectedObjColl.Item(i)
                        
                        If Not oNI Is Nothing Then
                            RemarkingLineName = oNI.Name
                        End If
                       
                        Set oCSColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oWireBody)
                     
                        For Each oCS In oCSColl
                            CreateMfgGeom3dObject oCS, STRMFG_PinJigRemarkingLine3D, _
                                              ConnectedObj, ReturnColl, _
                                              RemarkingLineName, STRMFG_PinJig_Remarking_Plate
                        Next
                    
                    End If
                End If
            End If
           
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
    Set oPlateColl = Nothing
    Set ReturnColl = Nothing
    Set oPartSupport = Nothing
    Set oPinJig = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5019, , "RULES")
End Function

Private Function IJDMfgPinJigRulePerRemarkingType_GetRemarkingTypesSetByRule(ByVal PartialPinJig As Object) As Long()

End Function
