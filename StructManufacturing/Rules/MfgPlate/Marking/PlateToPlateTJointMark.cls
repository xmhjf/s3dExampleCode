VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateToPlateTJointMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Plate To Plate T-Joint Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      TBH        feb. 6. 2002    created
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.PlateToPlateTJointMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

' ***********************************************************************************
' Public Function IJDMfgSystemMarkingRule_CreateBeforeUnfold
'
' Description:  function creates fitting marks on the this plate for each T - connected plate.
'               Input arguments: this part, marking side.
'               Output argument: collection of marking lines.
'
'
' Unresolved issue: GetConnectedTypeForContour() not exposed yet
' ***********************************************************************************
Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "PlateToPlateTJointMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    'Prepare the output collection of marking line's geometries
    Dim oGeomCol3D As IJMfgGeomCol3d
    Set oGeomCol3D = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
 
    CreateAPSMarkings STRMFG_PLATE_TO_PLATE_TJOINT_MARK, ReferenceObjColl, oGeomCol3D
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3D
    
    'Prepare collections
    Dim oSDPlateWrapper As New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part
    
    Dim oMfgRuleHlprPlatePart As MfgRuleHelpers.PlatePartHlpr
    Set oMfgRuleHlprPlatePart = New MfgRuleHelpers.PlatePartHlpr
    Set oMfgRuleHlprPlatePart.object = Part
    
    Dim oMfgPart As IJMfgPlatePart
    Dim oNeutralSurface As IJSurfaceBody

    If oMfgRuleHlprPlatePart.PlateHasMfgPart(oMfgPart) Then
        Dim oMfgPlateCreation_AE As IJMfgPlateCreation_AE
        Set oMfgPlateCreation_AE = oMfgPart.ActiveEntity
        Set oNeutralSurface = oMfgPlateCreation_AE.NeutralSurface
        Set oMfgPlateCreation_AE = Nothing
    Else
        Exit Function
    End If
    
    '--------------------------
    ' Get the connected objects
    '--------------------------
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Set oSDPartSupport.Part = Part
  
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetPhysicalConnectionData(Part, ReferenceObjColl, True)
    
    Dim oConnectionData As ConnectionData
    Dim iConnObjIndex As Integer 'iteration object
    
    'iterate through the collection of plate parts connected to this plate
    For iConnObjIndex = 1 To oConObjsCol.Count
        'get connection data
        oConnectionData = oConObjsCol.Item(iConnObjIndex)
            
        If TypeOf oConnectionData.ToConnectable Is IJCollarPart Then
             GoTo NextItem
        End If
        
        If Not TypeOf oConnectionData.ToConnectable Is IJPlatePart Then GoTo NextItem
        
        Dim bIsCrossOfTee As Boolean
        Dim eConnType As ContourConnectionType
        oSDPartSupport.GetConnectionTypeForContour oConnectionData.AppConnection, _
                                                   eConnType, _
                                                   bIsCrossOfTee
                                                   
        'If this is not TEE type conection, continue to the next item
        If eConnType <> PARTSUPPORT_CONNTYPE_TEE Then GoTo NextItem
        
        ' Check if there is one lateral face and one base/offset face
        Dim res1 As Integer
        Dim res2 As Integer
        Dim port1Flag As IMSStructConnection.eUSER_CTX_FLAGS
        Dim port2Flag As IMSStructConnection.eUSER_CTX_FLAGS

        Dim oStructPort1 As IJStructPort, oStructPort2 As IJStructPort
        Set oStructPort1 = oConnectionData.ConnectingPort
        Set oStructPort2 = oConnectionData.ToConnectedPort

        port1Flag = oStructPort1.ContextID
        port2Flag = oStructPort2.ContextID

        res1 = port1Flag And IMSStructConnection.CTX_LATERAL
        res2 = port2Flag And IMSStructConnection.CTX_LATERAL

        If res1 > 0 And res2 > 0 Then GoTo NextItem 'Butt connections
        If Not res1 > 0 And Not res2 > 0 Then GoTo NextItem  'Atleast one should be lateral
    
        'TODO : Make sure lap-connections don't trigger this

        'Get collection of mark positions
        Dim oMarkPosColl As Collection
        On Error Resume Next
        Set oMarkPosColl = GetFittingMarkPositions(oSDPlateWrapper.object, oConnectionData, UpSide, oNeutralSurface)
        On Error GoTo ErrorHandler
    
        If oMarkPosColl Is Nothing Then GoTo NextItem
        'Iterate through the collection of mark points in oMarkPosColl
        Dim oMarkPosObj As IJDPosition 'iteration object
    
        'Iterate through collection of mark points and create ML for each of the items
        For Each oMarkPosObj In oMarkPosColl
    
            'Get connected ports on both parts
            Dim oThisPartPort As IJPort
            Dim oConnPartPort As IJPort
            Set oThisPartPort = oConnectionData.ConnectingPort
            Set oConnPartPort = oConnectionData.ToConnectedPort
    
            'Find lateral connected port
            Dim oThisPlateWrapper As New StructDetailObjects.PlatePart
            Set oThisPlateWrapper.object = Part
        
            Dim oLateralPort As IJPort
            Dim resConn As Integer
            Dim resThis As Integer
            Dim portConnFlag As IMSStructConnection.eUSER_CTX_FLAGS
            Dim portThisFlag As IMSStructConnection.eUSER_CTX_FLAGS
        
            'Get struct ports of connected plates
            Dim oStructThisPort As IJStructPort
            Dim oStructConnPort As IJStructPort
            Set oStructThisPort = oThisPartPort
            Set oStructConnPort = oConnPartPort
        
            'get port flag
            portConnFlag = oStructConnPort.ContextID
            portThisFlag = oStructThisPort.ContextID
            resConn = portConnFlag And IMSStructConnection.CTX_LATERAL
            resThis = portThisFlag And IMSStructConnection.CTX_LATERAL
        
            Dim oPlateConnByLateralPort As IJPlatePart
    
            If resConn > 0 And resThis <= 0 Then 'oStructConnPort is lateral:
                Set oLateralPort = oConnectionData.ToConnectedPort
                Set oPlateConnByLateralPort = oConnectionData.ToConnectable
            Else
                If resThis > 0 And resConn <= 0 Then 'oStructThisPort is lateral:
                    Set oLateralPort = oConnectionData.ConnectingPort
                    Set oPlateConnByLateralPort = Part
                Else
                    'Plate is not T - connected. No mark points for this case
                    Exit Function
                End If
            End If
    
            'Get surfaces of ports
            Dim oThisPartSuface As IUnknown
            Dim oConnPartSurface As IUnknown
    
    '***************** GET FACE OF THE PLATE CONNECTED BY LATERAL ------->
            Dim oConnPlateWrapper As New StructDetailObjects.PlatePart
            Set oConnPlateWrapper.object = oConnectionData.ToConnectable
            Dim oFacePortOfLateralPlate As IJPort
    
            'Case 1: Connected part is engaged by its lateral
            If oPlateConnByLateralPort Is oConnectionData.ToConnectable Then
                Set oFacePortOfLateralPlate = oConnPlateWrapper.BasePort(BPT_Offset)
                Set oThisPartSuface = oThisPartPort.Geometry
                Set oConnPartSurface = oFacePortOfLateralPlate.Geometry
    
            'Case 2: This part is engaged by its lateral
            Else
                'Set oFacePortOfLateralPlate = oSDPlateWrapper.BasePort(BPT_Offset)
                'Set oThisPartSuface = oFacePortOfLateralPlate.Geometry
                Set oThisPartSuface = oNeutralSurface
                Set oConnPartSurface = oConnPartPort.Geometry
            End If
    '***************** GET FACE OF THE PLATE CONNECTED BY LATERAL <-------
    
            'Create mark line through oMarkPosObj
            Dim oCS As IJComplexString
            Set oCS = CreateLocationFittingMark(oMarkPosObj, oThisPartSuface, oConnPartSurface)
            
            If oCS Is Nothing Then
                GoTo NextMarkPosObj
            End If
    
            'Prepare some different assisting objects
            Dim oMarkingInfo As MarkingInfo
            Dim oGeom3d As IJMfgGeom3D
            Dim oMoniker As IMoniker
            Dim oObjSystemMark As IUnknown
            Dim oSystemMark As IJMfgSystemMark
            
            Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
            Dim ePlateSide As enumPlateSide
            ePlateSide = oMfgRuleHlprPlatePart.GetSide(oThisPartPort)
            If (ePlateSide = UnDefinedSide) Then
                oSystemMark.SetMarkingSide UpSide
            Else
                oSystemMark.SetMarkingSide ePlateSide
            End If
            'QI for the MarkingInfo object on the SystemMark
            Set oMarkingInfo = oSystemMark
    
            'Set marking information, if any
            oMarkingInfo.Name = "PLATE TJOINT FITTING MARK"
    
            'Add ML's geometry to collection
            Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
            oGeom3d.PutGeometry oCS
            oGeom3d.PutGeometrytype STRMFG_PLATE_TO_PLATE_TJOINT_MARK
    
            Set oObjSystemMark = oSystemMark
            
            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
            oGeom3d.PutMoniker oMoniker
            
            oSystemMark.Set3dGeometry oGeom3d
            oGeomCol3d.AddGeometry 1, oGeom3d
    'next position
NextMarkPosObj:
            Set oGeom3d = Nothing
            Set oMoniker = Nothing
            Set oObjSystemMark = Nothing
            Set oMarkingInfo = Nothing
            Set oSystemMark = Nothing
            Set oCS = Nothing
            Set oThisPartSuface = Nothing
            Set oConnPartSurface = Nothing
            Set oConnPartPort = Nothing
            Set oThisPartPort = Nothing
            Set oMarkPosObj = Nothing
        Next oMarkPosObj
    
NextItem:
        Set oMarkPosColl = Nothing
    Next iConnObjIndex
    
CleanUp:
    Set oFacePortOfLateralPlate = Nothing
    Set oPlateConnByLateralPort = Nothing
    Set oLateralPort = Nothing
    Set oStructThisPort = Nothing
    Set oStructConnPort = Nothing
    Set oThisPlateWrapper = Nothing
    Set oMfgRuleHlprPlatePart = Nothing
    Set oSDPlateWrapper = Nothing
    Set oConObjsCol = Nothing
    Set oSDPartSupport = Nothing
    Set oNeutralSurface = Nothing
    
    'Return collection of ML's geometry
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1020, , "RULES")
    GoTo CleanUp
End Function


 
