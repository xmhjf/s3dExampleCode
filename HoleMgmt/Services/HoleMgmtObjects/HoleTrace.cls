VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HoleTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
' Copyright (C) 1998-2006 Intergraph Corporation. All Rights Reserved.
'
' Project: S:\HoleMgmt\Data\HoleMgmtObjects
'
' File: HoleTrace.cls
'
' Author: Hole Mgmt Team
'
' Abstract: provides methods to manage hole trace objects
'
' Changes:
'   Date    By         Reason
' 06/02/03  CRS       Changed secondary orientation in PlaceEdgeReinforcement
'                     so that inside diameter of coaming coincides with hole.
'
' 06/18/03  CRS       Added sub SetProfilePartMaterialAndGrade and call it
'                     from PlaceProfileFitting to set material and grade of
'                     the newly created profile.
'
' 01/20/04  ASrivast  TR 50621 & 50622. Search for String TR 50621.
'                     Changes made to getOrientation Angle.
'                     Using the u&v vectors to get the correct alignment for
'                     the HoleFitting
'
' 04/26/06  Jealani   Changes have been made not to process holes on objects
'                     that do not support IJDStructApplyOperation. DI - 97082
'
' 06/13/06  PML       Updated for changes made in V7
'
'******************************************************************************

Option Explicit

Private m_oHoleTrace As IJHoleTraceAE

Private Const PI = 3.14159265

''  Constant for BO IID
Private Const strIJExtWeldPPInterfaceID = "{CA342C20-6433-4063-8A06-74C87D11BA7D}"
Private Const strCStandAlonePlatePart = "{E35693D3-E680-4F0B-BF83-ECBB28A45097}"

Private Const sSOURCEFILE As String = "HoleTrace.cls"
Private Const sDefRule As String = "HMBestFit.DefinitionRule"

'******************************************************************************
' METHOD:  Class_Terminate
'
' DESCRIPTION:  release all gloabl objects
'******************************************************************************
Private Sub Class_Terminate()
   Set m_oHoleTrace = Nothing
End Sub

'******************************************************************************
' METHOD: Object (Set)
'
' DESCRIPTION:  sets the internal hole trace object to the passed
'               hole trace from the calling method
'******************************************************************************
Public Property Set Object(ByVal oHoleTrace As Object)
    Set m_oHoleTrace = Nothing
    Set m_oHoleTrace = oHoleTrace
End Property

'******************************************************************************
' METHOD: Object (Get)
'
' DESCRIPTION:  return the internal hole trace object
'******************************************************************************
Public Property Get Object() As Object
    Set Object = m_oHoleTrace
End Property


'******************************************************************************
'******************************************************************************
'
' Start of methods used to get information about the hole trace
'
'******************************************************************************
'******************************************************************************


'******************************************************************************
' METHOD:  Name (Get)
'
' DESCRIPTION:  returns the hole trace's name
'******************************************************************************
Public Property Get Name() As String
    On Error GoTo ErrorHandler
    Const sMETHOD = "Name"
    
    'get IJNamedItem interface
    Dim oNamedItem As IJNamedItem
    Set oNamedItem = m_oHoleTrace
     
    'retrieve the name
    Name = oNamedItem.Name
     
Cleanup:
    Set oNamedItem = Nothing
    Exit Property

ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Property

'******************************************************************************
' METHOD:  PlateType (Get)
'
' DESCRIPTION:  return the plate type
'******************************************************************************
Public Property Get PlateType() As StructPlateType
    On Error GoTo ErrorHandler
    Const sMETHOD = "PlateType"
    
    Dim oStructure As Object
    Set oStructure = m_oHoleTrace.GetParentStructure
    If oStructure Is Nothing Then GoTo Cleanup
    
    If TypeOf oStructure Is IJPlate Then
        Dim oPlate As IJPlate
        Set oPlate = oStructure
        PlateType = oPlate.PlateType
    End If

Cleanup:
    Set oStructure = Nothing
    Set oPlate = Nothing
    Exit Property
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup:
End Property

'******************************************************************************
' METHOD:  PlateThickness (Get)
'
' DESCRIPTION:  return the plate thickness
'******************************************************************************
Public Property Get PlateThickness() As Double
    On Error GoTo ErrorHandler
    Const sMETHOD = "PlateThickness"
    
    Dim oStructure As Object
    Set oStructure = m_oHoleTrace.GetParentStructure
    If oStructure Is Nothing Then GoTo Cleanup
    
    If TypeOf oStructure Is IJPlate Then
        Dim oPlate As IJPlate
        Set oPlate = oStructure
        PlateThickness = oPlate.thickness
    End If
    
Cleanup:
    Set oStructure = Nothing
    Set oPlate = Nothing
    Exit Property
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup:
End Property

'******************************************************************************
' METHOD:  CuttingPlatePart (Get)
'
' DESCRIPTION:  return the plate object that is cut by the hole trace
'******************************************************************************
Public Property Get CuttingPlatePart() As Object
    On Error GoTo ErrorHandler
    Const sMETHOD = "CuttingPlatePart"

    Dim oPart As Object
    Dim oCollection As IJDObjectCollection

    Set oCollection = m_oHoleTrace.GetParentStructureParts
    For Each oPart In oCollection
        If Not (oPart Is Nothing) Then
            If TypeOf oPart Is IJPlate Then
                Set CuttingPlatePart = oPart
                Exit For
            End If
        End If
    Next

Cleanup:
    Set oCollection = Nothing
    Set oPart = Nothing

    Exit Property

ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Property

'******************************************************************************
' METHOD:  HoleFitting (Get)
'
' DESCRIPTION:  get the actual fitting for the hole trace
'******************************************************************************
Public Property Get HoleFitting() As Object
    On Error GoTo ErrorHandler
    Const sMETHOD = "HoleFitting"

    ' get the fitting active entity
    Dim oFittingAE As IJHoleFitting
    On Error Resume Next
    Set oFittingAE = m_oHoleTrace.GetHoleFitting
    On Error GoTo ErrorHandler
    If oFittingAE Is Nothing Then GoTo Cleanup
    
    Dim oFitting As Object
    Set oFitting = oFittingAE.HoleFitting
    If Not oFitting Is Nothing Then
        Set HoleFitting = oFitting
    End If
    
Cleanup:
    Set oFittingAE = Nothing
    Set oFitting = Nothing
    
    Exit Property
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Property

'******************************************************************************
' METHOD:  HoleFitting (Get)
'
' DESCRIPTION:  get the actual fitting for the hole trace
'******************************************************************************
Public Property Get FittingType() As HMFittingType
    On Error GoTo ErrorHandler
    Const sMETHOD = "FittingType"

    ' get the fitting active entity first
    Dim oFittingAE As IJHoleFitting
    On Error Resume Next
    Set oFittingAE = m_oHoleTrace.GetHoleFitting
    On Error GoTo ErrorHandler
    If oFittingAE Is Nothing Then GoTo Cleanup
    
    FittingType = oFittingAE.FittingType
    
Cleanup:
    Set oFittingAE = Nothing
    Exit Property
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Property

'******************************************************************************
' METHOD:  HoleMatrix (Get)
'
' DESCRIPTION:  compute the matrix of the hole trace
'******************************************************************************
Public Property Get HoleMatrix() As IJDT4x4
    On Error GoTo ErrorHandler
    Const sMETHOD = "HoleMatrix"
    
    Dim oPlane As IJPlane
    Dim oPosition As DPosition
    Dim oUVector As DVector
    Dim oVVector As DVector
    Dim oMatrix As IJDT4x4
    Dim oXFormHelper As ConstructXform
    
    Dim dblArray(2) As Double
    
    Set oPlane = m_oHoleTrace.GetWorkingPlane
    
    Set oPosition = New DPosition
    oPlane.GetRootPoint dblArray(0), dblArray(1), dblArray(2)
    oPosition.Set dblArray(0), dblArray(1), dblArray(2)
    
    Set oUVector = New DVector
    oPlane.GetUDirection dblArray(0), dblArray(1), dblArray(2)
    oUVector.Set dblArray(0), dblArray(1), dblArray(2)
    
    Set oVVector = New DVector
    oPlane.GetVDirection dblArray(0), dblArray(1), dblArray(2)
    oVVector.Set dblArray(0), dblArray(1), dblArray(2)
    
    Set oXFormHelper = New ConstructXform
    oXFormHelper.CoordinateTransform oPosition, oUVector, oVVector, oMatrix
    
    Set HoleMatrix = oMatrix
    
Cleanup:
    Set oUVector = Nothing
    Set oVVector = Nothing
    Set oPosition = Nothing
    Set oXFormHelper = Nothing
    Set oMatrix = Nothing
    Set oPlane = Nothing
    
    Exit Property
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Property

'******************************************************************************
' METHOD:  OrientationAngle (Get)
'
' DESCRIPTION:  if parent outfitting has cross-section - is it rotated
'******************************************************************************
Public Property Get OrientationAngle() As Double
    On Error GoTo ErrorHandler
    Const sMETHOD = "OrientationAngle"
    
    'set default in case something bad happens
    OrientationAngle = 0

'   Changes for TR 50621 and TR50622.
'   Getting the Vvector of the Hole Trace
'   Get the Plane
    Dim oPlane As IJPlane
    Dim oVVectorNormal As IJDVector
    Dim oVVectorHoleTrace As IJDVector
    Dim oVVectorHoleFitting As IJDVector
    
    Dim dblArrayNM(2) As Double   ' Array to hold Normal vector
    Dim dblArrayHT(2) As Double   ' Array to hold Hole Trace V Vector
    Dim dblArrayFT(5) As Double   ' Array to hold fitting V Vector
    
    Set oPlane = m_oHoleTrace.GetWorkingPlane
    Set oVVectorNormal = New DVector
    oPlane.GetNormal dblArrayNM(0), dblArrayNM(1), dblArrayNM(2)
    oVVectorNormal.Set dblArrayNM(0), dblArrayNM(1), dblArrayNM(2)
    
    Set oVVectorHoleTrace = New DVector

    oPlane.GetVDirection dblArrayHT(0), dblArrayHT(1), dblArrayHT(2)
    oVVectorHoleTrace.Set dblArrayHT(0), dblArrayHT(1), dblArrayHT(2) ' Got HoleTrace V Vector
    
    Dim oObject As Object
    Dim oOutfitting As IJDObjectCollection
    Dim oRteFeatureUtility As IJRtePathCrossSectUtility
    
    'this is just going to look at the first outfitting
    Set oOutfitting = m_oHoleTrace.GetParentOutfitting
    For Each oObject In oOutfitting
        If Not oObject Is Nothing Then Exit For
    Next oObject
    
    Dim ePlateType As StructPlateType
    ePlateType = PlateType
    
    'this only needs to be checked on duct and cableway
    If TypeOf oObject Is IJRteDuctPathFeat Or TypeOf oObject Is IJRteCablewayPathFeat Then
        Set oRteFeatureUtility = oObject
        ' need to add 90 degrees to take into account the switching of the
        '  width and depth going from V3 to V4
        
'       Get HoleFitting ( Cable Tray V Vector)
        oRteFeatureUtility.GetWidthAndDepthAxis Nothing, _
                                                dblArrayFT(0), _
                                                dblArrayFT(1), _
                                                dblArrayFT(2), _
                                                dblArrayFT(3), _
                                                dblArrayFT(4), _
                                                dblArrayFT(5)
        Set oVVectorHoleFitting = New DVector
'       V Vector is the 3rd,4th and 5th component.
        oVVectorHoleFitting.Set dblArrayFT(3), dblArrayFT(4), dblArrayFT(5)
        
'       Since the Alignment angle is being found by 1st principle( u&v dir)
'       it is same for all plates ( Deck, Long Bulkhead or Trans Bulkhead and any other slope)
        OrientationAngle = oVVectorHoleTrace.Angle(oVVectorHoleFitting, oVVectorNormal)
    End If

Cleanup:
    Set oObject = Nothing
    Set oOutfitting = Nothing
    Set oRteFeatureUtility = Nothing
    
    Set oPlane = Nothing
    Set oVVectorNormal = Nothing
    Set oVVectorHoleTrace = Nothing
    Set oVVectorHoleFitting = Nothing
    
    Exit Property
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Property


'******************************************************************************
'******************************************************************************
'
' Start of construction methods for the hole trace
'
'******************************************************************************
'******************************************************************************


'******************************************************************************
' METHOD:  PlaceEdgeReinforcement
'
' DESCRIPTION:  constructs the edge reinforcement (profile) fitting
'******************************************************************************
Public Function PlaceEdgeReinforcement(oXSect As IJCrossSection) As Object
    On Error GoTo ErrorHandler
    Const sMETHOD = "PlaceEdgeReinforcement"
    
    'which type of fitting is this
    Dim eFittingType As HMFittingType
    eFittingType = FittingType
    
    'get the resource manager
    Dim oResourceMgr As Object
    Set oResourceMgr = GetResourceMgrFromObject(m_oHoleTrace)
    If oResourceMgr Is Nothing Then GoTo Cleanup
    
    'get the parent structure - will be parent of fitting
    Dim oParentStructure As Object
    Set oParentStructure = m_oHoleTrace.GetParentStructure
    If oParentStructure Is Nothing Then GoTo Cleanup
        
    'need the plate type (deck, bulkhead, etc)
    Dim ePlateType As StructPlateType
    ePlateType = PlateType
    
    'get the plate part that owns the cutout
    Dim oPlatePart As IJPlate
    Set oPlatePart = CuttingPlatePart
    If oPlatePart Is Nothing Then GoTo Cleanup
    
    'create a plate part wrapper
    Dim oSDOPlatePart As StructDetailObjects.PlatePart
    Set oSDOPlatePart = New StructDetailObjects.PlatePart
    If oSDOPlatePart Is Nothing Then GoTo Cleanup
    Set oSDOPlatePart.Object = oPlatePart
    
    'get the hole feature that cut the structure
    Dim oCurve As IJStructGenericContour
    On Error Resume Next
    Set oCurve = m_oHoleTrace.GetHoleFeature
    On Error GoTo ErrorHandler
    
    Dim oPortOrHoleTrace As Object
    
    If oCurve Is Nothing Then
        Set oPortOrHoleTrace = m_oHoleTrace
    Else
        Set oPortOrHoleTrace = oSDOPlatePart.CutoutEdgePort(oCurve)
        If oPortOrHoleTrace Is Nothing Then GoTo Cleanup
    End If
        
    'create edge reinforcement using StructDetail wrapper
    Dim eProfileFaceName As ProfileFaceName
    Dim lLoadingPt As Long
    Dim ePrimOrientation As StructMoldedOrientation
    Dim eSecondOrientation As StructMoldedOrientation
    Dim eOrientationRule As ProfileOrientationRule
    Dim ePosition As ERPosition
    Dim dOffset As Double

    ePosition = ER_OnFace
    ePrimOrientation = AboveOrient
    eSecondOrientation = BelowOrient
    eProfileFaceName = ValidMountingFace(oResourceMgr, oXSect)
    lLoadingPt = ValidLoadPoint(oResourceMgr, oXSect, eProfileFaceName)
    eOrientationRule = ProfileOrientationNormal
    dOffset = 0#
    
    'set the values based on plate type and coaming type
    If eFittingType = HM_FittingCoaming Then
        ePosition = ER_OnFace
        eSecondOrientation = ER_TowardPlate
        eProfileFaceName = 4
        lLoadingPt = 3
       
        Select Case ePlateType
            Case DeckPlate
                ePrimOrientation = AboveOrient
            Case LBulkheadPlate
                ePrimOrientation = OutboardOrient
            Case TBulkheadPlate
                ePrimOrientation = AftOrient
            Case Else
                ePrimOrientation = AboveOrient
        End Select
        
    ElseIf eFittingType = HM_FittingPenetrated Then
        ePosition = ER_OnEdgeCentered
        eSecondOrientation = ER_FromPlate
        eProfileFaceName = 9
        lLoadingPt = 25
        
        Select Case ePlateType
            Case DeckPlate
                ePrimOrientation = AboveOrient
            Case LBulkheadPlate
                ePrimOrientation = OutboardOrient
            Case TBulkheadPlate
                ePrimOrientation = AftOrient
            Case Else
                ePrimOrientation = AboveOrient
        End Select
    
    End If
    
    'create the actual fitting
    Dim oER As StructDetailObjects.EdgeReinforcement
    Set oER = New StructDetailObjects.EdgeReinforcement
    oER.CreateEdgeReinforcementProfilePart oResourceMgr, _
                                           oPlatePart, _
                                           oPortOrHoleTrace, _
                                           oXSect, _
                                           eProfileFaceName, _
                                           lLoadingPt, _
                                           ePrimOrientation, _
                                           eSecondOrientation, _
                                           eOrientationRule, _
                                           ePosition, _
                                           dOffset
    If oER.Object Is Nothing Then GoTo Cleanup
    
    'set the fitting as a child of the hole trace
    Dim oDesignParent As IJDesignParent
    Dim oDesignChild As IJDesignChild
    On Error Resume Next
    Set oDesignParent = oPlatePart
    If oDesignParent Is Nothing Then
        Set oDesignParent = oParentStructure
    End If
    Set oDesignChild = oER.Object
    On Error GoTo ErrorHandler
    oDesignParent.AddChild oDesignChild
    
    'set the properties of the fitting
    SetERProfilePartMaterial oER.Object
'    SetERProfilePartProperties oER.Object, ePrimOrientation
    SetERProfilePartNamingCategory oResourceMgr, oER.Object
    
    'return the fitting
    Set PlaceEdgeReinforcement = oER.Object
    
Cleanup:
    Set oResourceMgr = Nothing
    Set oParentStructure = Nothing
    Set oCurve = Nothing
    Set oPlatePart = Nothing
    Set oSDOPlatePart = Nothing
    Set oPortOrHoleTrace = Nothing
    Set oER = Nothing
    Set oDesignParent = Nothing
    Set oDesignChild = Nothing
                                   
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Function

'******************************************************************************
' METHOD:  PlaceCatalogFitting
'
' DESCRIPTION:  Constructs the fitting from the catalog
'******************************************************************************
Public Function PlaceCatalogFitting() As IJPartOcc
    On Error GoTo ErrorHandler
    Const sMETHOD = "PlaceCatalogFitting"
    
    'get the resource manager
    Dim oResourceMgr As Object
    Set oResourceMgr = GetResourceMgrFromObject(m_oHoleTrace)
    If oResourceMgr Is Nothing Then GoTo Cleanup
    
    'need a factory
    Dim oHoleFactory As IJHoleMgmtFactory
    Set oHoleFactory = New CHoleMgmtFactory
    If oHoleFactory Is Nothing Then GoTo Cleanup
    
    'get the part from the fitting active entity
    Dim oFittingAE As IJHoleFitting
    On Error Resume Next
    Set oFittingAE = m_oHoleTrace.GetHoleFitting
    On Error GoTo ErrorHandler
    If oFittingAE Is Nothing Then GoTo Cleanup
    
    Dim oPart As IJDPart
    Set oPart = oFittingAE.FittingPart
    If oPart Is Nothing Then GoTo Cleanup
    
    'check for orientation angle of outfitting
    Dim dAngle As Double
    dAngle = OrientationAngle
    
    'create the actual occurrence
    Dim oPartOcc As IJPartOcc
    Set oPartOcc = oHoleFactory.CreateHoleFittingOcc(oResourceMgr, oPart)
    If oPartOcc Is Nothing Then GoTo Cleanup
    
    'set the matrix on the part
    Dim oMatrix As IJDT4x4
    Set oMatrix = HoleMatrix

    'rotate the matrix if needed
    If dAngle > 0 Then
        Dim oVector As DVector
        Set oVector = New DVector
        oVector.Set 0, 0, 1
        oMatrix.Rotate dAngle, oVector
        Set oVector = Nothing
    End If
    
    oPartOcc.SetMatrix oMatrix
        
    'set the part type on the fitting
    SetPartType oResourceMgr, oPartOcc
    
    'return the fitting
    Set PlaceCatalogFitting = oPartOcc
    
Cleanup:
    Set oResourceMgr = Nothing
    Set oHoleFactory = Nothing
    Set oFittingAE = Nothing
    Set oPart = Nothing
    Set oPartOcc = Nothing
    Set oMatrix = Nothing
    Set oVector = Nothing
                                   
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Function

'******************************************************************************
' METHOD:  PlacePlateFitting
'
' DESCRIPTION:  constructs the plate part fitting
'******************************************************************************
Public Function PlacePlateFitting() As Object
    On Error GoTo ErrorHandler
    Const sMETHOD = "PlacePlateFitting"
    
    'get the resource manager
    Dim oResourceMgr As Object
    Set oResourceMgr = GetResourceMgrFromObject(m_oHoleTrace)
    If oResourceMgr Is Nothing Then GoTo Cleanup
    
    'which type of fitting is this
    Dim eFittingType As HMFittingType
    eFittingType = FittingType
    
    'how many outfitting are there
    Dim lCount As Long
    Dim oOutfitting As IJDObjectCollection
    Set oOutfitting = m_oHoleTrace.GetParentOutfitting
    lCount = oOutfitting.Count
    
    'get the hole feature
    Dim oCurve As Object
    
    Set oCurve = m_oHoleTrace.GetHoleFeature
    On Error GoTo ErrorHandler
    If oCurve Is Nothing Then
        ' Fitting on un-cut plate,use hole trace curve
        Set oCurve = m_oHoleTrace
    End If

    'get the plate part that owns the cutout
    Dim oPlatePart As IJPlate
    Set oPlatePart = CuttingPlatePart
    If oPlatePart Is Nothing Then GoTo Cleanup
    
    'create a plate part wrapper
    Dim oSDOPlatePart As StructDetailObjects.PlatePart
    Set oSDOPlatePart = New StructDetailObjects.PlatePart
    If oSDOPlatePart Is Nothing Then GoTo Cleanup
    Set oSDOPlatePart.Object = oPlatePart
    
    'get the parent structure - will be parent of fitting
    Dim oParentStructure As Object
    Set oParentStructure = m_oHoleTrace.GetParentStructure
    If oParentStructure Is Nothing Then GoTo Cleanup
    
    'get the parent system (try cutting plate part first - then parent structure)
    Dim oSystem As IJSystem
    On Error Resume Next
    Set oSystem = oPlatePart
    If oSystem Is Nothing Then
        Set oSystem = oParentStructure
    End If
    On Error GoTo ErrorHandler
    If oSystem Is Nothing Then GoTo Cleanup
    
    ' get the thickness direction of the "parent" plate
    Dim eThicknessDir As StructMoldedDirection
    eThicknessDir = oSDOPlatePart.ThicknessDirection
    
    ' Valid thickness direction for plate edge reinforcement standalone part is
    ' Above/Below/Fore/Aft/Port/Starboard only.  Convert Inboard/Outboard/Centered
    ' to a valid value
    If eThicknessDir = InboardDir Then
        eThicknessDir = AboveDir  ' Correponds to Port for LBulkhead or Fore for TBulkhead
    ElseIf eThicknessDir = OutboardDir Then
        eThicknessDir = BelowDir  ' Correponds to Starboard for LBulkhead or Aft for TBulkhead
    ElseIf eThicknessDir = Centered Then
        eThicknessDir = AboveDir
    End If
    
    'if the fitting is a center flange and there is only one outfitting, the
    'resulting fitting is the same as a double ring so just use the same code
    'if the fitting is a center flange and there are more than one outfitting,
    'then need to place a special center flange
    
    Dim oER As StructDetailObjects.EdgeReinforcement
    Set oER = New StructDetailObjects.EdgeReinforcement
    If oER Is Nothing Then GoTo Cleanup
    
    'create the fitting based on type and outfitting count
    Dim oPlateFitting As Object
    If (lCount >= 1 And eFittingType = HM_FittingCenterFlange) Then
        'need to place a center flange with inner cuts
        '!!!! hardcode offset value to 0.065
        oER.CreateEdgeReinforcementPlanarPlatePart oResourceMgr, _
                                                   oCurve, _
                                                   eThicknessDir, _
                                                   0.065, _
                                                   -999999.99, _
                                                   oSystem, _
                                                   oPlatePart
        'get the fitting object
        Set oPlateFitting = oER.Object
        If oPlateFitting Is Nothing Then GoTo Cleanup
    
        'need to cut the center flange
        CutCenterFlange oPlateFitting
        
    Else
        'can just create the EdgeReinforcementPlanarPlatePart
        '!!!! hardcode offset values to 0.065 and .015
        oER.CreateEdgeReinforcementPlanarPlatePart oResourceMgr, _
                                                   oCurve, _
                                                   eThicknessDir, _
                                                   0.065, _
                                                   0.015, _
                                                   oSystem, _
                                                   oPlatePart
        'get the fitting object
        Set oPlateFitting = oER.Object
        If oPlateFitting Is Nothing Then GoTo Cleanup
    End If
    
    'set the attribute PartType value on IJExternalWeldedPipePart
    SetPartType oResourceMgr, oPlateFitting
    
    'return the fitting
    Set PlacePlateFitting = oPlateFitting
    
Cleanup:
    Set oResourceMgr = Nothing
    Set oOutfitting = Nothing
    Set oCurve = Nothing
    Set oPlatePart = Nothing
    Set oSDOPlatePart = Nothing
    Set oParentStructure = Nothing
    Set oSystem = Nothing
    Set oER = Nothing
    Set oPlateFitting = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Function

'******************************************************************************
' METHOD:  PlaceProfileFitting
'
' DESCRIPTION:  Constructs the profile part fitting
'                   Gets the curve from Sketch2D
'                   Gets the plane from Sketch2D
'                   Creates ProfileByProjection
'******************************************************************************
Public Function PlaceProfileFitting(oXSect As IJCrossSection) As Object
    On Error GoTo ErrorHandler
    Const sMETHOD = "PlaceProfileFitting"
    
    'get the resource manager
    Dim oResourceMgr As Object
    Set oResourceMgr = GetResourceMgrFromObject(m_oHoleTrace)
    If oResourceMgr Is Nothing Then GoTo Cleanup
    
    'need the fitting active entity
    Dim oFittingAE As IJHoleFitting
    On Error Resume Next
    Set oFittingAE = m_oHoleTrace.GetHoleFitting
    On Error GoTo ErrorHandler
    If oFittingAE Is Nothing Then GoTo Cleanup
    
    'get the sketching plane and curve from the fitting ae
    Dim oPlane As IJPlane
    Set oPlane = oFittingAE.FittingSketchPlane
    If oPlane Is Nothing Then GoTo Cleanup
    
    Dim oContour As IJComplexString
    Set oContour = oFittingAE.FittingSketchCurve
    If oContour Is Nothing Then GoTo Cleanup
    
    'get the plate system
    Dim oStructure As Object
    Set oStructure = m_oHoleTrace.GetParentStructure
    If oStructure Is Nothing Then GoTo Cleanup
    
    If Not TypeOf oStructure Is IJPlate Then GoTo Cleanup
    
    Dim oPlate As IJPlate
    Set oPlate = oStructure
    If oPlate Is Nothing Then GoTo Cleanup
    
    'create the wrapper to use
    Dim oSDOProfilePart As StructDetailObjects.ProfilePart
    Set oSDOProfilePart = New StructDetailObjects.ProfilePart
    oSDOProfilePart.CreateByProjection oResourceMgr, _
                                       oPlate, _
                                       oPlane, _
                                       oContour, _
                                       oXSect, _
                                       oPlate, _
                                       BottomSideOfInferiorFlange, _
                                       1, _
                                       primOrientationVertical, _
                                       ForeOrient
    
    'get the actual fitting
    Dim oProfileFitting As Object
    Set oProfileFitting = oSDOProfilePart.Object
    If oProfileFitting Is Nothing Then GoTo Cleanup
    
    'set the material and grade of the new profile part
    SetProfilePartMaterialAndGrade oProfileFitting
    
    'return the fitting
    Set PlaceProfileFitting = oProfileFitting
    
Cleanup:
    Set oResourceMgr = Nothing
    Set oFittingAE = Nothing
    Set oPlane = Nothing
    Set oContour = Nothing
    Set oStructure = Nothing
    Set oPlate = Nothing
    Set oSDOProfilePart = Nothing
    Set oProfileFitting = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Function

'***********************************************************************
' METHOD:  CutCenterFlange
'
' DESCRIPTION:  cut the center flange for pipes
'***********************************************************************
Public Sub CutCenterFlange(oPlateFitting As Object)
    On Error GoTo ErrorHandler
    Const sMETHOD = "CreateCenterFlangeInnerCuts"
    
    'get the resource manager
    Dim oResourceMgr As Object
    Set oResourceMgr = GetResourceMgrFromObject(m_oHoleTrace)
    If oResourceMgr Is Nothing Then GoTo Cleanup
        
    'get the parent structure
    Dim oPenetratedPlate As IJPlate
    Set oPenetratedPlate = CuttingPlatePart
    If oPenetratedPlate Is Nothing Then GoTo Cleanup
    
    'create the inner cuts
    Dim oInnerCuts As IJDObjectCollection
    Set oInnerCuts = CreateCenterFlangeInnerCuts(oResourceMgr)
    If oInnerCuts Is Nothing Then GoTo Cleanup
    
    'going to need a middle helper to do the actual cutting
    Dim oHMHelper As IJHMMiddleHelper
    Set oHMHelper = New CHMMiddleHelper
    
    oHMHelper.CutCenterFlange oResourceMgr, _
                              oPlateFitting, _
                              oInnerCuts, _
                              m_oHoleTrace, _
                              oPenetratedPlate

Cleanup:
    Set oResourceMgr = Nothing
    Set oPenetratedPlate = Nothing
    Set oInnerCuts = Nothing
    Set oHMHelper = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Sub

'***********************************************************************
' METHOD:  CreateCenterFlangeInnerCuts
'
' DESCRIPTION:  create the inner cuts for the center flange
'***********************************************************************
Public Function CreateCenterFlangeInnerCuts(oResourceMgr As Object) As IJDObjectCollection
    On Error GoTo ErrorHandler
    Const sMETHOD = "CreateCenterFlangeInnerCuts"
    
    'create the collection
    Dim oInnerCuts As IJDObjectCollection
    Set oInnerCuts = New JObjectCollection
    
    'need a port on the fitting
    Dim oStructure As Object
    Set oStructure = Nothing 'there is none at this time - ComputeCenterFlangeTrace will accept "Nothing"
    
    'need the working plane
    Dim oWorkingPlane As IJPlane
    Set oWorkingPlane = m_oHoleTrace.GetWorkingPlane

    'need the list of outfitting
    Dim oOutfitting As IJDObjectCollection
    Set oOutfitting = m_oHoleTrace.GetParentOutfitting
    
    'need the plate thickness
    Dim dPlateThickness As Double
    dPlateThickness = PlateThickness
    
    'going to use the best-fit symbol helper to do this
    Dim bSelectedRun As Boolean
    Dim oDefinitionRule As Object
    Dim oBestFitHelper As IJHMBestFitHelper
    Dim oRepresentOutfitting As IJDObjectCollection
    Dim oSingleOutfitting As IJDObjectCollection
          
''' In V8, the custom VB symbols are not registered
'''    Set oDefinitionRule = New HMBestFit.DefinitionRule
    Set oDefinitionRule = SP3DCreateObject(sDefRule)
    
    Set oBestFitHelper = New CHMBestFitHelper
    Set oRepresentOutfitting = New JObjectCollection
    Set oSingleOutfitting = New JObjectCollection
    
    Dim oGeometry As IJComplexString
    
    Dim strName As String
    
    Dim oObject As Object
    For Each oObject In oOutfitting
        oSingleOutfitting.Clear
        oSingleOutfitting.Add oObject
        oRepresentOutfitting.Clear
        bSelectedRun = oDefinitionRule.ComputeCenterFlangeTrace(oStructure, oSingleOutfitting, oRepresentOutfitting)
        If bSelectedRun Then
            Set oGeometry = oBestFitHelper.TraceGeneration(oResourceMgr, _
                                                           oWorkingPlane, _
                                                           dPlateThickness, _
                                                           oRepresentOutfitting)
            oInnerCuts.Add oGeometry
            Set oGeometry = Nothing
        End If
        'Exit For 'just for now - lets try and get one first
    Next oObject
    
    Set CreateCenterFlangeInnerCuts = oInnerCuts
    
Cleanup:
    Set oInnerCuts = Nothing
    Set oWorkingPlane = Nothing
    Set oOutfitting = Nothing
    Set oDefinitionRule = Nothing
    Set oBestFitHelper = Nothing
    Set oRepresentOutfitting = Nothing
    Set oSingleOutfitting = Nothing
    Set oGeometry = Nothing
    Set oObject = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Function


'***********************************************************************
'***********************************************************************
'
' Start of private methods used by this object
'
'***********************************************************************
'***********************************************************************


'***********************************************************************
' METHOD:  ValidMountingFace
'
' DESCRIPTION:  return a mounting face for the cross section
'***********************************************************************
Private Property Get ValidMountingFace(oResourceMgr As Object, _
                                       oXSect As IJCrossSection) As ProfileFaceName
    On Error GoTo ErrorHandler
    Const sMETHOD = "ValidMountingFace"
    
    Dim oMiddleHelper As IJHMMiddleHelper
    Set oMiddleHelper = New CHMMiddleHelper
    
    Dim oSectionType As IJDCrossSectionType
    Set oSectionType = oMiddleHelper.GetSectionTypeFromSection(oResourceMgr, oXSect)
    
    Dim mountingFacePictures As IJDObjectCollection
    Set mountingFacePictures = oSectionType.GetMountingFacePictures
    ValidMountingFace = oSectionType.GetMountingFace(0)
    
Cleanup:
    Set oMiddleHelper = Nothing
    Set oSectionType = Nothing
    
    Exit Property
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Property

'***********************************************************************
' METHOD:  ValidLoadPoint
'
' DESCRIPTION:  return a load point for the cross section
'***********************************************************************
Private Property Get ValidLoadPoint(oResourceMgr As Object, oXSect As IJCrossSection, _
                                    eFaceName As ProfileFaceName) As Long
    On Error GoTo ErrorHandler
    Const sMETHOD = "ValidLoadPoint"
    
    Dim oMiddleHelper As IJHMMiddleHelper
    Set oMiddleHelper = New CHMMiddleHelper
    
    Dim oSectionType As IJDCrossSectionType
    Set oSectionType = oMiddleHelper.GetSectionTypeFromSection(oResourceMgr, oXSect)
    
    Dim loadingPointPictures As IJDObjectCollection
    Set loadingPointPictures = oSectionType.GetLoadPointPictures(eFaceName)
    ValidLoadPoint = oSectionType.GetLoadPoint(oSectionType.GetIndexForMountingFace(eFaceName), 0)
    
Cleanup:
    Set oMiddleHelper = Nothing
    Set oSectionType = Nothing
    
    Exit Property
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Property

'***********************************************************************
' METHOD:  SetERProfilePartMaterial
'
' DESCRIPTION:  set the material for the profile part
'***********************************************************************
Private Sub SetERProfilePartMaterial(oERObject As Object)
    On Error GoTo ErrorHandler
    Const sMETHOD = "SetERProfilePartMaterial"
    
    Dim arrMaterial() As String
    Dim arrGrade() As String
    
    Dim oSpecMgr As New GSCADMoldedFormSpecMgr.JMoldedFormSpecMgr
    Dim oStructQuery2 As IJDStructQuery2
    Set oStructQuery2 = oSpecMgr.GetSpecs.Item(1)
    Set oSpecMgr = Nothing

    Dim oActiveMaterialObject As Object
    Dim arrMaterialNames() As String, arrMaterialGrades() As String
    
    oStructQuery2.GetMaterials MaterialForProfile, arrMaterialNames
    oStructQuery2.GetGradesForMaterial MaterialForProfile, arrMaterialNames(0), arrMaterialGrades
    oStructQuery2.GetMaterial arrMaterialNames(0), arrMaterialGrades(0), oActiveMaterialObject
    
    If Not oActiveMaterialObject Is Nothing Then
        Dim oStructureMaterial As IJStructureMaterial
        Set oStructureMaterial = oERObject
        oStructureMaterial.Material = oActiveMaterialObject
        
        Set oStructureMaterial = Nothing
        Set oActiveMaterialObject = Nothing
    End If
    
Cleanup:
    Set oStructQuery2 = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Sub

'***********************************************************************
' METHOD:  SetERProfilePartNamingCategory
'
' DESCRIPTION:  set the naming category for the profile part
'***********************************************************************
Private Sub SetERProfilePartNamingCategory(oResourceMgr As Object, oERObject As Object)
    On Error GoTo ErrorHandler
    Const sMETHOD = "SetERProfilePartNamingCategory"

    'default naming category to first non-negative value
    Dim oQueryUtil As IJMetaDataCategoryQuery
    Dim oIJProfile As IJProfile
    
    Dim strLongNames() As String
    Dim strShortNames() As String
    Dim lPriority() As Long
    Dim idx As Long

    Set oQueryUtil = New CMetaDataCategoryQuery
    oQueryUtil.GetEdgeReinforcementCategoryInfo oResourceMgr, _
                                                strLongNames, _
                                                strShortNames, _
                                                lPriority

    Set oIJProfile = oERObject
    oIJProfile.NamingCategory = -1 'init to indicate value not set
    For idx = LBound(lPriority) To UBound(lPriority)
        If lPriority(idx) >= 0 Then
            oIJProfile.NamingCategory = lPriority(idx)
            Exit For
        End If
    Next idx

Cleanup:
    Set oQueryUtil = Nothing
    Set oIJProfile = Nothing
    
    Erase strLongNames
    Erase strShortNames
    Erase lPriority
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Sub

'***********************************************************************
' METHOD:  SetERProfilePartProperties
'
' DESCRIPTION:  set the properties for the profile part
'***********************************************************************
Private Sub SetERProfilePartProperties(oERObject As Object, _
                                       ePrimaryOrient As StructMoldedOrientation)
    On Error GoTo ErrorHandler
    Const sMETHOD = "SetERProfilePartMaterial"
    
    Dim oMoldedConventions As IJDProfileMoldedConventions
    Set oMoldedConventions = oERObject
    oMoldedConventions.Orientation = ePrimaryOrient
    
Cleanup:
    Set oMoldedConventions = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Sub

'***********************************************************************
' METHOD:  SetProfilePartMaterialAndGrade
'
' DESCRIPTION:  Set the material and grade of the specified profile part
'               to be that of the parent plate part.
'***********************************************************************
Private Sub SetProfilePartMaterialAndGrade(oProfileObject As Object)
    On Error GoTo ErrorHandler
    Const sMETHOD = "SetProfilePartMaterialAndGrade"

    ' Get the parent plate part of the profile.
    Dim oParentPlatePart As IJPlate
    Set oParentPlatePart = CuttingPlatePart

    ' Get the material of the parent plate part.
    Dim oPlateMaterial As IJStructureMaterial
    If TypeOf oParentPlatePart Is IJStructureMaterial Then
        Set oPlateMaterial = oParentPlatePart
    Else
        GoTo Cleanup
    End If
    
    ' Get the interface to the material of the profile part.
    Dim oProfileMaterial As IJStructureMaterial
    If TypeOf oProfileObject Is IJStructureMaterial Then
        Set oProfileMaterial = oProfileObject
    Else
        GoTo Cleanup
    End If
    
    ' Set the material of the profile to that of the plate.
    oProfileMaterial.Material = oPlateMaterial.Material

Cleanup:
    Set oParentPlatePart = Nothing
    Set oPlateMaterial = Nothing
    Set oProfileMaterial = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Sub

'***********************************************************************
' METHOD:  SetPartType
'
' DESCRIPTION:  set the part type to allow the fitting to participate
'               in a pipe spool
'***********************************************************************
Private Sub SetPartType(ByVal oResourceMgr As Object, ByVal oFittingObj As Object)
    On Error GoTo ErrorHandler
    Const sMETHOD = "SetPartType"
       
    Dim oJDAttributes As IMSAttributes.IJDAttributes
    Dim oJDAttributesCol As IMSAttributes.IJDAttributesCol
    Dim oJDAttr As IMSAttributes.IJDAttribute
            
    On Error Resume Next
    Set oJDAttributes = oFittingObj
    If Not oJDAttributes Is Nothing Then
        Set oJDAttributesCol = oJDAttributes.CollectionOfAttributes(strIJExtWeldPPInterfaceID)
        If Not oJDAttributesCol Is Nothing Then
            If oJDAttributesCol.Count > 0 Then
                Set oJDAttr = oJDAttributesCol.Item("PartType")
                If Not oJDAttr Is Nothing Then
                    oJDAttr.Value = 5 'Penetration Plate
                End If
            End If
        End If
    End If
    
Cleanup:
    Set oJDAttr = Nothing
    Set oJDAttributesCol = Nothing
    Set oJDAttributes = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, sMETHOD).Number
    GoTo Cleanup
End Sub
 
