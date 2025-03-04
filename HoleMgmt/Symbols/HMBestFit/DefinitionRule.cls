VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefinitionRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'Copyright (C) 1998, Intergraph Corporation. All rights reserved.
'
'Project: Hole Management - HoleMgmt\Middle\Symbols\HMBestFit
'
'File: DefinitionRule.bas
'
'Abstract: Definition Rule to Compute HoleTrace of BestFit symbol
'
'Author: sypark@ship.samsung.co.kr
'
'Revision:
'     05/22/01 - sypark@ship.samsung.co.kr - initial release
'
'Note: ReportAndRaiseUnanticipatedError method is located in
'      CommonApp/Client/Bin/GSCADUtilities.dll and may have to
'      be replaced
'
' Changes:
'   Date    By    Description
' 06/09/03  CRS  Consider pipe turn features when computing projections
'                to pass to the best fit helper.
' 09/09/03  CRS  Added function ComputePipeWithSleeveTrace to handle the
'                special case of pipe sleeve catalog fittings.
'
'*******************************************************************
Option Explicit

Private Const MODULE = "HMBestFit.DefinitionRule(DefinitionRule.bas)"
Private m_oErrors As New IMSErrorLog.JServerErrors

'For Pipe Simplification
Dim m_PipeHoleMinDist As Double 'If the nearest distance between two holes is less than 50 mm,
                                ' then two hole size is same

'For Pipe Clearance
Dim m_PipeFlangeCr As Double
Dim m_PipeSquareFlangeCr As Double
Dim m_PipeSleeveTightCr As Double
Dim m_PipeSleeveNoTightInsulationCr As Double
Dim m_PipeSleeveNoTightNoInsulationCr As Double
Dim m_PipeNoneTightCr As Double
Dim m_PipeNoneInsulationCr As Double
Dim m_PipeNoneNoInsulationCr As Double

'For Pipe Insulation thickness
Dim m_PipeInsulationThickness As Double

'For Pipe Round Value
Dim m_PipeFlangeRv As Boolean
Dim m_PipeSquareFlangeRv As Boolean
Dim m_PipeSleeveTightRv As Boolean
Dim m_PipeSleeveNoTightInsulationRv As Boolean
Dim m_PipeSleeveNoTightNoInsulationRv As Boolean
Dim m_PipeNoneTightRv As Boolean
Dim m_PipeNoneInsulationRv As Boolean
Dim m_PipeNoneNoInsulationRv As Boolean

'For Duct Clearance
Dim m_DuctFlangeCr As Double
Dim m_DuctFlangeInsulationCr As Double
Dim m_DuctFlangeNoInsulationCr As Double
Dim m_DuctSleeveCr As Double
Dim m_DuctSleeveInsulationCr As Double
Dim m_DuctSleeveNoInsulationCr As Double
Dim m_DuctNoneCr As Double
Dim m_DuctNoneInsulationCr As Double
Dim m_DuctNoneNoInsulationCr As Double

'For Duct Insulation thickness
Dim m_DuctInsulationThickness As Double

'For Duct corner radius
Dim m_DuctCorRadius As Double

'For Duct Round Value
Dim m_DuctFlangeRv As Boolean
Dim m_DuctFlangeInsulationRv As Boolean
Dim m_DuctFlangeNoInsulationRv As Boolean
Dim m_DuctSleeveRv As Boolean
Dim m_DuctSleeveInsulationRv As Boolean
Dim m_DuctSleeveNoInsulationRv As Boolean
Dim m_DuctNoneRv As Boolean
Dim m_DuctNoneInsulationRv As Boolean
Dim m_DuctNoneNoInsulationRv As Boolean

'For Cableway Clearance
Dim m_CablewayFlangeCr As Double
Dim m_CablewayFlangeInsulationCr As Double
Dim m_CablewayFlangeNoInsulationCr As Double
Dim m_CablewaySleeveCr As Double
Dim m_CablewaySleeveInsulationCr As Double
Dim m_CablewaySleeveNoInsulationCr As Double
Dim m_CablewayNoneCr As Double
Dim m_CablewayNoneInsulationCr As Double
Dim m_CablewayNoneNoInsulationCr As Double

'For Cableway Insulation thickness
Dim m_CablewayInsulationThickness As Double

'For Cableway corner radius
Dim m_CablewayCorRadius As Double

'For Cableway Round Value
Dim m_CablewayFlangeRv As Boolean
Dim m_CablewayFlangeInsulationRv As Boolean
Dim m_CablewayFlangeNoInsulationRv As Boolean
Dim m_CablewaySleeveRv As Boolean
Dim m_CablewaySleeveInsulationRv As Boolean
Dim m_CablewaySleeveNoInsulationRv As Boolean
Dim m_CablewayNoneRv As Boolean
Dim m_CablewayNoneInsulationRv As Boolean
Dim m_CablewayNoneNoInsulationRv As Boolean

'For Conduit Simplification
Dim m_ConduitHoleMinDist As Double   'If the nearest distance between two holes is less than 50 mm,
                                     ' then two hole size is same

'For Conduit Clearance (note that conduits do not have insulation)
Dim m_ConduitFlangeCr As Double
Dim m_ConduitSquareFlangeCr As Double
Dim m_ConduitSleeveTightCr As Double
Dim m_ConduitSleeveNoTightCr As Double
Dim m_ConduitNoneTightCr As Double
Dim m_ConduitNoneCr As Double

'For Conduit Round Value
Dim m_ConduitFlangeRv As Boolean
Dim m_ConduitSquareFlangeRv As Boolean
Dim m_ConduitSleeveTightRv As Boolean
Dim m_ConduitSleeveNoTightRv As Boolean
Dim m_ConduitNoneTightRv As Boolean
Dim m_ConduitNoneRv As Boolean

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InitializeClearances
'
'Interface: Private function
'
'Abstract: This method is to get the clearance and Simplification rule
'
'Attention : The unit of measure should be mm in Registry.
'
'Application joint type    Tightness  Insulation Round value                          Hole size
'Pipe        Flange            -           -     5(m_PipeFlangeRv)                    Flange O.D + 10mm ( m_PipeFlangeCr )
'            Square flange     -           -     5(m_PipeSquareFlangeRv)              Flange diagonal + 10 mm ( m_PipeSquareFlangeCr )
'            Sleeve           Yes          -     1(m_PipeSleeveTightRv)               Sleeve O.D + 2mm ( m_PipeSleeveTightCr)
'                             No          Yes    5(m_PipeSleeveNoTightInsulationRv)   Pipe O.D + I.T * 2 + 25 mm ( m_PipeSleeveNoTightInsulationCr )
'                                         No     5(m_PipeSleeveNoTightNoInsulationCRv)Pipe O.D + 25 mm ( m_PipeSleeveNoTightNoInsulationCr )
'            None or Others   Yes          -     1(m_PipeNoneTightRv)                 Pipe O.D + 2 mm ( m_PipeNoneTightCr )
'                             No          Yes    5(m_PipeNoneInsulationRv)            Pipe O.D + I.T * 2 + 25 mm ( m_PipeNoneInsulationCr )
'                                         No     5(m_PipeNoneNoInsulationRv)          Pipe O.D + 25 mm (  m_PipeNoneNoInsulationCr )
'
'HVAC        Flange            -           -     5(m_DuctFlangeRv)                    Flange O-dim + 10 mm ( m_DuctFlangeCr )
'                             No          Yes    5(m_DuctFlangeInsulationRv)          Flange O-dim + 25 mm ( m_DuctFlangeInsulationCr )
'                                                                                     or Duct O-dim + I.T * 2 + 25 mm ( m_DuctFlangeInsulationCr ) (large one)
'                                         No     5(m_DuctFlangeNoInsulationRv)        Flange O-dim + 10 mm  ( m_DuctFlangeNoInsulationCr )
'            Sleeve           Yes          -     1(m_DuctSleeveRv)                    Sleeve O-dim + 2 mm ( m_DuctSleeveCr )
'                             No          Yes    1(m_DuctSleeveInsulationRv)          Sleeve O-dim + I.T * 2  + 2 mm ( m_DuctSleeveInsulationCr )
'                                         No     5(m_DuctSleeveNoInsulationRv)        Sleeve O-dim + 25 mm ( m_DuctSleeveNoInsulationCr )
'            None or Others   Yes          -     1(m_DuctNoneRv)                      Duct O-dim + 2 mm (  m_DuctNoneCr)
'                             No          Yes    5(m_DuctNoneInsulationRv)            Duct O-dim + I.T * 2 + 25 mm (  m_DuctNoneInsulationCr )
'                                         No     5(m_DuctNoneNoInsulationRv)          Duct O-dim + 25 mm (  m_DuctNoneNoInsulationCr )
'
'
'********************************************************************
Private Sub InitializeClearances(oHoleTrace As IJHoleTraceAE)

    Const Method = "InitializeClearances"
    On Error GoTo ErrorHandler
    
    'FIRST SET THE DEFAULT VALUES FROM ABOVE TABLE
    
    'For the nearest distance between two holes.
    m_PipeHoleMinDist = 0.05             'If the nearest distance between two holes is less than 50 mm,
                                         ' then two hole size is same
   
    'For Pipe Clearance
    m_PipeFlangeCr = 0.01
    m_PipeSquareFlangeCr = 0.01
    m_PipeSleeveTightCr = 0.002
    m_PipeSleeveNoTightInsulationCr = 0.025
    m_PipeSleeveNoTightNoInsulationCr = 0.025
    m_PipeNoneTightCr = 0.002
    m_PipeNoneInsulationCr = 0.025
    m_PipeNoneNoInsulationCr = 0.025

    'For Pipe Round Value
    m_PipeFlangeRv = True
    m_PipeSquareFlangeRv = True
    m_PipeSleeveTightRv = False
    m_PipeSleeveNoTightInsulationRv = True
    m_PipeSleeveNoTightNoInsulationRv = True
    m_PipeNoneTightRv = False
    m_PipeNoneInsulationRv = True
    m_PipeNoneNoInsulationRv = True

    'For Duct Clearance
    m_DuctFlangeCr = 0.001
    m_DuctFlangeInsulationCr = 0.025
    m_DuctFlangeNoInsulationCr = 0.01
    m_DuctSleeveCr = 0.002
    m_DuctSleeveInsulationCr = 0.002
    m_DuctSleeveNoInsulationCr = 0.025
    m_DuctNoneCr = 0.002
    m_DuctNoneInsulationCr = 0.025
    m_DuctNoneNoInsulationCr = 0.025

   'For Duct corner radius
    m_DuctCorRadius = 0.025

    'For Duct Round Value
    m_DuctFlangeRv = True
    m_DuctFlangeInsulationRv = True
    m_DuctFlangeNoInsulationRv = True
    m_DuctSleeveRv = False
    m_DuctSleeveInsulationRv = False
    m_DuctSleeveNoInsulationRv = True
    m_DuctNoneRv = False
    m_DuctNoneInsulationRv = True
    m_DuctNoneNoInsulationRv = True
    
    'For Cableway Clearance
    m_CablewayFlangeCr = 0.001
    m_CablewayFlangeInsulationCr = 0.025
    m_CablewayFlangeNoInsulationCr = 0.01
    m_CablewaySleeveCr = 0.002
    m_CablewaySleeveInsulationCr = 0.002
    m_CablewaySleeveNoInsulationCr = 0.025
    m_CablewayNoneCr = 0.002
    m_CablewayNoneInsulationCr = 0.025
    m_CablewayNoneNoInsulationCr = 0.025

   'For Cableway corner radius
    m_CablewayCorRadius = 0.025

    'For Cableway Round Value
    m_CablewayFlangeRv = True
    m_CablewayFlangeInsulationRv = True
    m_CablewayFlangeNoInsulationRv = True
    m_CablewaySleeveRv = False
    m_CablewaySleeveInsulationRv = False
    m_CablewaySleeveNoInsulationRv = True
    m_CablewayNoneRv = False
    m_CablewayNoneInsulationRv = True
    m_CablewayNoneNoInsulationRv = True
    
    'For the nearest distance between two holes
    m_ConduitHoleMinDist = 0.05          'If the nearest distance between two holes is less than 50 mm,
                                         ' then two hole size is same
   
    'For Conduit Clearance (note that conduits do not have insulation)
    m_ConduitFlangeCr = 0.01
    m_ConduitSquareFlangeCr = 0.01
    m_ConduitSleeveTightCr = 0.002
    m_ConduitSleeveNoTightCr = 0.025
    m_ConduitNoneTightCr = 0.002
    m_ConduitNoneCr = 0.025

    'For Conduit Round Value
    m_ConduitFlangeRv = True
    m_ConduitSquareFlangeRv = True
    m_ConduitSleeveTightRv = False
    m_ConduitSleeveNoTightRv = True
    m_ConduitNoneTightRv = False
    m_ConduitNoneRv = True

    'NOW TRY TO GET THE CLEARANCE RULE FROM THE DATABASE
    GetDatabaseClearances oHoleTrace
    
Cleanup:
    Exit Sub
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Sub

'///////////////////////////////////////////////////////////////////////
'***********************************************************************
'Method: ComputeTrace
'
'Interface: public function
'
'Abstract: Produce a collection of IJProjection objects (cylinders) that
'          represent the outfitting parts
'
'Arguments:
'  oStructure: Parent structure of the HoleTraceAE.  This is required to
'    determine the Plate Tightness.
'  oOutfitting: Collection of Path Features and/or Equipment for which a
'    hole trace is to be computed.
'  oOutputObjects: Resulting collection of projections (cylinder or
'    point) that represent the outfitting for which the hole trace is to
'    be determined.
'
'Return: Boolean
'  True: all objects in the outfitting collection are path features.
'  False: at least one of the objects in the outfitting collection is
'    an equipment.
'***********************************************************************
Public Function ComputeTrace _
       (oHoleTrace As IJHoleTraceAE, oStructure As Object, _
        oOutfitting As IMSCoreCollections.IJDObjectCollection, _
        oWorkingPlane As IJPlane, _
        ByRef oOutputObjects As IMSCoreCollections.IJDObjectCollection) As Boolean
    Const Method = "ComputeTrace"
    On Error GoTo ErrorHandler
    
    Dim oObj As Object
    
    oOutputObjects.Clear
    ComputeTrace = True
    
      
    'Define the Clearance and Minimun distance between Holes.
    InitializeClearances oHoleTrace
        
    'Get the plate tightness
    Dim bPlateTightness As Boolean
    bPlateTightness = PlateTightness(oStructure)
        
    'For Loop input object ( 1 .. icount .. n )
    For Each oObj In oOutfitting
        If TypeOf oObj Is IJRtePipePathFeat Then
            ComputePipeTrace oObj, oOutfitting, oWorkingPlane, bPlateTightness, oOutputObjects
            
        ElseIf TypeOf oObj Is IJRteDuctPathFeat Then
            ComputeDuctTrace oObj, oOutfitting, bPlateTightness, oOutputObjects
        
        ElseIf TypeOf oObj Is IJRteCablewayPathFeat Then
            ComputeCablewayTrace oObj, oOutfitting, bPlateTightness, oOutputObjects
        
        ElseIf TypeOf oObj Is IJRteConduitPathFeat Then
            ComputeConduitTrace oObj, oOutfitting, bPlateTightness, oOutputObjects
        
        ElseIf TypeOf oObj Is IJEquipment Or TypeOf oObj Is IJPort Then
            'If object is equipment, convert equipment to port, and returns port object
            ComputeTrace = False

            Dim oPort As IJPort
            Dim oConnectable As IJConnectable
            Dim oEquipment As IJEquipment
            
            Set oPort = oObj
            Set oConnectable = oPort.Connectable
            Set oEquipment = oConnectable
                    
            'If the Pipe & Duct has already selected, then oOutputObjects should be cleared.
            'Equip can't be merged into a hole with another piece of outfitting
            oOutputObjects.Clear
            oOutputObjects.Add oEquipment
            
            Set oPort = Nothing
            Set oConnectable = Nothing
            Set oEquipment = Nothing

            Exit For
        Else
            MsgBox "UNKNOWN OUTFITTING TYPE IN COMPUTETRACE"
        End If
    Next oObj
        
Cleanup:
    Set oObj = Nothing
    
    Exit Function
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function  ' ComputeTrace

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: ComputePipeTrace
'
'Interface: public function
'
'Abstract: Produce a collection of IJProjection objects (cylinders) that
'          represent an individual pipe part
'********************************************************************
Public Sub ComputePipeTrace(oPipe As Object, _
        oOutfitting As IMSCoreCollections.IJDObjectCollection, _
        oWorkingPlane As IJPlane, _
        bTightness As Boolean, _
        ByRef oOutputObjects As IMSCoreCollections.IJDObjectCollection)
    Const Method = "ComputePipeTrace"
    On Error GoTo ErrorHandler
    
    Dim oRtePathFeat As IJRtePathFeat
    Dim oProjection As IJProjection
    
    Dim eJointType As JointType
    
    Dim dFlangeDia As Double
    Dim dPipeDia As Double
    Dim dHoleSize As Double
    Dim dInsulationThickness As Double
    
    Set oRtePathFeat = oPipe
    
    'To get the flange size on the pipe feature,
    ' If can't find the flange on the pipe feature, and then ask the
    ' back and forward pipe on the pipe run if they have a flange.
    dFlangeDia = GetPipeFlangeDimensions(oRtePathFeat, eJointType)
    
    'If there is no flange on the Pipe Run, The Pipe size should be returned.
    If dFlangeDia <= 0 Then
        dPipeDia = GetPipeDimensions(oRtePathFeat)
    End If
    
    'To check the simplification among less than three pipes, Change the
    ' radius due to Simplification rule
    If oOutfitting.Count > 1 And oOutfitting.Count <= 3 Then
        If dFlangeDia > 0 Then
            dFlangeDia = PipeSimplification(dFlangeDia, m_PipeHoleMinDist, oRtePathFeat, oOutfitting)
        Else
            dPipeDia = PipeSimplification(dPipeDia, m_PipeHoleMinDist, oRtePathFeat, oOutfitting)
        End If
    End If
                               
    'Get the insulation thickness
    dInsulationThickness = GetInsulationThickness(oRtePathFeat)
                      
    'Select the hole size according to the joint type, plate tightness
    ' and insulation
    If eJointType = Flange Then
        dHoleSize = dFlangeDia + m_PipeFlangeCr
        dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_PipeFlangeRv)
        
    ElseIf eJointType = SquareFlange Then
        dHoleSize = dFlangeDia + m_PipeSquareFlangeCr
        dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_PipeSquareFlangeRv)
        
    ElseIf eJointType = Sleeve Then
        If bTightness = True Then
            dHoleSize = dFlangeDia + m_PipeSleeveTightCr
            dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_PipeSleeveTightRv)
            
        Else  ' no water-tight integrity.
            If dInsulationThickness > 0 Then
                dHoleSize = dPipeDia + dInsulationThickness * 2 + m_PipeSleeveNoTightInsulationCr
                dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_PipeSleeveNoTightInsulationRv)
            Else
                dHoleSize = dPipeDia + m_PipeSleeveNoTightNoInsulationCr
                dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_PipeSleeveNoTightNoInsulationRv)
            End If
        End If
        
    Else
        If bTightness = True Then
            dHoleSize = dPipeDia + m_PipeNoneTightCr
            dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_PipeNoneTightRv)
            
        Else  ' no water-tight integrity.
            If dInsulationThickness > 0 Then
                dHoleSize = dPipeDia + dInsulationThickness * 2 + m_PipeNoneInsulationCr
                dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_PipeNoneInsulationRv)
            Else
                dHoleSize = dPipeDia + m_PipeNoneNoInsulationCr
                dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_PipeNoneNoInsulationRv)
            End If
        End If
    End If

    'Call IJProjection and represent the pipe
    If TypeOf oPipe Is IJRteTurnPathFeat Then
        
        ' In the event that the feature is a turn feature, two
        ' projections are added to the collection. The first is the
        ' projection created by considering the end points of the turn
        ' feature to be the end points of a straight feature so that the
        ' hole will be sufficiently large that the inside of the bend
        ' will clear.  The second is obtained by considering a straight
        ' feature that would penetrate the plate at same point on the
        ' centerline and at an angle equal to the angle that the turn
        ' feature is actually penetrating the plate.  This assures that
        ' the outside of the bend will clear.
        Set oProjection = CreatePipeProjection(oRtePathFeat, dHoleSize)
        oOutputObjects.Add oProjection
        Set oProjection = CreatePipeTurnProjection(oRtePathFeat, oWorkingPlane, dHoleSize)
    
    Else
            Set oProjection = CreatePipeProjection(oRtePathFeat, dHoleSize)
    
    End If
            
    'Add the projection to the collection
    oOutputObjects.Add oProjection

Cleanup:
    Set oRtePathFeat = Nothing
    Set oProjection = Nothing
    
    Exit Sub
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Sub  ' ComputePipeTrace

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: ComputePipeWithSleeveTrace
'
'Interface: public function
'
'Abstract: Produce a collection of IJProjection objects (cylinders) that
'          represent an individual pipe part
'********************************************************************
Public Function ComputePipeWithSleeveTrace(oHoleTrace As IJHoleTraceAE, _
                                           oOutfitting As IMSCoreCollections.IJDObjectCollection, _
                                           ByRef oOutputObjects As IMSCoreCollections.IJDObjectCollection) As Boolean
    Const Method = "ComputePipeWithSleeveTrace"
    On Error GoTo ErrorHandler
    
    ' Initialize the return value on the assumption of success.
    ComputePipeWithSleeveTrace = True
    
    ' Verify that there is a single item in the outfitting collection.
    If oOutfitting.Count <> 1 Then
        ComputePipeWithSleeveTrace = False
        GoTo Cleanup
    End If
    
    ' Get the part from the hole fitting object
    Dim oPart As IJDPart
    Dim oHoleFitting As IJHoleFitting
    Set oHoleFitting = oHoleTrace.GetHoleFitting
    Set oPart = oHoleFitting.FittingPart
    
    If oPart Is Nothing Then
        ComputePipeWithSleeveTrace = False
        GoTo Cleanup
    End If
    
    ' Get the outside diameter of the pipe sleeve.
    Dim dSleeveOD As Double
    dSleeveOD = GetPipeSleeveOD(oPart)
    
    If dSleeveOD = 0# Then
        ComputePipeWithSleeveTrace = False
        GoTo Cleanup
    End If
    
    ' Get the pipe.  Even though there is only one item in the outfitting
    ' collection I am using a loop because I don't know how to access a
    ' specific item in an IJDObjectCollection.
    Dim oObject As Object
    Dim oRtePathFeat As IJRtePathFeat
    For Each oObject In oOutfitting
        Set oRtePathFeat = oObject
    Next
    
    ' Get the pipe diameter.
    Dim dPipeDia As Double
    dPipeDia = GetPipeDimensions(oRtePathFeat)
    
    ' Check for a reasonable value of sleeve OD wrt pipe diameter.
    If dSleeveOD < dPipeDia Or dSleeveOD > 2 * dPipeDia Then
        ComputePipeWithSleeveTrace = False
        GoTo Cleanup
    End If
    
    ' The size of the hole will be the OD of the pipe sleeve plus a
    ' clearance.
    Dim dHoleSize As Double
    dHoleSize = dSleeveOD + m_PipeSleeveTightCr
    
    ' Create an IJProjection to represent the pipe.
    Dim oProjection As IJProjection
    Set oProjection = CreatePipeProjection(oRtePathFeat, dHoleSize)
            
    ' Add the projection to the output collection.
    oOutputObjects.Add oProjection
    
    GoTo Cleanup

ErrorHandler:
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    
Cleanup:
    Set oPart = Nothing
    Set oObject = Nothing
    Set oRtePathFeat = Nothing
    Set oProjection = Nothing
End Function ' ComputePipeWithSleeveTrace

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: ComputeDuctTrace
'
'Interface: public function
'
'Abstract: Produce a collection of IJProjection objects (cylinders) that
'          represent an individual duct part
'********************************************************************
Public Sub ComputeDuctTrace(oDuct As Object, oOutfitting As IMSCoreCollections.IJDObjectCollection, bTightness As Boolean, _
                            ByRef oOutputObjects As IMSCoreCollections.IJDObjectCollection)
    
    Const Method = "ComputeDuctTrace"
    On Error GoTo ErrorHandler
    
    Dim oRtePathFeat As IJRtePathFeat
    Dim oProjection As IJProjection
    Dim oGeomColl As IMSCoreCollections.IJDObjectCollection
    
    Dim eJointType As JointType
    
    Dim dFlangeDia As Double
    Dim dHoleSize As Double
    Dim dInsulationThickness As Double
    
    Dim eSelectedType As CrossSectionShapeTypes
    Dim dWidth As Double, dDepth As Double, dRadius As Double
    
    Set oRtePathFeat = oDuct
            
    ' get the HVAC and get the flange on the HVAC feature.
    eSelectedType = GetDuctFlangeDimensions(oRtePathFeat, dWidth, dDepth, eJointType)
        
    'If can 't find the flange, and then ask the back and forward duct if they has a flange.
    If dWidth <= 0 Or dDepth <= 0 Then
        eSelectedType = GetDuctDimensions(oRtePathFeat, dWidth, dDepth, dRadius)
    End If
            
    'Get the insulation thickness
    dInsulationThickness = GetInsulationThickness(oRtePathFeat)
        
    'CrossSectionShapeTypes
    '    Rectangular = 1
    '    Oval = 2
    '    FlatOval = 3
    '    Round = 4
    '    UShape = 5
    '    UserDefined = 6
    
    'Following code is for the duct which has a round type.
    'If Duct is  spiral duct, 4 means round type of duct
    If eSelectedType = 4 Then
        ' Compare between width and Depth of duct
        Dim dSelectedDiamater As Double
        dSelectedDiamater = 0#
                     
        If dWidth > dDepth Then
            dSelectedDiamater = dWidth
        Else
            dSelectedDiamater = dDepth
        End If
        
        'Select the hole size according to the joint type and plate tightness and insulation
        'To add the clearance to selected duct's diamater
        If eJointType = Flange Then
            If bTightness = False Then
                If dInsulationThickness > 0 Then
                    dSelectedDiamater = dSelectedDiamater + m_DuctFlangeInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_DuctFlangeInsulationRv)
                Else
                    dSelectedDiamater = dSelectedDiamater + m_DuctFlangeNoInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_DuctFlangeNoInsulationRv)
                End If
            Else
                dSelectedDiamater = dSelectedDiamater + m_DuctFlangeCr
                dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_DuctFlangeRv)
            End If
        ElseIf eJointType = Sleeve Then
            If bTightness = True Then
                dSelectedDiamater = dSelectedDiamater + m_DuctSleeveCr
                dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_DuctSleeveRv)
            Else
                If dInsulationThickness > 0 Then
                    dSelectedDiamater = dSelectedDiamater + dInsulationThickness * 2 + m_DuctSleeveInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_DuctSleeveInsulationRv)
                Else
                    dSelectedDiamater = dSelectedDiamater + m_DuctSleeveNoInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_DuctSleeveNoInsulationRv)
                End If
            End If
        Else
            If bTightness = True Then
                dSelectedDiamater = dSelectedDiamater + m_DuctNoneCr
                dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_DuctNoneRv)
            Else
                If dInsulationThickness > 0 Then
                    dSelectedDiamater = dSelectedDiamater + dInsulationThickness * 2 + m_DuctNoneInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_DuctNoneInsulationRv)
                Else
                    dSelectedDiamater = dSelectedDiamater + m_DuctNoneNoInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_DuctNoneNoInsulationRv)
                End If
            End If
        End If
                                
        'Call IJProjection and represent the duct
        Set oProjection = CreatePipeProjection(oRtePathFeat, dSelectedDiamater)
                
        'Add the projection to the collection
        oOutputObjects.Add oProjection
        
    ElseIf eSelectedType = 3 Then 'duct is flat oval
        'Select the hole size according to the joint type and plate tightness and insulation
        'To add the clearance to selected duct's diamater
        If eJointType = Flange Then
            If bTightness = False Then
                If dInsulationThickness > 0 Then
                    dWidth = dWidth + dInsulationThickness * 2 + m_DuctFlangeInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctFlangeInsulationRv)
                    dDepth = dDepth + dInsulationThickness * 2 + m_DuctFlangeInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctFlangeInsulationRv)
                Else
                    dWidth = dWidth + m_DuctFlangeNoInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctFlangeNoInsulationRv)
                    dDepth = dDepth + m_DuctFlangeNoInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctFlangeNoInsulationRv)
                End If
            Else
                dWidth = dWidth + m_DuctFlangeCr
                dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctFlangeRv)
                dDepth = dDepth + m_DuctFlangeCr
                dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctFlangeRv)
            End If
        ElseIf eJointType = Sleeve Then
            If bTightness = True Then
                dWidth = dWidth + m_DuctSleeveCr
                dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctSleeveRv)
                dDepth = dDepth + m_DuctSleeveCr
                dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctSleeveRv)
            Else
                If dInsulationThickness > 0 Then
                    dWidth = dWidth + dInsulationThickness * 2 + m_DuctSleeveInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctSleeveInsulationRv)
                    dDepth = dDepth + dInsulationThickness * 2 + m_DuctSleeveInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctSleeveInsulationRv)
                Else
                    dWidth = dWidth + m_DuctSleeveNoInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctSleeveNoInsulationRv)
                    dDepth = dDepth + m_DuctSleeveNoInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctSleeveNoInsulationRv)
                End If
            End If
        Else
            If bTightness = True Then
                dWidth = dWidth + m_DuctNoneCr
                dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctNoneRv)
                dDepth = dDepth + m_DuctNoneCr
                dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctNoneRv)
            Else
                If dInsulationThickness > 0 Then
                    dWidth = dWidth + dInsulationThickness * 2 + m_DuctNoneInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctNoneInsulationRv)
                    dDepth = dDepth + dInsulationThickness * 2 + m_DuctNoneInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctNoneInsulationRv)
                Else
                    dWidth = dWidth + m_DuctNoneNoInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctNoneNoInsulationRv)
                    dDepth = dDepth + m_DuctNoneNoInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctNoneNoInsulationRv)
                End If
            End If
        End If
                    
        'Call IJProjection and represent the Duct
        Set oGeomColl = CreateFlatOvalProjection(oRtePathFeat, dWidth, dDepth)

        ' returns the collection of cylinder
        For Each oProjection In oGeomColl
            oOutputObjects.Add oProjection
        Next oProjection
                
        Set oProjection = Nothing
        Set oGeomColl = Nothing
    
    Else 'duct is not round or flat oval
        'Select the hole size according to the joint type and plate tightness and insulation
        'To add the clearance to selected duct's diamater
        If eJointType = Flange Then
            If bTightness = False Then
                If dInsulationThickness > 0 Then
                    dWidth = dWidth + dInsulationThickness * 2 + m_DuctFlangeInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctFlangeInsulationRv)
                    dDepth = dDepth + dInsulationThickness * 2 + m_DuctFlangeInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctFlangeInsulationRv)
                    If dRadius <= 0 Then dRadius = dRadius + dInsulationThickness * 2 + m_DuctFlangeInsulationCr
                Else
                    dWidth = dWidth + m_DuctFlangeNoInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctFlangeNoInsulationRv)
                    dDepth = dDepth + m_DuctFlangeNoInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctFlangeNoInsulationRv)
                    If dRadius <= 0 Then dRadius = dRadius + m_DuctFlangeNoInsulationCr
                End If
            Else
                dWidth = dWidth + m_DuctFlangeCr
                dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctFlangeRv)
                dDepth = dDepth + m_DuctFlangeCr
                dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctFlangeRv)
                If dRadius <= 0 Then dRadius = dRadius + m_DuctFlangeCr
            End If
        ElseIf eJointType = Sleeve Then
            If bTightness = True Then
                dWidth = dWidth + m_DuctSleeveCr
                dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctSleeveRv)
                dDepth = dDepth + m_DuctSleeveCr
                dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctSleeveRv)
                If dRadius <= 0 Then dRadius = dRadius + m_DuctSleeveCr
            Else
                If dInsulationThickness > 0 Then
                    dWidth = dWidth + dInsulationThickness * 2 + m_DuctSleeveInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctSleeveInsulationRv)
                    dDepth = dDepth + dInsulationThickness * 2 + m_DuctSleeveInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctSleeveInsulationRv)
                    If dRadius <= 0 Then dRadius = dRadius + dInsulationThickness * 2 + m_DuctSleeveInsulationCr
                Else
                    dWidth = dWidth + m_DuctSleeveNoInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctSleeveNoInsulationRv)
                    dDepth = dDepth + m_DuctSleeveNoInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctSleeveNoInsulationRv)
                    If dRadius <= 0 Then dRadius = dRadius + m_DuctSleeveNoInsulationCr
                End If
            End If
        Else
            If bTightness = True Then
                dWidth = dWidth + m_DuctNoneCr
                dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctNoneRv)
                dDepth = dDepth + m_DuctNoneCr
                dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctNoneRv)
                If dRadius <= 0 Then dRadius = dRadius + m_DuctNoneCr
            Else
                If dInsulationThickness > 0 Then
                    dWidth = dWidth + dInsulationThickness * 2 + m_DuctNoneInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctNoneInsulationRv)
                    dDepth = dDepth + dInsulationThickness * 2 + m_DuctNoneInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctNoneInsulationRv)
                    If dRadius <= 0 Then dRadius = dRadius + dInsulationThickness * 2 + m_DuctNoneInsulationCr
                Else
                    dWidth = dWidth + m_DuctNoneNoInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_DuctNoneNoInsulationRv)
                    dDepth = dDepth + m_DuctNoneNoInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_DuctNoneNoInsulationRv)
                    If dRadius <= 0 Then dRadius = dRadius + m_DuctNoneNoInsulationCr
                End If
            End If
        End If
                    
        If dRadius <= 0 Then dRadius = m_DuctCorRadius
        
        'Find the four point of HVAC and put the corner radius
        'Call IJProjection and represent the Duct
        Set oGeomColl = CreateDuctProjection(oRtePathFeat, dWidth, dDepth, dRadius)
        
        ' returns the collection of cylinder
        For Each oProjection In oGeomColl
            oOutputObjects.Add oProjection
        Next oProjection
                
        Set oProjection = Nothing
        Set oGeomColl = Nothing
    End If
        
Cleanup:
    Set oRtePathFeat = Nothing
    Set oProjection = Nothing
    
    Exit Sub
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Sub  ' ComputeDuctTrace

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: ComputeCablewayTrace
'
'Interface: public function
'
'Abstract: Produce a collection of IJProjection objects (cylinders) that
'          represent an individual cableway part
'********************************************************************
Public Sub ComputeCablewayTrace(oCableway As Object, oOutfitting As IMSCoreCollections.IJDObjectCollection, bTightness As Boolean, _
                                ByRef oOutputObjects As IMSCoreCollections.IJDObjectCollection)
    Const Method = "ComputeCablewayTrace"
    On Error GoTo ErrorHandler

    Dim oRtePathFeat As IJRtePathFeat
    Dim oProjection As IJProjection
    
    Dim eJointType As JointType
    
    Dim dFlangeDia As Double
    Dim dHoleSize As Double
    Dim dInsulationThickness As Double
    
    Dim dWidth As Double, dDepth As Double
    Dim eSelectedType As CrossSectionShapeTypes
    
    Set oRtePathFeat = oCableway
            
    'get the cableway and get the flange on the cableway feature.
    eSelectedType = GetCablewayFlangeDimensions(oRtePathFeat, dWidth, dDepth, eJointType)
        
    'If can 't find the flange, and then ask the back and forward flange if they has a flange.
    If dWidth <= 0 Or dDepth <= 0 Then
        eSelectedType = GetCablewayDimensions(oRtePathFeat, dWidth, dDepth)
    End If
            
    'Get the insulation thickness
'    dInsulationThickness = GetInsulationThickness(oRtePathFeat)
    dInsulationThickness = 0  'it doesn't appear that cableway have insulation
    
    'Following code is for the cableway which has a round type.
    'If cableway is  spiral cableway, 4 means round type of cableway
    If eSelectedType = 4 Then
        ' Compare between width and Depth of cableway
        Dim dSelectedDiamater As Double
        dSelectedDiamater = 0#
                     
        If dWidth > dDepth Then
            dSelectedDiamater = dWidth
        Else
            dSelectedDiamater = dDepth
        End If
                
        'Select the hole size according to the joint type and plate tightness and insulation
        'To add the clearance to selected cableway's diamater
        If eJointType = Flange Then
            If bTightness = False Then
                If dInsulationThickness > 0 Then
                    dSelectedDiamater = dSelectedDiamater + m_CablewayFlangeInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_CablewayFlangeInsulationRv)
                Else
                    dSelectedDiamater = dSelectedDiamater + m_CablewayFlangeNoInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_CablewayFlangeNoInsulationRv)
                End If
            Else
                dSelectedDiamater = dSelectedDiamater + m_CablewayFlangeCr
                dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_CablewayFlangeRv)
            End If
        ElseIf eJointType = Sleeve Then
            If bTightness = True Then
                dSelectedDiamater = dSelectedDiamater + m_CablewaySleeveCr
                dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_CablewaySleeveRv)
            Else
                If dInsulationThickness > 0 Then
                    dSelectedDiamater = dSelectedDiamater + dInsulationThickness * 2 + m_CablewaySleeveInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_CablewaySleeveInsulationRv)
                Else
                    dSelectedDiamater = dSelectedDiamater + m_CablewaySleeveNoInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_CablewaySleeveNoInsulationRv)
                End If
            End If
        Else
            If bTightness = True Then
                dSelectedDiamater = dSelectedDiamater + m_CablewayNoneCr
                dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_CablewayNoneRv)
            Else
                If dInsulationThickness > 0 Then
                    dSelectedDiamater = dSelectedDiamater + dInsulationThickness * 2 + m_CablewayNoneInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_CablewayNoneInsulationRv)
                Else
                    dSelectedDiamater = dSelectedDiamater + m_CablewayNoneNoInsulationCr
                    dSelectedDiamater = GetHoleSizeByRoundValue(dSelectedDiamater, m_CablewayNoneNoInsulationRv)
                End If
            End If
        End If
                                
        'Call IJProjection and represent the cableway
        Set oProjection = CreatePipeProjection(oRtePathFeat, dSelectedDiamater)
                
        'Add the projection to the collection
        oOutputObjects.Add oProjection

    Else 'cableway is not round
        Dim oFourCylinderCollection As IMSCoreCollections.IJDObjectCollection
                    
        'Select the hole size according to the joint type and plate tightness and insulation
        'To add the clearance to selected cableway's diamater
        If eJointType = Flange Then
            If bTightness = False Then
                If dInsulationThickness > 0 Then
                    dWidth = dWidth + dInsulationThickness * 2 + m_CablewayFlangeInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_CablewayFlangeInsulationRv)
                    dDepth = dDepth + dInsulationThickness * 2 + m_CablewayFlangeInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_CablewayFlangeInsulationRv)
                Else
                    dWidth = dWidth + m_CablewayFlangeNoInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_CablewayFlangeNoInsulationRv)
                    dDepth = dDepth + m_CablewayFlangeNoInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_CablewayFlangeNoInsulationRv)
                End If
            Else
                dWidth = dWidth + m_CablewayFlangeCr
                dWidth = GetHoleSizeByRoundValue(dWidth, m_CablewayFlangeRv)
                dDepth = dDepth + m_CablewayFlangeCr
                dDepth = GetHoleSizeByRoundValue(dDepth, m_CablewayFlangeRv)
            End If
        ElseIf eJointType = Sleeve Then
            If bTightness = True Then
                dWidth = dWidth + m_CablewaySleeveCr
                dWidth = GetHoleSizeByRoundValue(dWidth, m_CablewaySleeveRv)
                dDepth = dDepth + m_CablewaySleeveCr
                dDepth = GetHoleSizeByRoundValue(dDepth, m_CablewaySleeveRv)
            Else
                If dInsulationThickness > 0 Then
                    dWidth = dWidth + dInsulationThickness * 2 + m_CablewaySleeveInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_CablewaySleeveInsulationRv)
                    dDepth = dDepth + dInsulationThickness * 2 + m_CablewaySleeveInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_CablewaySleeveInsulationRv)
                Else
                    dWidth = dWidth + m_CablewaySleeveNoInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_CablewaySleeveNoInsulationRv)
                    dDepth = dDepth + m_CablewaySleeveNoInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_CablewaySleeveNoInsulationRv)
                End If
            End If
        Else
            If bTightness = True Then
                dWidth = dWidth + m_CablewayNoneCr
                dWidth = GetHoleSizeByRoundValue(dWidth, m_CablewayNoneRv)
                dDepth = dDepth + m_CablewayNoneCr
                dDepth = GetHoleSizeByRoundValue(dDepth, m_CablewayNoneRv)
            Else
                If dInsulationThickness > 0 Then
                    dWidth = dWidth + dInsulationThickness * 2 + m_CablewayNoneInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_CablewayNoneInsulationRv)
                    dDepth = dDepth + dInsulationThickness * 2 + m_CablewayNoneInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_CablewayNoneInsulationRv)
                Else
                    dWidth = dWidth + m_CablewayNoneNoInsulationCr
                    dWidth = GetHoleSizeByRoundValue(dWidth, m_CablewayNoneNoInsulationRv)
                    dDepth = dDepth + m_CablewayNoneNoInsulationCr
                    dDepth = GetHoleSizeByRoundValue(dDepth, m_CablewayNoneNoInsulationRv)
                End If
            End If
        End If
                    
        'Find the four point of HVAC and put the corner radius
        'Call IJProjection and represent the cableway
        Set oFourCylinderCollection = CreateDuctProjection(oRtePathFeat, dWidth, dDepth, m_CablewayCorRadius)

        ' returns the collection of cylinder
        Dim oCylinderProjection As IJProjection
        For Each oCylinderProjection In oFourCylinderCollection
            oOutputObjects.Add oCylinderProjection
        Next oCylinderProjection
                
        Set oCylinderProjection = Nothing
        Set oFourCylinderCollection = Nothing
    End If
        
Cleanup:
    Set oRtePathFeat = Nothing
    Set oProjection = Nothing
    
    Exit Sub

ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Sub  ' ComputeCablewayTrace

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: ComputeConduitTrace
'
'Interface: public function
'
'Abstract: Produce a collection of IJProjection objects (cylinders) that
'          represent an individual conduit part
'********************************************************************
Public Sub ComputeConduitTrace(oConduit As Object, oOutfitting As IMSCoreCollections.IJDObjectCollection, bTightness As Boolean, _
                               ByRef oOutputObjects As IMSCoreCollections.IJDObjectCollection)
    Const Method = "ComputeConduitTrace"
    On Error GoTo ErrorHandler

    Dim oRtePathFeat As IJRtePathFeat
    Dim oProjection As IJProjection
    
    Dim eJointType As JointType
    
    Dim dFlangeDia As Double
    Dim dConduitDia As Double
    Dim dHoleSize As Double
    
    Set oRtePathFeat = oConduit
    
    'To get the flange size on the conduit feature,
    ' If can't find the flange on the conduit feature, and then ask the back and forward conduits
    ' on the conduit run if they has a flange.
    dFlangeDia = GetConduitFlangeDimensions(oRtePathFeat, eJointType)
    
    'If there is no flange on the conduit run, the conduit size should be returned.
    If dFlangeDia <= 0 Then
        dConduitDia = GetConduitDimensions(oRtePathFeat)
    End If
    
    'To check the simlipication among less then three conduits, change the radius due to Simplification rule
'    If oOutfitting.Count <= 3 Then
'        If dFlangeDia > 0 Then
'            dFlangeDia = ConduitSimplification(dFlangeDia, m_ConduitHoleMinDist, oRtePathFeat, oOutfitting)
'        Else
'            dPipeDia = ConduitSimplification(dConduitDia, m_ConduitHoleMinDist, oRtePathFeat, oOutfitting)
'        End If
'    End If
                               
    'Select the hole size according to the joint type and plate tightness
    If eJointType = Flange Then
        dHoleSize = dFlangeDia + m_ConduitFlangeCr
        dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_ConduitFlangeRv)
    ElseIf eJointType = SquareFlange Then
        dHoleSize = dFlangeDia + m_ConduitSquareFlangeCr
        dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_ConduitSquareFlangeRv)
    ElseIf eJointType = Sleeve Then
        If bTightness = True Then
            dHoleSize = dFlangeDia + m_ConduitSleeveTightCr
            dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_ConduitSleeveTightRv)
        Else
            dHoleSize = dConduitDia + m_ConduitSleeveNoTightCr
            dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_ConduitSleeveNoTightRv)
        End If
    Else
        If bTightness = True Then
            dHoleSize = dConduitDia + m_ConduitNoneTightCr
            dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_ConduitNoneTightRv)
        Else
            dHoleSize = dConduitDia + m_ConduitNoneCr
            dHoleSize = GetHoleSizeByRoundValue(dHoleSize, m_ConduitNoneRv)
        End If
    End If

    'Call IJProjection and represent the conduit
    Set oProjection = CreatePipeProjection(oRtePathFeat, dHoleSize)
            
    'Add the projection to the collection
    oOutputObjects.Add oProjection

Cleanup:
    Set oRtePathFeat = Nothing
    Set oProjection = Nothing

    Exit Sub
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
    
End Sub  'ComputeConduitTrace

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: ComputeCenterFlangeTrace
'
'Interface: public function
'
'Abstract: Produce a collection of IJProjection object ( Cylinders ) that
'          represent the outfitting part
'
'Note: this is being used to place the cuts in a center flangeplate fitting (double ring or
'      center flange) in which there is not a port yet for the plate. so if
'      the oStructure is nothing then there is no PlateThickness
'********************************************************************
Public Function ComputeCenterFlangeTrace(oStructure As Object, oOutfitting As IMSCoreCollections.IJDObjectCollection, _
                                         ByRef oOutputObjects As IMSCoreCollections.IJDObjectCollection) As Boolean
    Const Method = "ComputeCenterFlangeTrace"
    On Error GoTo ErrorHandler
    
    Dim oObj As Object
    Dim oRtePathFeat As IJRtePathFeat
    Dim oProjection As IJProjection
    Dim oGeomColl As IMSCoreCollections.IJDObjectCollection
    
    oOutputObjects.Clear
    ComputeCenterFlangeTrace = True
        
    'Get the plate tightness
    Dim bPlateTigthness As Boolean
    bPlateTigthness = PlateTightness(oStructure)
        
    'For Loop input object ( 1 .. icount .. n )
    For Each oObj In oOutfitting
        If TypeOf oObj Is IJRtePipePathFeat Then
            Set oRtePathFeat = oObj
          
            'Get the pipe diameter
            Dim dPipeDia As Double
            dPipeDia = GetPipeDimensions(oRtePathFeat)

            'Call IJProjection and represent the pipe
            Set oProjection = CreatePipeProjection(oRtePathFeat, dPipeDia)
                
            'Add the projection to the collection
            oOutputObjects.Add oProjection

            Set oRtePathFeat = Nothing

        ElseIf TypeOf oObj Is IJRteDuctPathFeat Then
            Set oRtePathFeat = oObj
                
            ' get the HVAC dimensions
            Dim eSelectedType As CrossSectionShapeTypes
            Dim dWidth As Double, dDepth As Double, dRadius As Double

            eSelectedType = GetDuctDimensions(oRtePathFeat, dWidth, dDepth, dRadius)
            If eSelectedType = 4 Then   'duct is round
                ' Compare between width and Depth of duct
                Dim dSelectedDiamater As Double
                dSelectedDiamater = 0#
                     
                If dWidth > dDepth Then
                    dSelectedDiamater = dWidth
                Else
                    dSelectedDiamater = dDepth
                End If
                                               
                'Call IJProjection and represent the pipe
                Set oProjection = CreatePipeProjection(oRtePathFeat, dSelectedDiamater)
                
                'Add the projection to the collection
                oOutputObjects.Add oProjection

            ElseIf eSelectedType = 3 Then  'duct is flat oval
                'Call IJProjection and represent the Duct
                Set oGeomColl = CreateFlatOvalProjection(oRtePathFeat, dWidth, dDepth)

                ' returns the collection of cylinder
                For Each oProjection In oGeomColl
                    oOutputObjects.Add oProjection
                Next oProjection
            
                Set oProjection = Nothing
                Set oGeomColl = Nothing
            Else
                'Find the four point of HVAC and put the corner radius
                'Call IJProjection and represent the Duct
                If dRadius <= 0 Then dRadius = 0.025
                Set oGeomColl = CreateDuctProjection(oRtePathFeat, dWidth, dDepth, dRadius)

                ' returns the collection of cylinder
                For Each oProjection In oGeomColl
                    oOutputObjects.Add oProjection
                Next oProjection
                
                Set oProjection = Nothing
                Set oGeomColl = Nothing
            End If
            
        ElseIf TypeOf oObj Is IJEquipment Or TypeOf oObj Is IJPort Then
            'If object is equipment, convert equipment to port, and returns port object
            ComputeCenterFlangeTrace = False
        
        End If
    Next oObj
        
Cleanup:
    Set oRtePathFeat = Nothing
    Set oObj = Nothing
    Set oProjection = Nothing
    
    Exit Function
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Function  ' ComputeCenterFlangeTrace

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: GetDatabaseClearances
'
'Interface: public function
'
'Abstract:
'********************************************************************
Private Sub GetDatabaseClearances(oHoleTrace As IJHoleTraceAE)
    Const Method = "GetDatabaseClearances"
    On Error GoTo ErrorHandler
    
    Dim oSRDQuery As IJSRDQuery
    Dim oRuleQuery As IJSRDRuleQuery
    Dim oRule As IJSRDRule
    
    'Get the Clearance Rule from the HoleTrace object
    oHoleTrace.GetHoleClearanceRule oRule
    
    Dim oRuleObject As Object
    'CoCreate the rule progid
    Set oRuleObject = SP3DCreateObject(oRule.ProgId)
    If oRuleObject Is Nothing Then GoTo Cleanup
 
    
    Dim oPipeClearance As IJHMPipeClearance
    Dim oDuctClearance As IJHMDuctClearance
    Dim oCablewayClearance As IJHMCablewayClearance
    Dim oConduitClearance As IJHMConduitClearance
    
    Set oPipeClearance = oRuleObject
    Set oDuctClearance = oRuleObject
    Set oCablewayClearance = oRuleObject
    Set oConduitClearance = oRuleObject
    
    'For the nearest distance between two holes.
    m_PipeHoleMinDist = oPipeClearance.HoleMinDist
    
    'For Pipe Clearance
    m_PipeFlangeCr = oPipeClearance.FlangeClr
    m_PipeSquareFlangeCr = oPipeClearance.SquareFlangeClr
    m_PipeSleeveTightCr = oPipeClearance.SleeveClr(True, True)
    m_PipeSleeveNoTightInsulationCr = oPipeClearance.SleeveClr(False, True)
    m_PipeSleeveNoTightNoInsulationCr = oPipeClearance.SleeveClr(False, False)
    m_PipeNoneTightCr = oPipeClearance.OtherClr(True, True, 0)
    m_PipeNoneInsulationCr = oPipeClearance.OtherClr(False, True, 0)
    m_PipeNoneNoInsulationCr = oPipeClearance.OtherClr(False, False, 0)

    'For Pipe Round Value
    m_PipeFlangeRv = oPipeClearance.FlangeRound
    m_PipeSquareFlangeRv = oPipeClearance.SquareFlangeRound
    m_PipeSleeveTightRv = oPipeClearance.SleeveRound(True, True)
    m_PipeSleeveNoTightInsulationRv = oPipeClearance.SleeveRound(False, True)
    m_PipeSleeveNoTightNoInsulationRv = oPipeClearance.SleeveRound(False, False)
    m_PipeNoneTightRv = oPipeClearance.OtherRound(True, True, 0)
    m_PipeNoneInsulationRv = oPipeClearance.OtherRound(False, True, 0)
    m_PipeNoneNoInsulationRv = oPipeClearance.OtherRound(False, False, 0)

    'For Duct Clearance
    m_DuctFlangeCr = oDuctClearance.FlangeClr(True, True)
    m_DuctFlangeInsulationCr = oDuctClearance.FlangeClr(False, True)
    m_DuctFlangeNoInsulationCr = oDuctClearance.FlangeClr(False, False)
    m_DuctSleeveCr = oDuctClearance.SleeveClr(True, True)
    m_DuctSleeveInsulationCr = oDuctClearance.SleeveClr(False, True)
    m_DuctSleeveNoInsulationCr = oDuctClearance.SleeveClr(False, False)
    m_DuctNoneCr = oDuctClearance.OtherClr(True, True, 0)
    m_DuctNoneInsulationCr = oDuctClearance.OtherClr(False, True, 0)
    m_DuctNoneNoInsulationCr = oDuctClearance.OtherClr(False, False, 0)

    'For Duct corner radius
    m_DuctCorRadius = oDuctClearance.CornerRadius

    'For Duct Round Value
    m_DuctFlangeRv = oDuctClearance.FlangeRound(True, True)
    m_DuctFlangeInsulationRv = oDuctClearance.FlangeRound(False, True)
    m_DuctFlangeNoInsulationRv = oDuctClearance.FlangeRound(False, False)
    m_DuctSleeveRv = oDuctClearance.SleeveRound(True, True)
    m_DuctSleeveInsulationRv = oDuctClearance.SleeveRound(False, True)
    m_DuctSleeveNoInsulationRv = oDuctClearance.SleeveRound(False, False)
    m_DuctNoneRv = oDuctClearance.OtherRound(True, True, 0)
    m_DuctNoneInsulationRv = oDuctClearance.OtherRound(False, True, 0)
    m_DuctNoneNoInsulationRv = oDuctClearance.OtherRound(False, False, 0)
    
    'For Cableway Clearance
    m_CablewayFlangeCr = oCablewayClearance.FlangeClr(True, True)
    m_CablewayFlangeInsulationCr = oCablewayClearance.FlangeClr(False, True)
    m_CablewayFlangeNoInsulationCr = oCablewayClearance.FlangeClr(False, False)
    m_CablewaySleeveCr = oCablewayClearance.SleeveClr(True, True)
    m_CablewaySleeveInsulationCr = oCablewayClearance.SleeveClr(False, True)
    m_CablewaySleeveNoInsulationCr = oCablewayClearance.SleeveClr(False, False)
    m_CablewayNoneCr = oCablewayClearance.OtherClr(True, True, 0)
    m_CablewayNoneInsulationCr = oCablewayClearance.OtherClr(False, True, 0)
    m_CablewayNoneNoInsulationCr = oCablewayClearance.OtherClr(False, False, 0)

    'For Cableway corner radius
    m_CablewayCorRadius = oCablewayClearance.CornerRadius
    
    'For Cableway Round Value
    m_CablewayFlangeRv = oCablewayClearance.FlangeRound(True, True)
    m_CablewayFlangeInsulationRv = oCablewayClearance.FlangeRound(False, True)
    m_CablewayFlangeNoInsulationRv = oCablewayClearance.FlangeRound(False, False)
    m_CablewaySleeveRv = oCablewayClearance.SleeveRound(True, True)
    m_CablewaySleeveInsulationRv = oCablewayClearance.SleeveRound(False, True)
    m_CablewaySleeveNoInsulationRv = oCablewayClearance.SleeveRound(False, False)
    m_CablewayNoneRv = oCablewayClearance.OtherRound(True, True, 0)
    m_CablewayNoneInsulationRv = oCablewayClearance.OtherRound(False, True, 0)
    m_CablewayNoneNoInsulationRv = oCablewayClearance.OtherRound(False, False, 0)
    
    'For the nearest distance between two holes
    m_ConduitHoleMinDist = oConduitClearance.HoleMinDist
   
    'For Conduit Clearance (note that conduits do not have insulation)
    m_ConduitFlangeCr = oConduitClearance.FlangeClr
    m_ConduitSquareFlangeCr = oConduitClearance.SquareFlangeClr
    m_ConduitSleeveTightCr = oConduitClearance.SleeveClr(True)
    m_ConduitSleeveNoTightCr = oConduitClearance.SleeveClr(False)
    m_ConduitNoneTightCr = oConduitClearance.OtherClr(True, 0)
    m_ConduitNoneCr = oConduitClearance.OtherClr(False, 0)

    'For Conduit Round Value
    m_ConduitFlangeRv = oConduitClearance.FlangeRound
    m_ConduitSquareFlangeRv = oConduitClearance.SquareFlangeRound
    m_ConduitSleeveTightRv = oConduitClearance.SleeveRound(True)
    m_ConduitSleeveNoTightRv = oConduitClearance.SleeveRound(False)
    m_ConduitNoneTightRv = oConduitClearance.OtherRound(True, 0)
    m_ConduitNoneRv = oConduitClearance.OtherRound(False, 0)

Cleanup:
    Set oRule = Nothing
    Set oSRDQuery = Nothing
    Set oRuleQuery = Nothing
    Set oRuleObject = Nothing
    
    Set oPipeClearance = Nothing
    Set oDuctClearance = Nothing
    Set oCablewayClearance = Nothing
    Set oConduitClearance = Nothing
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, MODULE & " - " & Method, Err.Description
    GoTo Cleanup
End Sub  ' GetDatabaseClearances
 
