Attribute VB_Name = "StructSupServiceUtils"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   StructSupServiceUtils.bas
'   ProgID:
'   Author:         Yan Ji
'   Creation Date:  28.Feb.2003
'   Description:
'
'
'   Change History:
'     Date              Name               Description
'   28.Feb.2003         Yan Ji             Creation Date
'   24.May.2006       Chinta,Mahesh        DI CP:95927(Address the code review comments for  PF Assembly Common Modules)
'   02.Jun.2006       Ravi ippili          DI-CP·96213 Modified the GetAngleCSDimension to get the backtoback distance from hanger beam
'   Nov 29, 2010      Ramya                CR-CP-190404  Modified GetIsLugEndOffsetApplied, GetIndexedStructPortName for Place By Refrence command
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const MODULE = "AssemblyInformationRuleUtils"
Private Const SUPPORT_RECOMPUTE = 4

Public Const EPSILON = 0.00001
'Public Const Pi = 3.1415926
Public Enum GBBConfigOrientType
    GBBVertical
    GBBHorizontal
End Enum
' Define 2 structures in order along route direction (Route-X)
Public Enum StructOrderType
    LegtRightStructsAlongRoute     ' 2 structure from Left to Right along Y-axis of route
    RearFrontStructsAlongRoute     ' 2 structure from Rear to Front along X-axis of route
End Enum

'*************************************************************************************
' Method:   AddConnPortBetweenStructPorts
'
' Description:    This function adds a connection port between 2 structure ports so that some support component
' can be positioined properly base on this connection port position
'
' Parameters:
'           pInputConfigHlpr -   Input (Object Type)
'           JointCollection  -   Input( Collection)
'           Idx_Connection   -   Input index for "Connnection" part in part class list
'
'
'*************************************************************************************

Public Function AddConnPortBetweenStructPorts(ByVal pInputConfigHlpr As Object, _
                                              ByVal JointCollection As Collection, _
                                              ByVal Idx_Connection As Long)

    Const METHOD = "AddConnPortBetweenStructPorts:"
    On Error GoTo ErrorHandler
    
    If pInputConfigHlpr Is Nothing Or JointCollection Is Nothing Then
        Exit Function
    End If
    
    'Get IJHgrInputConfig Hlpr Interface off of passed Helper
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pInputConfigHlpr
    
    'Get ports.
    Dim oRefPortColl    As IJElements
    Dim oObject         As Object
    Dim oHgrPort        As IJHgrPort
    Dim oStructPort1    As IJHgrPort
    Dim oStructPort2    As IJHgrPort
    Dim offset          As Double
    Dim StructYaxis1    As New DVector
    Dim AssemblyJoint   As Object
    Dim JointFactory    As New HgrJointFactory
    
    ' Get reference port collection
    Set oRefPortColl = my_IJHgrInputConfigHlpr.GetRefPorts()
    
    ' Extract expected ports from oRefPortColl
    For Each oObject In oRefPortColl
        Set oHgrPort = oObject
        If oHgrPort.name = "Structure" Then
            Set oStructPort1 = oHgrPort
        ElseIf oHgrPort.name = "Struct_2" Then
            Set oStructPort2 = oHgrPort
        End If
        Set oHgrPort = Nothing
    Next
      
    If oStructPort1 Is Nothing Or oStructPort1.name <> "Structure" Then
        Exit Function
    End If
    
    If oStructPort2 Is Nothing Then
        Set AssemblyJoint = JointFactory.MakeRigidJoint(-1, oStructPort1.name, Idx_Connection, "Connection")
        JointCollection.Add AssemblyJoint
        Exit Function
    End If
    
    Dim StructOrigin1   As New DPosition    ' position for structure port 1
    Dim StructXaxis1    As New DVector
    Dim StructZaxis1    As New DVector
    
    Dim StructOrigin2   As New DPosition    ' position for structure port 2
    Dim StructXaxis2    As New DVector
    Dim StructZaxis2    As New DVector
    
    Dim StructVec       As New DVector
    Dim OriginX As Double, OriginY As Double, OriginZ As Double
    Dim XaxisX As Double, XaxisY As Double, XaxisZ As Double
    Dim ZaxisX As Double, ZaxisY As Double, ZaxisZ As Double
    
    ' Get orientation from structure prot 1
    oStructPort1.GetOrientation OriginX, OriginY, OriginZ, _
                                XaxisX, XaxisY, XaxisZ, _
                                ZaxisX, ZaxisY, ZaxisZ
    StructOrigin1.Set OriginX, OriginY, OriginZ
    StructXaxis1.Set XaxisX, XaxisY, XaxisZ
    StructZaxis1.Set ZaxisX, ZaxisY, ZaxisZ
    
    ' Get orientation from structure prot 2
    oStructPort2.GetOrientation OriginX, OriginY, OriginZ, _
                                XaxisX, XaxisY, XaxisZ, _
                                ZaxisX, ZaxisY, ZaxisZ
    StructOrigin2.Set OriginX, OriginY, OriginZ
    StructXaxis2.Set XaxisX, XaxisY, XaxisZ
    StructZaxis2.Set ZaxisX, ZaxisY, ZaxisZ
    
    ' Genetate a vector from the origin of structure port 1 to the origin of port 2
    StructVec.Set StructOrigin2.x - StructOrigin1.x, _
                  StructOrigin2.y - StructOrigin1.y, _
                  StructOrigin2.z - StructOrigin1.z
        
    ' Generate Y-axis of port 1
    Set StructYaxis1 = StructZaxis1.Cross(StructXaxis1)
    
    ' Calculate offset by projecting StructVec onto Y-axis of port 1.
    offset = 0.5 * StructYaxis1.Dot(StructVec)
    
    ' Add connection port at the middle point between 2 ports
    Set AssemblyJoint = JointFactory.MakeRigidJoint(-1, oStructPort1.name, Idx_Connection, "Connection", 10596, 0#, 0#, offset)
    JointCollection.Add AssemblyJoint
    
    Set my_IJHgrInputConfigHlpr = Nothing
    Set oObject = Nothing
    Set oRefPortColl = Nothing
    Set oStructPort1 = Nothing
    Set oStructPort2 = Nothing
    Set StructYaxis1 = Nothing
    Set AssemblyJoint = Nothing
    Set JointFactory = Nothing
    Set StructOrigin1 = Nothing
    Set StructOrigin2 = Nothing
    Set StructXaxis1 = Nothing
    Set StructXaxis2 = Nothing
    Set StructZaxis1 = Nothing
    Set StructZaxis2 = Nothing
    Set StructVec = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetStructPortNamesInOrder
'
' Desc:    This function is used to get structure port names in order so that, either from left to right
' or from rear to front based on input order type. Basically, it can be used to check relative
' orientation or position of 2 structures about the route
'
'  Param:
'           pInputConfigHlpr -
'           BBX_Name -      such as "Route", "BBR_Low", etc.
'           OrientFlag -    order type, either from left to right along Y-axis of route, or
'                           from rear to front along X-axis of route.
'                           Value: LegtRightStructsAlongRoute / RearFrontStructsAlongRoute
'
'
'
'*************************************************************************************

Public Function GetStructPortNamesInOrder(ByVal pInputConfigHlpr As Object, _
                                          ByVal RoutePortName As String, _
                                          ByVal OrientFlag As StructOrderType) As String()

    Const METHOD = "GetStructPortNamesInOrder:"
    On Error GoTo ErrorHandler
    
    'Get IJHgrInputConfig Hlpr Interface off of passed Helper
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pInputConfigHlpr

    Dim oStructs        As Object
    Dim oStruct1        As Object
    Dim oStruct2        As Object
    Dim oObject         As Object
    Dim oHgrPort        As IJHgrPort
    Dim oRoutePort      As IJHgrPort
    Dim RouteXaxis      As New DVector
    Dim RouteZaxis      As New DVector
    Dim StructVec       As New DVector
    Dim StructCnt1      As New DPosition
    Dim StructCnt2      As New DPosition
    Dim DotProduct      As Double
    Dim OriginX As Double, OriginY  As Double, OriginZ  As Double
    Dim XaxisX  As Double, XaxisY   As Double, XaxisZ   As Double
    Dim ZaxisX  As Double, ZaxisY   As Double, ZaxisZ   As Double
    
    Dim StructPorts(1 To 2) As String
    Dim oRefPortColl        As IJElements

    ' Get structure collection
    my_IJHgrInputConfigHlpr.GetStructure oStructs
    
    If oStructs.Count < 1 Then
        Exit Function
    End If
    
    ' If only one structure input, set structure names and return
    If oStructs.Count = 1 Then
        StructPorts(1) = "Structure"
        StructPorts(2) = "Structure"
        GetStructPortNamesInOrder = StructPorts
        Exit Function
    End If
        
    ' We sure there are 2 structures input
    Set oStruct1 = oStructs.Item(1)
    Set oStruct2 = oStructs.Item(2)

    ' Get geometry centers from 2 structures, respectively
    Set StructCnt1 = my_IJHgrInputConfigHlpr.GetStructGeomCenter(oStruct1)
    Set StructCnt2 = my_IJHgrInputConfigHlpr.GetStructGeomCenter(oStruct2)
                
    ' Create a vector from structure 1 to structure 2
    StructVec.Set StructCnt2.x - StructCnt1.x, _
                  StructCnt2.y - StructCnt1.y, _
                  StructCnt2.z - StructCnt1.z
    
    ' Get route port based on input name
    Set oRefPortColl = my_IJHgrInputConfigHlpr.GetRefPorts()
    For Each oObject In oRefPortColl
        Set oHgrPort = oObject
        If oHgrPort.name = RoutePortName Then '"Route" Then
            Set oRoutePort = oHgrPort
        End If
        Set oHgrPort = Nothing
    Next
    
    If oRoutePort Is Nothing Then
        Exit Function
    End If

    ' Get orientation from route port
    oRoutePort.GetOrientation OriginX, OriginY, OriginZ, _
                                XaxisX, XaxisY, XaxisZ, _
                                ZaxisX, ZaxisY, ZaxisZ
    RouteXaxis.Set XaxisX, XaxisY, XaxisZ
    RouteZaxis.Set ZaxisX, ZaxisY, ZaxisZ
                
    ' Check orientation of 2 structures about route
    If OrientFlag = LegtRightStructsAlongRoute Then
        Dim RouteYaxis      As New DVector
        Set RouteYaxis = RouteZaxis.Cross(RouteXaxis)
        DotProduct = RouteYaxis.Dot(StructVec)
    Else
        DotProduct = RouteXaxis.Dot(StructVec)
    End If
    
    ' Set output for structure port names in required order
    If DotProduct > 0# Then
        StructPorts(1) = "Structure"
        StructPorts(2) = "Struct_2"
    Else
        StructPorts(1) = "Struct_2"
        StructPorts(2) = "Structure"
    End If
    
    GetStructPortNamesInOrder = StructPorts
    
    Set my_IJHgrInputConfigHlpr = Nothing
    Set oStructs = Nothing
    Set oStruct1 = Nothing
    Set oStruct2 = Nothing
    Set oObject = Nothing
    Set oRoutePort = Nothing
    Set RouteXaxis = Nothing
    Set RouteZaxis = Nothing
    Set StructCnt1 = Nothing
    Set StructCnt2 = Nothing
                   
 Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   AttachConnectionsToGBBR
'
' Desc:    This method used to attach Connection ports to GBBR_Z so that user
' can use them as desired Low and High ports.
'
'  Param:
'           Output: Width and Height of bounding box as desired
'
'
'
'*************************************************************************************
Public Function AttachConnectionsToGBBR( _
                            ByVal pDispInputConfigHlpr As Object, _
                            ByVal JointCollection As Collection, _
                            ByVal LowConnIdx As Long, _
                            ByVal HighConnIdx As Long, _
                            ByVal enumOrient As GBBConfigOrientType) As Double()
    
    Const METHOD = "AttachConnectionsToGBBR:"
    On Error GoTo ErrorHandler
    
    If enumOrient <> GBBHorizontal And enumOrient <> GBBVertical Then
        Exit Function
    End If
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr

    Dim dimension(1 To 2) As Double
    dimension(1) = 0#
    dimension(2) = 0#
    
    ' Get bounding box dimension
    Dim IJElements_BBInfo As IJElements
    Dim BB_Position As IJDPosition
    Dim oBBInfo As Object
    Dim oBBHigh As Object
    Dim dWidth As Double
    Dim dHeight As Double
    Dim Config_Idx_Low As Long
    Dim Config_Idx_High As Long
    Dim plnOffset As Double
    Dim axisOffset As Double
    
    Set oBBInfo = my_IJHgrInputConfigHlpr.GetBoundingBoxInfo("GBBR_Z")
    Set IJElements_BBInfo = oBBInfo
    Set oBBHigh = IJElements_BBInfo.Item(1)
    Set BB_Position = oBBHigh
    dWidth = BB_Position.x
    dHeight = BB_Position.y
    
    ' Get orienntation information
    Dim dAngleRZ_GZ As Double
    Dim dAngleRY_GZ As Double
    Dim dAngleSZ_GZ As Double
    dAngleRZ_GZ = my_IJHgrInputConfigHlpr.GetPortOrientAngle(ORIENT_DIRECT, "GBBR_Z_Low", 3, "WORLD", 3)
    dAngleRY_GZ = my_IJHgrInputConfigHlpr.GetPortOrientAngle(ORIENT_DIRECT, "Route", 2, "WORLD", 3)
    dAngleSZ_GZ = my_IJHgrInputConfigHlpr.GetPortOrientAngle(ORIENT_DIRECT, "Structure", 3, "WORLD", 3)
    
    'Create a Joint Factory
    Dim JointFactory As New HgrJointFactory
    Dim AssemblyJoint As Object
    Dim PI2 As Double
    PI2 = 0.5 * (4 * Atn(1))  '4 * atn(1) =Pi
    

    Select Case enumOrient
        Case GBBHorizontal
            Set AssemblyJoint = JointFactory.MakeRigidJoint(-1, "GBBR_Z_Low", LowConnIdx, "Connection")
            JointCollection.Add AssemblyJoint
            Set AssemblyJoint = JointFactory.MakeRigidJoint(-1, "GBBR_Z_High", HighConnIdx, "Connection")
            JointCollection.Add AssemblyJoint
            dimension(1) = dWidth
            dimension(2) = dHeight
            
        Case GBBVertical
            'If Abs(dAngleRZ_GZ) < EPSILON Or Abs(dAngleRZ_GZ - PI) < EPSILON Then
            If Abs(dAngleRY_GZ - PI2) < EPSILON Then
                dimension(1) = dWidth
                dimension(2) = dHeight
                Set AssemblyJoint = JointFactory.MakeRigidJoint(-1, "GBBR_Z_Low", LowConnIdx, "Connection", 9444)
                JointCollection.Add AssemblyJoint
                Set AssemblyJoint = JointFactory.MakeRigidJoint(-1, "GBBR_Z_High", HighConnIdx, "Connection", 9444)
                JointCollection.Add AssemblyJoint
            Else
                dimension(1) = dHeight
                dimension(2) = dWidth
                If (dAngleSZ_GZ > PI2 And dAngleRY_GZ >= PI2) Or _
                   (dAngleSZ_GZ <= PI2 And dAngleRY_GZ < PI2) Then
                    
                    Config_Idx_Low = 9452
                    plnOffset = 0#
                    axisOffset = dWidth
                Else
                    Config_Idx_Low = 9388
                    plnOffset = dHeight
                    axisOffset = 0#
                End If
                Set AssemblyJoint = JointFactory.MakeRigidJoint(-1, "GBBR_Z_Low", LowConnIdx, "Connection", Config_Idx_Low, plnOffset, axisOffset)
                JointCollection.Add AssemblyJoint
                Set AssemblyJoint = JointFactory.MakeRigidJoint(-1, "GBBR_Z_High", HighConnIdx, "Connection", Config_Idx_High, -plnOffset, -axisOffset)
                JointCollection.Add AssemblyJoint
            End If
    End Select
    
    AttachConnectionsToGBBR = dimension
    
    Set my_IJHgrInputConfigHlpr = Nothing
    Set BB_Position = Nothing
    Set oBBInfo = Nothing
    Set oBBHigh = Nothing
    Set JointFactory = Nothing
    Set AssemblyJoint = Nothing
    
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetMaxExternalPipeDiameter
'
' Desc:    This method retrieved the index of the pipe with max diameter among multi pipes.
' Param:
'           Output: route index
'
'
'*************************************************************************************

Public Function GetMaxExternalPipeDiameter(ByVal pDispInputConfigHlpr As Object) As Double
    
    Const METHOD = "GetExternalPipeDiameter:"
    On Error GoTo ErrorHandler
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
      
    Dim dblPipeDia                  As Double
    Dim dblMaxDia                   As Double
    Dim Index                       As Integer
    Dim numPipes                    As Integer
    Dim oPipeCollection             As Object
    Dim IJElements_PipeCollection   As IJElements
    
    my_IJHgrInputConfigHlpr.GetPipes oPipeCollection
    Set IJElements_PipeCollection = oPipeCollection
    numPipes = IJElements_PipeCollection.Count
    
    dblMaxDia = 0#
    For Index = 1 To numPipes
        dblPipeDia = GetExternalPipeDiameter(my_IJHgrInputConfigHlpr, Index)
        If dblPipeDia > dblMaxDia Then
            dblMaxDia = dblPipeDia
        End If
    Next

    GetMaxExternalPipeDiameter = dblMaxDia
    Set my_IJHgrInputConfigHlpr = Nothing
    Set IJElements_PipeCollection = Nothing
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetMaxExternalPipeIndex
'
' Desc:    This method retrieved the index of the pipe with max diameter among multi pipes.

' Param:
'           Output: route index
'
'*************************************************************************************

Public Function GetMaxExternalPipeIndex(ByVal pDispInputConfigHlpr As Object) As Integer
    
    Const METHOD = "GetExternalPipeDiameter:"
    On Error GoTo ErrorHandler
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
      
    Dim dblPipeDia                  As Double
    Dim dblMaxDia                   As Double
    Dim Index                       As Integer
    Dim maxPipeIndex                As Integer
    Dim numPipes                    As Integer
    Dim oPipeCollection             As Object
    Dim IJElements_PipeCollection   As IJElements
    
    my_IJHgrInputConfigHlpr.GetPipes oPipeCollection
    Set IJElements_PipeCollection = oPipeCollection
    numPipes = IJElements_PipeCollection.Count
    
    dblMaxDia = 0#
    maxPipeIndex = 1
    For Index = 1 To numPipes
        dblPipeDia = GetExternalPipeDiameter(my_IJHgrInputConfigHlpr, Index)
        If dblPipeDia > dblMaxDia Then
            dblMaxDia = dblPipeDia
            maxPipeIndex = Index
        End If
    Next

    GetMaxExternalPipeIndex = maxPipeIndex
    Set my_IJHgrInputConfigHlpr = Nothing
    Set IJElements_PipeCollection = Nothing
    Set oPipeCollection = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetExternalPipeDiameter
'
' Desc:    This method retrieved external diameter of pipe.


' Param:
'           Output: diameter of pipe
'
'*************************************************************************************
Public Function GetExternalPipeDiameter(ByVal pDispInputConfigHlpr As Object, ByVal RouteIndex As Integer) As Double
    
    Const METHOD = "GetExternalPipeDiameter:"
    On Error GoTo ErrorHandler
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Dim my_IJHgrInputObjectInfo As IJHgrInputObjectInfo
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    Set my_IJHgrInputObjectInfo = pDispInputConfigHlpr
      
    Dim lInsulPurpose           As Long
    Dim lInsulMat               As Long
    Dim dInsulThk               As Double
    Dim bInsulFlag              As Boolean
    Dim dblPipeDia              As Double
    Dim oPipes                  As Object
    Dim pPipeColl               As IJElements
    
    my_IJHgrInputObjectInfo.GetInsulationData RouteIndex, lInsulPurpose, lInsulMat, dInsulThk, bInsulFlag
    my_IJHgrInputConfigHlpr.GetPipes oPipes
    Set pPipeColl = oPipes
    
    dblPipeDia = my_IJHgrInputConfigHlpr.GetExternalPipeDiameter(pPipeColl.Item(RouteIndex))
    If bInsulFlag <> False Then
        dblPipeDia = dblPipeDia + 2# * dInsulThk
    End If

    GetExternalPipeDiameter = dblPipeDia
    
    Set my_IJHgrInputConfigHlpr = Nothing
    Set my_IJHgrInputObjectInfo = Nothing
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetBoundaryObjectDimension
'
' Desc:    This method retrieved the dimension of the boundary objects on the defined
'          bounding box.

' Param:
'          Output: dimension array (Left/Right/Top/Bottom)
'*************************************************************************************

Public Function GetBoundaryObjectDimension(ByVal pDispInputConfigHlpr As Object, _
                                           ByVal BBXName As String, _
                                           ByVal BBXRefPlane As HgrBBXRefPlane) As Double()
    
    Const METHOD = "GetBoundaryObjectDimension:"
    On Error GoTo ErrorHandler
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    
    Dim my_IJHgrBBXHlpr As IJHgrBBXHlpr
    Set my_IJHgrBBXHlpr = pDispInputConfigHlpr
    
    Dim pStdBBX As Object
    Dim StdBBX As IJHgrBBX
    Dim BBXBoundary(1 To 4) As HgrBBXBoundary
    Dim BBXOffset(1 To 4) As Double
    
    BBXBoundary(1) = LEFT
    BBXBoundary(2) = RIGHT
    BBXBoundary(3) = Bottom
    BBXBoundary(4) = Top
    
    'get standard bounding box definition
    my_IJHgrBBXHlpr.GetBBX BBXName, pStdBBX '"StdBBX", pStdBBX
    
    Set StdBBX = pStdBBX
    
    Dim idx As Integer
    For idx = 1 To 4
        Dim pRteObjColl As Object
        Dim pRteElements As IJElements
        Dim pRteObj As Object
        
        'get boundary route objects on the BBR bounding box
        StdBBX.GetBoundaryRteObj BBXRefPlane, BBXBoundary(idx), pRteObjColl
        Set pRteElements = pRteObjColl
        Set pRteObj = pRteElements.Item(1)
        
        'get first route object dimension
        'BBXOffset(idx) = my_IJHgrInputConfigHlpr.GetPipeDiameter(pRteObj)
        Dim eShape As CrossSectionShapeTypes
        Dim pRadius As Double
        Dim dWidth As Double
        Dim dDepth As Double
        Dim strUnit As String
        my_IJHgrInputConfigHlpr.GetRteCrossSectParams pRteObj, eShape, strUnit, pRadius, dWidth, dDepth
       
        If Not strUnit = "" Then
            Dim oUOMService         As IJUomVBInterface
            Dim unitID              As Units
            Set oUOMService = New UnitsOfMeasureServicesLib.UomVBInterface
            unitID = oUOMService.GetUnitId(UNIT_DISTANCE, strUnit)
            pRadius = oUOMService.ConvertUnitToDbu(UNIT_DISTANCE, pRadius, unitID)
        End If
        
        If pRadius = 0# Then
            BBXOffset(idx) = dWidth
        Else
            BBXOffset(idx) = 2# * pRadius
        End If
        
        Set pRteObjColl = Nothing
        Set pRteObj = Nothing
    Next
    
    GetBoundaryObjectDimension = BBXOffset
    Set my_IJHgrBBXHlpr = Nothing
    Set my_IJHgrInputConfigHlpr = Nothing
    Set pStdBBX = Nothing
    Set StdBBX = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetIsLugEndOffsetApplied
'
' Desc:    This method checks whether a offset value needs to
'          be explicitly specified on both end of the leg
'          If the input structure is either a slab or a non-parallel member,
'          then the offset needs to be specified. Else no offset value needs to
'          be specified
' Param:
'          Output:    boolean array ( 1: LeftOffsetAppliedFlag
'                                     2: RightOffsetAppliedFlag)
'*************************************************************************************

Public Function GetIsLugEndOffsetApplied(ByVal pDispInputConfigHlpr As Object) As Boolean()
    
    Const METHOD = "GetLugEndOffset:"
    On Error GoTo ErrorHandler
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Dim RteObjects As Object
    Dim RteObj As Object
    Dim StructObjects As Object
    
    'first route object is set as primary route object
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    
    my_IJHgrInputConfigHlpr.GetSupportedCollection RteObjects
    Set RteObj = RteObjects.Item(1)
    
    my_IJHgrInputConfigHlpr.GetStructure StructObjects
    
    Dim varIsOffsetApplied() As Variant
    Dim bIsOffsetApplied(1 To 2) As Boolean
    
    bIsOffsetApplied(1) = True
    bIsOffsetApplied(2) = True
    
    Dim idx As Integer
    If Not StructObjects Is Nothing Then
        If StructObjects.Count >= 2 Then
            For idx = 1 To 2 'StructObjects.Count
        
                Dim inputObjColl As New Collection
        
                inputObjColl.Add pDispInputConfigHlpr
        
                'put primary route object into collection
                inputObjColl.Add RteObj
        
                'put structure into collection
                inputObjColl.Add StructObjects.Item(idx)
        
                'check if offset is to be applied
                varIsOffsetApplied = my_IJHgrInputConfigHlpr.GetDataByRule("HgrSupIsOffsetApplied", inputObjColl)
                bIsOffsetApplied(idx) = varIsOffsetApplied(1)
                
                Set inputObjColl = Nothing
            Next
        End If
    End If
    GetIsLugEndOffsetApplied = bIsOffsetApplied
    
    Set my_IJHgrInputConfigHlpr = Nothing
    Set RteObj = Nothing
    Set RteObjects = Nothing
    Set StructObjects = Nothing
     
Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetBoundingBoxDimension
'
' Desc:    This method returns the dimension of bounding box itself
' Param:
'         Output:    double array
'                   (1: Width   2: Depth)
'*************************************************************************************

Public Function GetBoundingBoxDimension(ByVal pDispInputConfigHlpr As Object) As Double()
    
    Const METHOD = "GetBoundingBoxDimension:"
    On Error GoTo ErrorHandler
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    
    Dim strBBName As String
    If my_IJHgrInputConfigHlpr.IsPlaceByStructure Then
        strBBName = "BBSR"
    Else
        strBBName = "BBR"
    End If

    Dim oBBInfo As Object
    Set oBBInfo = my_IJHgrInputConfigHlpr.GetBoundingBoxInfo(strBBName)
    
    Dim IJElements_BBInfo As IJElements
    Set IJElements_BBInfo = oBBInfo

    Dim oBBHigh As Object
    Set oBBHigh = IJElements_BBInfo.Item(1)

    Dim BB_Position As IJDPosition
    Set BB_Position = oBBHigh

    Dim dimension(1 To 2) As Double
    dimension(1) = BB_Position.x
    dimension(2) = BB_Position.y
    
    GetBoundingBoxDimension = dimension
    
    Set my_IJHgrInputConfigHlpr = Nothing
    Set BB_Position = Nothing
    Set oBBHigh = Nothing
    Set oBBInfo = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetPipeStructuralASMAttributes
'
' Desc:   This method get/set all the occurrence attributes defined on the
'         pipe structural support assemblies
' Param:
'        Input :pDispInputConfigHlpr(object)
'        OutPut:Variant
'*************************************************************************************

Public Function GetPipeStructuralASMAttributes(ByVal pDispInputConfigHlpr As Object) As Variant()
    Const METHOD = "GetPipeStructuralASMAttributes:"
    On Error GoTo ErrorHandler
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    
    Dim AttributeColl(1 To 4) As Variant
    
    'retrieve the support data (attributes)
    Dim D As Double
    Dim G As Double
   
    Dim leftPad As Boolean
    Dim rightPad As Boolean
    
    Dim varD As Variant
    Dim varG As Variant
    Dim varPCS As Variant
    Dim SupAttrib() As Variant
    Dim Pads As Variant
    
    Dim hgrStatus As Long
    my_IJHgrInputConfigHlpr.GetHgrSupStatus hgrStatus

    If (hgrStatus And SUPPORT_RECOMPUTE) > 0 Then
        '=============================
        '(RECOMPUTE PROCESS)retrieve the occurrence value
        '=============================
        'retrieve D and G
        my_IJHgrInputConfigHlpr.GetAttributeValue "D", Nothing, varD
        my_IJHgrInputConfigHlpr.GetAttributeValue "Gap", Nothing, varG
        D = varD
        G = varG

        'retrieve pad configuration
        my_IJHgrInputConfigHlpr.GetAttributeValue "LeftPad", Nothing, Pads
        leftPad = Pads
        my_IJHgrInputConfigHlpr.GetAttributeValue "RightPad", Nothing, Pads
        rightPad = Pads
    Else
        '=====================================
        '(INITIAL PLACEMENT)recalculate all the support attributes
        'based on predefined rule
        '=====================================
        'Set D and G
        SupAttrib = my_IJHgrInputConfigHlpr.GetDataByRule("HgrSupAngleByLF")

        D = SupAttrib(4)
        my_IJHgrInputConfigHlpr.SetAttributeValue "D", Nothing, D
        SupAttrib = my_IJHgrInputConfigHlpr.GetDataByRule("HgrSup_G")
        G = SupAttrib(1)
        my_IJHgrInputConfigHlpr.SetAttributeValue "Gap", Nothing, G

        my_IJHgrInputConfigHlpr.GetAttributeValue "LeftPad", Nothing, varPCS
        leftPad = varPCS
        my_IJHgrInputConfigHlpr.GetAttributeValue "RightPad", Nothing, varPCS
        rightPad = varPCS
    End If
    
    AttributeColl(1) = D
    AttributeColl(2) = G
    AttributeColl(3) = leftPad
    AttributeColl(4) = rightPad
    
    GetPipeStructuralASMAttributes = AttributeColl
    Set my_IJHgrInputConfigHlpr = Nothing
    Exit Function
    
ErrorHandler:
     Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetDuctStructuralASMAttributes
'
' Desc:   'This method get/set all the occurrence attributes defined on the
'          duct structural support assemblies
' Param:
'          Input :pDispInputConfigHlpr (Object)
'          Output:Variant
'*************************************************************************************

Public Function GetDuctStructuralASMAttributes(ByVal pDispInputConfigHlpr As Object) As Variant()
    Const METHOD = "GetDuctStructuralASMAttributes:"
    On Error GoTo ErrorHandler
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    
    Dim AttributeColl(1 To 4) As Variant
    
    'retrieve the support data (attributes)
    Dim D As Double
    Dim G As Double
    Dim leftPad As Boolean
    Dim rightPad As Boolean
    
    Dim varD As Variant
    Dim varG As Variant
    Dim varLPad As Variant
    Dim varRPad As Variant
    Dim SupAttrib() As Variant
    
    Dim hgrStatus As Long
    my_IJHgrInputConfigHlpr.GetHgrSupStatus hgrStatus
    
    If (hgrStatus And SUPPORT_RECOMPUTE) > 0 Then
        '=============================
        '(RECOMPUTE PROCESS)retrieve the occurrence value
        '=============================
        'retrieve D and G
        my_IJHgrInputConfigHlpr.GetAttributeValue "D", Nothing, varD
        my_IJHgrInputConfigHlpr.GetAttributeValue "Gap", Nothing, varG
        D = varD
        G = varG
        
        'retrieve pad configuration
        my_IJHgrInputConfigHlpr.GetAttributeValue "LeftPad", Nothing, varLPad
        my_IJHgrInputConfigHlpr.GetAttributeValue "RightPad", Nothing, varRPad
        leftPad = varLPad
        rightPad = varRPad
    Else
        '=====================================
        '(INITIAL PLACEMENT)recalculate all the support attributes
        'based on predefined rule
        '=====================================
        
        ' Retrieve duct shape
        Dim inputObjects As Object
        Dim pRteObj As Object
        Dim eShape As CrossSectionShapeTypes
        Dim pRadius As Double
        Dim dWidth As Double
        Dim dDepth As Double
        Dim strUnit As String
        Dim strRule As String
        
        my_IJHgrInputConfigHlpr.GetDucts inputObjects
        Set pRteObj = inputObjects.Item(1)
        my_IJHgrInputConfigHlpr.GetRteCrossSectParams pRteObj, eShape, strUnit, pRadius, dWidth, dDepth

        'Set D and G
        If eShape = Rectangular Then
            SupAttrib = my_IJHgrInputConfigHlpr.GetDataByRule("HgrSupDuctAngleOffset")
            D = SupAttrib(1)
            G = SupAttrib(2)
        Else
            SupAttrib = my_IJHgrInputConfigHlpr.GetDataByRule("HgrSupDuctClamp_D")
            D = SupAttrib(1)
            SupAttrib = my_IJHgrInputConfigHlpr.GetDataByRule("HgrSup_G")
            G = SupAttrib(1)
        End If
        my_IJHgrInputConfigHlpr.SetAttributeValue "D", Nothing, D
        my_IJHgrInputConfigHlpr.SetAttributeValue "Gap", Nothing, G
        
        my_IJHgrInputConfigHlpr.GetAttributeValue "LeftPad", Nothing, varLPad
        my_IJHgrInputConfigHlpr.GetAttributeValue "RightPad", Nothing, varRPad
        
        leftPad = varLPad
        rightPad = varRPad
        
        my_IJHgrInputConfigHlpr.SetAttributeValue "LeftPad", Nothing, leftPad
        my_IJHgrInputConfigHlpr.SetAttributeValue "RightPad", Nothing, rightPad
    End If
    
    AttributeColl(1) = D
    AttributeColl(2) = G
    AttributeColl(3) = leftPad
    AttributeColl(4) = rightPad
    
    GetDuctStructuralASMAttributes = AttributeColl
    
    Set my_IJHgrInputConfigHlpr = Nothing
    Set pRteObj = Nothing
    
 Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetAngleCSDimension
'
' Desc:   'This method is used to get cross section dimension of HgrBeam

' Param:  Input:pAngleStructOcc as Object
'         Output:Double
'
'
'*************************************************************************************

Public Function GetAngleCSDimension(ByVal pAngleStructOcc As Object) As Double()
    Const METHOD = "GetAngleCSDimension:"
    On Error GoTo ErrorHandler
    
    Dim oPartOcc As PartOcc
    Dim oHgrPart As IJDPart
    Dim oHgrPortHlpr As HgrSymbolPortHlpr
    Dim oCrossSectionColl As IJDTargetObjectCol
    Dim oCrossSectionServices As CrossSectionServices
    Dim oCSObj As Object
    Dim varValue As Variant
    Dim Width As Double, Depth As Double  ' primary cross section width and depth
    Dim dimension() As Double
    
    Set oPartOcc = pAngleStructOcc
    oPartOcc.GetPart oHgrPart
    Set oHgrPortHlpr = New HgrSymbolPortHlpr
    Call oHgrPortHlpr.GetHgrAssociation(oHgrPart, HGR_CROSSSECREL, oCrossSectionColl)
    Set oCSObj = oCrossSectionColl.Item(1)
    Set oCrossSectionServices = New CrossSectionServices
    ReDim dimension(1 To 5) As Double
    oCrossSectionServices.GetCrossSectionAttributeValue oCSObj, "Width", varValue
    dimension(1) = varValue
    oCrossSectionServices.GetCrossSectionAttributeValue oCSObj, "Depth", varValue
    dimension(2) = varValue
    oCrossSectionServices.GetCrossSectionAttributeValue oCSObj, "tf", varValue
    dimension(3) = varValue
    oCrossSectionServices.GetCrossSectionAttributeValue oCSObj, "tw", varValue
    dimension(4) = varValue
    On Error Resume Next
    oCrossSectionServices.GetCrossSectionAttributeValue oCSObj, "bb", varValue
    dimension(5) = varValue
    On Error GoTo ErrorHandler
    
    GetAngleCSDimension = dimension
    
    Set oPartOcc = Nothing
    Set oHgrPart = Nothing
    Set oHgrPortHlpr = Nothing
    Set oCrossSectionColl = Nothing
    Set oCSObj = Nothing
    
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetIndexedStructPortName
'
' Desc:   This method differentiate multiple structure input object based on
'         relative position. For U shape support, when two structure are inputted.
'         This method returns the port name on left and right side of the route
'         object. (Left is on the negative Y axis of bounding box coord. sys., right
'         is on the positive Y axis of the bounding box coord. sys.)

' Param:
'       Outputs:   string array
'           1: Left structure reference port name
'           2: right structure reference port name
'
'*************************************************************************************

Public Function GetIndexedStructPortName(ByVal pDispInputConfigHlpr As Object, _
                                         ByRef varIsOffsetApplied() As Boolean) As String()
    
    Const METHOD = "GetIndexedStructPortName:"
    On Error GoTo ErrorHandler
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    
    Dim RefColl As New Collection
    Dim StructPort(1 To 2) As String
    Dim varIdxStructPort As Variant
    
    RefColl.Add pDispInputConfigHlpr
    RefColl.Add varIsOffsetApplied
    
    Dim eHgrCmdType As HgrCmdType
    Dim oIJHgrSupport As IJHgrSupport
    Set oIJHgrSupport = my_IJHgrInputConfigHlpr
       
    eHgrCmdType = oIJHgrSupport.CommandType
    
    Set oIJHgrSupport = Nothing
    
    If eHgrCmdType = HgrByReferencePointCmdType Then
        StructPort(1) = "Structure"
        StructPort(2) = "Structure"
    Else
        varIdxStructPort = my_IJHgrInputConfigHlpr.GetDataByRule("HgrSupIdxStructPortName", RefColl)
        
        StructPort(1) = varIdxStructPort(1)
        StructPort(2) = varIdxStructPort(2)
    End If
    'switch the OffsetApplied flag
    If StructPort(1) = "Struct_2" Then
        Dim TmpFlag As Boolean
        
        TmpFlag = varIsOffsetApplied(1)
        varIsOffsetApplied(1) = varIsOffsetApplied(2)
        varIsOffsetApplied(2) = TmpFlag
    End If
    
    GetIndexedStructPortName = StructPort
    Set my_IJHgrInputConfigHlpr = Nothing
    Set RefColl = Nothing
    
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetLSupportConfiguration
'
' Desc:   This method retrieve the configuration of L shape strutural support
'         There are two configuration of the L shape support
'          config 1: the top supporting member is on the right of the route object
'          config 2: the top supporting member is on the left of the route object
'         This method also returns whether an offset is needed for a certain end and als
'         determines the offset value
'          offset(1) defines the offset between BBX and vertical leg
'          offset(2) defines the offset between BBX and the horizontal leg
' Parameter:
'            Input :pDispInputConfigHlpr (as Object)
'                   bIsOffsetApplied( as Boolean)
'                   LugoffSet(as Double)
'            output as long
'
'*************************************************************************************

Public Function GetLSupportConfiguration(ByVal pDispInputConfigHlpr As Object, _
                                         ByRef bIsOffsetApplied() As Boolean, _
                                         ByRef LugOffset() As Double) As Long
    Const METHOD = "GetLSupportConfig:"
    On Error GoTo ErrorHandler

    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    
    Dim BBXOffset() As Double
    Dim TopStructPort As String
    Dim BottomStructPort As String
    Dim RouteRefPort As String
    Dim StructObjects As Object
    Dim RteStructConfig As Long
    Dim angle As Double
    
    Dim getOffsetApplied() As Boolean
    
    'Get offset value from rule
    Dim refPlane As HgrBBXRefPlane
    If my_IJHgrInputConfigHlpr.IsPlaceByStructure Then
        refPlane = BBSR
    Else
        refPlane = BBR
    End If
    BBXOffset = GetBoundaryObjectDimension(pDispInputConfigHlpr, "StdBBX", refPlane)
    
    my_IJHgrInputConfigHlpr.GetStructure StructObjects
    
    'default value for one structure input
    bIsOffsetApplied(1) = True
    bIsOffsetApplied(2) = True
    
    TopStructPort = "Structure"
    BottomStructPort = "Struct_2"
    
    RteStructConfig = -1
    
    'if first input is a slab, offset is needed
    getOffsetApplied = GetIsLugEndOffsetApplied(pDispInputConfigHlpr)
    bIsOffsetApplied(1) = getOffsetApplied(1)
    bIsOffsetApplied(2) = getOffsetApplied(2)
    
    If my_IJHgrInputConfigHlpr.IsPlaceByStructure Then
        RouteRefPort = "BBSR_Low"
    Else
        RouteRefPort = "BBR_Low" '"Route"
    End If
    
    If Not bIsOffsetApplied(1) Then
        'check whether it is on the left  or right
        angle = my_IJHgrInputConfigHlpr.GetRouteStructConfigAngle(RouteRefPort, TopStructPort, HGRPORT_Y)
        
        If Abs(angle) < (4 * Atn(1)) / 2# Then
            'the member is on the right side of the route
            RteStructConfig = 1
        Else
            RteStructConfig = 2
        End If
    End If
    
    'check the second input
    If StructObjects.Count = 2 Then
        bIsOffsetApplied(2) = False
        
        If RteStructConfig = -1 Then
            'check whether it is on the left  or right
            angle = my_IJHgrInputConfigHlpr.GetRouteStructConfigAngle(RouteRefPort, BottomStructPort, HGRPORT_Y)
            
            If Abs(angle) < (4 * Atn(1)) / 2# Then
                'the top member is on the left side of the route
                RteStructConfig = 2
            Else
                RteStructConfig = 1
            End If
        End If
    End If
    
    'define appropriate offset value
    Dim varD As Variant
    Dim D As Double
    
    my_IJHgrInputConfigHlpr.GetAttributeValue "D", Nothing, varD
    D = varD
    
    Dim AngleConfig() As Variant
    AngleConfig = my_IJHgrInputConfigHlpr.GetDataByRule("HgrSupAngleByLF")
    LugOffset(1) = D / 2# - BBXOffset(1) / 2#
    LugOffset(2) = AngleConfig(4) - BBXOffset(2) / 2#
    
    If RteStructConfig = 1 Then
        LugOffset(1) = BBXOffset(2)
        LugOffset(2) = BBXOffset(1)
    End If

    GetLSupportConfiguration = RteStructConfig
    Set my_IJHgrInputConfigHlpr = Nothing
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetLSupportConfiguration
'
' Desc:   This method retrieve the configuration of L shape strutural support for DUCTS
'   There are two configuration of the L shape support
'   config 1: the top supporting member is on the right of the route object
'   config 2: the top supporting member is on the left of the route object
'   This method also returns whether an offset is needed for a certain end
'Parameters:
'              Input -pDispInputConfigHlpr(Object)
'                    -bIsOffsetApplied(Boolean)
'              Output-Long
'*************************************************************************************
Public Function GetDuctLSupportConfiguration(ByVal pDispInputConfigHlpr As Object, _
                                             ByRef bIsOffsetApplied() As Boolean) As Long
    Const METHOD = "GetLSupportConfig:"
    On Error GoTo ErrorHandler

    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    
    Dim BBXOffset() As Double
    Dim TopStructPort As String
    Dim BottomStructPort As String
    Dim RouteRefPort As String
    Dim StructObjects As Object
    Dim RteStructConfig As Long
    Dim angle As Double
    
    Dim getOffsetApplied() As Boolean

    'Get offset value from rule
    Dim refPlane As HgrBBXRefPlane
    If my_IJHgrInputConfigHlpr.IsPlaceByStructure Then
        refPlane = BBSR
    Else
        refPlane = BBR
    End If
    BBXOffset = GetBoundaryObjectDimension(pDispInputConfigHlpr, "StdBBX", refPlane)
    
    my_IJHgrInputConfigHlpr.GetStructure StructObjects
    
    'default value for one structure input
    bIsOffsetApplied(1) = True
    bIsOffsetApplied(2) = True
    
    TopStructPort = "Structure"
    BottomStructPort = "Struct_2"
    
    RteStructConfig = -1#
    
    
    'if first input is a slab, offset is needed
    getOffsetApplied = GetIsLugEndOffsetApplied(pDispInputConfigHlpr)
    bIsOffsetApplied(1) = getOffsetApplied(1)
    bIsOffsetApplied(2) = getOffsetApplied(2)
    
    If my_IJHgrInputConfigHlpr.IsPlaceByStructure Then
        RouteRefPort = "BBSR_Low"
    Else
        RouteRefPort = "BBR_Low" '"Route"
    End If
    
    If Not bIsOffsetApplied(1) Then
        'check whether it is on the left  or right
        angle = my_IJHgrInputConfigHlpr.GetRouteStructConfigAngle(RouteRefPort, TopStructPort, HGRPORT_Y)
        
        If Abs(angle) < (4 * Atn(1)) / 2# Then
            'the member is on the right side of the route
            RteStructConfig = 1
        Else
            RteStructConfig = 2
        End If
    End If
    
    'check the second input
    If StructObjects.Count = 2 Then
        bIsOffsetApplied(2) = False
        
        If RteStructConfig = -1 Then
            'check whether it is on the left  or right
            angle = my_IJHgrInputConfigHlpr.GetRouteStructConfigAngle(RouteRefPort, BottomStructPort, HGRPORT_Y)
            
            If Abs(angle) < (4 * Atn(1)) / 2# Then
                'the top member is on the left side of the route
                RteStructConfig = 2
            Else
                RteStructConfig = 1
            End If
        End If
    End If
    
    'define appropriate offset value
    Dim varD As Variant
    Dim D As Double
    
    my_IJHgrInputConfigHlpr.GetAttributeValue "D", Nothing, varD
    D = varD

    GetDuctLSupportConfiguration = RteStructConfig
    Set my_IJHgrInputConfigHlpr = Nothing
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   CrossSectionAttributeValue
'
' Desc:
' Parameters:
'              Input -oSupportComp(PartOcc)
'
'              Output-Double
'*************************************************************************************

Public Function CrossSectionAttributeValue(ByVal oSupportComp As PartOcc) As Double()
Const METHOD = "CrossSectionAttributeValue:"
On Error GoTo ErrorHandler

    Dim oHgrPart As IJDPart
    Dim dblWidth As Double
    Dim dblDepth As Double
    Dim dblThick As Double
    Dim varWidth As Variant
    Dim varDepth As Variant
    Dim varThick As Variant
    Dim oHgrPortHlpr As HgrSymbolPortHlpr
    Dim oCrossSectionColl As IJDTargetObjectCol
    Dim oCrossSectionServices As CrossSectionServices
    Dim oCSObj As Object
    Dim oCSOcc As Object
    
    oSupportComp.GetPart oHgrPart
    
    Set oHgrPortHlpr = New HgrSymbolPortHlpr
    Call oHgrPortHlpr.GetHgrAssociation(oHgrPart, 3, oCrossSectionColl)
      
    Set oCSObj = oCrossSectionColl.Item(1)
    Set oCrossSectionServices = New CrossSectionServices

    oCrossSectionServices.GetCrossSectionAttributeValue oCSObj, "Width", varWidth
    dblWidth = varWidth
    oCrossSectionServices.GetCrossSectionAttributeValue oCSObj, "Depth", varDepth
    dblDepth = varDepth
    oCrossSectionServices.GetCrossSectionAttributeValue oCSObj, "tnom", varThick
    dblThick = varThick

    Dim CSAttributes(1 To 3) As Double
    CSAttributes(1) = dblWidth
    CSAttributes(2) = dblDepth
    CSAttributes(3) = dblThick
    CrossSectionAttributeValue = CSAttributes
    
    Set oCrossSectionColl = Nothing
    Set oCSObj = Nothing
    Set oCrossSectionServices = Nothing
    Set oHgrPortHlpr = Nothing
    Set oHgrPart = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function


