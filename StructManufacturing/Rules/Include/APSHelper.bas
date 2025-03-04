Attribute VB_Name = "APSHelper"
Option Explicit

'----------------------------------------------------------------------------------------------------------------------------------------------------
' Copyright (C) 2009 Intergraph Corporation. All rights reserved.
'
' File Info:
'     Module:  APSHelper.bas
'
' Abstract:
'      Helper for the Mfg Marking rules
'
' History:
'         April 24. 2009      Created.
'
'----------------------------------------------------------------------------------------------------------------------------------------------------

' Enumerator for the mode by which the user defined marking is created in the marking line command.
Public Enum enumUserMarkingLineMode
    eIntersectionMode = 0   ' Intersection
    eSketchMode = 1         ' 2D Projection
    eReferenceMode = 2      ' Reference Curve
End Enum

' ***********************************************************************************
' Private Function GetRelatedPartThickness()
'
' Description:  Function gets the appropriate profile side from the profile marking side
'               Input arguments: Marking Line Active Entity.
'               Output argument: Profile Side as Long
'
' ***********************************************************************************
Public Function GetRelatedPartThickness(oMarkingLineAE As IJMfgMarkingLines_AE) As Double
    Const METHOD = "GetRelatedPartThickness"
    On Error GoTo ErrorHandler

    Dim oRelatedObject As Object
    Set oRelatedObject = oMarkingLineAE.GetMfgMarkingRelatedObject
    
    If Not oRelatedObject Is Nothing Then
        If TypeOf oRelatedObject Is IJPlate Then
            Dim oSDConPlateWrapper As New StructDetailObjects.PlatePart
            Set oSDConPlateWrapper.object = oRelatedObject
            GetRelatedPartThickness = oSDConPlateWrapper.PlateThickness
        ElseIf TypeOf oRelatedObject Is IJStiffener Then
            Dim oSDConProfileWrapper As New StructDetailObjects.ProfilePart
            Set oSDConProfileWrapper.object = oRelatedObject
            GetRelatedPartThickness = oSDConProfileWrapper.WebThickness
        ElseIf TypeOf oRelatedObject Is IJBeam Then
            Dim oSDConBeamWrapper As New StructDetailObjects.BeamPart
            Set oSDConBeamWrapper.object = oRelatedObject
            GetRelatedPartThickness = oSDConBeamWrapper.WebThickness
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************************************
' Private Function GetProfileMarkingSide()
'
' Description:  Function gets the appropriate profile side from the profile marking side
'               Input arguments: Marking Line Active Entity.
'               Output argument: Profile Side as Long
'
' ***********************************************************************************
Public Function GetProfileMarkingSide(oMarkingLineAE As IJMfgMarkingLines_AE) As Long
    Const METHOD = "GetProfileMarkingSide"
    On Error GoTo ErrorHandler

    Dim oMarkingPart As Object
    Set oMarkingPart = oMarkingLineAE.GetMfgMarkingPart
    
    Dim oMarkingPort As IJPort
    Set oMarkingPort = oMarkingLineAE.GetMfgMarkingPortForProfilePart(oMarkingPart)
    
    Dim oStructPort As IJStructPort
    Set oStructPort = oMarkingPort
    
    Dim oProfileWrapper       As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oMarkingPart
    
    GetProfileMarkingSide = oProfileWrapper.GetSide(oStructPort)

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************************************
' Private Function GetMarkingLineAE()
'
' Description:  Function returns the MarkingLine Active Entity from the Geom3d object
'               Input arguments: Geom3d object
'               Output argument: Marking Line Active Entity.
'
' ***********************************************************************************
Public Function GetMarkingLineAE(oGeom3d As IJMfgGeom3D) As IJMfgMarkingLines_AE
    Const METHOD = "GetMarkingLineAE"
    On Error GoTo ErrorHandler
    
    Dim oMoniker            As IMoniker
    Dim oUnkMLData          As Object
    Dim oMarkingLineData    As IJMfgMarkingLinesData
    Dim oMfgRuleHelper      As MfgRuleHelpers.Helper
    
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper

    Set oMoniker = oGeom3d.GetMoniker
    Set oUnkMLData = oMfgRuleHelper.BindToObject(GetActiveConnection.GetResourceManager(GetActiveConnectionName), oMoniker)
    Set oMarkingLineData = oUnkMLData
    
    Set GetMarkingLineAE = oMarkingLineData.GetMfgMarkingLines_AE

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'**********************************************************
'* Determine the face port from a given edge port *
'**********************************************************
Public Sub Get_FacePortFromEdgePort(oEdgePort As IJStructPort, _
                                    oFacePort As IJStructPort)

Const sMETHOD = "Get_FacePortFromEdgePort"
On Error GoTo ErrorHandler
Dim sText As String

    Dim iIndex As Long
    Dim nPorts As Long
    
    Dim lXId As Long
    Dim lCtxId As Long
    Dim lOptId As Long
    Dim lOprId As Long
    
    Dim oPort As IJPort
    Dim oStructPort As IJStructPort
    Dim oConnectable As IJConnectable
    Dim oStructConnectable As IJStructConnectable
    Dim oFacePortList As IJElements
    
    Set oPort = oEdgePort
    Set oConnectable = oPort.Connectable
    Set oStructConnectable = oConnectable
    
    Dim oTopologyLocate As GSCADStructGeomUtilities.IJTopologyLocate
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate

    lCtxId = oEdgePort.ContextID
    lOptId = oEdgePort.OperationID
    lOprId = oEdgePort.OperatorID
    lXId = oEdgePort.SectionID
    
    oStructConnectable.enumConnectablePortsByOperationAndTopology _
                                oFacePortList, _
                                vbNullString, _
                                JS_TOPOLOGY_FILTER_ALL_LFACES, _
                                True

    nPorts = oFacePortList.Count

    For iIndex = 1 To nPorts
        If TypeOf oFacePortList.Item(iIndex) Is IJStructPort Then
            Set oStructPort = oFacePortList.Item(iIndex)
            Set oFacePort = oStructPort
            Exit For
        End If
    Next iIndex
  
Exit Sub

ErrorHandler:
  Err.Raise Err.Number, Err.Source, Err.Description
  
End Sub

' ***********************************************************************************
' Private Function GetDirVector
'
' Description:  Returns a vector in the specified direction.
'               Input arguments: direction, marking side.
'               Output argument: direction vector.
'
' ***********************************************************************************
Public Function GetDirVector(strDir As String, side As Long) As IJDVector
    Const METHOD = "GetDirVector"
    On Error GoTo ErrHandler
    
    Dim oVectorCustom As IJDVector
    Set oVectorCustom = New DVector
    
    ' Side: This input in only needed when strDir is In or Out
    ' -1 means marking is on Starboard side
    '  1 means marking is on Port side.
    
    Select Case strDir
    Case "aft":
        oVectorCustom.x = -1
        oVectorCustom.y = 0
        oVectorCustom.z = 0
    Case "fore":
        oVectorCustom.x = 1
        oVectorCustom.y = 0
        oVectorCustom.z = 0
    Case "in":
        oVectorCustom.x = 0
        oVectorCustom.y = -side
        oVectorCustom.z = 0
    Case "out":
        oVectorCustom.x = 0
        oVectorCustom.y = side
        oVectorCustom.z = 0
    Case "starboard":
        oVectorCustom.x = 0
        oVectorCustom.y = -1
        oVectorCustom.z = 0
    Case "port":
        oVectorCustom.x = 0
        oVectorCustom.y = 1
        oVectorCustom.z = 0
    Case "upper":
        oVectorCustom.x = 0
        oVectorCustom.y = 0
        oVectorCustom.z = 1
    Case "lower":
        oVectorCustom.x = 0
        oVectorCustom.y = 0
        oVectorCustom.z = -1
    End Select
    
    Set GetDirVector = oVectorCustom
Exit Function
ErrHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************************************
' Private Function GetUserDefinedMarkingLineMode()
'
' Description:  Function returns the Mode by which the Marking Line was created in Marking Line command
'               Input arguments: Marking Line Active Entity
'               Output argument: Mode by which the Marking Line was created in Marking Line command
'
' ***********************************************************************************
Public Function GetUserDefinedMarkingLineMode(oMarkingLineAE As IJMfgMarkingLines_AE) As enumUserMarkingLineMode
    Const METHOD = "GetUserDefinedMarkingLineMode"
    On Error GoTo ErrorHandler

    Dim oRefInput As Object
    Set oRefInput = oMarkingLineAE.GetMfgMarkingRefInput
    
    If oMarkingLineAE.Symbol Is Nothing Then
        If (TypeOf oRefInput Is IJPlane) Or (TypeOf oRefInput Is ISPGRadialCylinder) Then
            GetUserDefinedMarkingLineMode = eIntersectionMode
        Else
            GetUserDefinedMarkingLineMode = eReferenceMode
        End If
    Else
        GetUserDefinedMarkingLineMode = eSketchMode
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
' ***********************************************************************************
' Private Function GetMarkingSideAndPort()
'
' Description:  Function gets the appropriate plate marking side and port from the Geom3d object
'               Input arguments: Marking Line Active Entity.
'               Output argument: Plate Side as Long
'
' ***********************************************************************************
Public Sub GetMarkingSideAndPort(oGeom3d As IJMfgGeom3D, ByRef lSide As Long, Optional ByRef oPort As IJPort)
    Const METHOD = "GetMarkingSideAndPort"
    On Error GoTo ErrorHandler

    Dim oMarkingLineAE As IJMfgMarkingLines_AE
    Dim oMarkingPart As Object
   
    'Get the Marking Line AE from Geom3d object
    Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
   
    'Get the Part on which the marking exists
    Set oMarkingPart = oMarkingLineAE.GetMfgMarkingPart
    
    'Get the Port on which the marking exists
    Set oPort = oMarkingLineAE.GetMfgMarkingPortForPlatePart(oMarkingPart)
    
    'Get Marking Line Side.
    'Important : Cannot use directly "oMarkingLineAE.MfgMarkingSide" because this may give sides pertaining to direction like Aft, Fore, etc.
    ' Instead use "GetMarkingEnumMfgSide" which gives the actual plate side (Base/Offset)except for Manufacturing Upside/ Anti - Manufacturing Upside

    Dim lMarkingSide As Long
    oMarkingLineAE.GetMarkingEnumMfgSide lMarkingSide
    
    lSide = lMarkingSide

    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function GetCodeListStringValue(ByVal strCodelistTableName As String, ByVal lValue As Long) As String
Const METHOD = "GetCodeListStringValue"
On Error GoTo ErrorHandler

    GetCodeListStringValue = vbNullString

    Dim oCodeListMetaData   As IJDCodeListMetaData
    Set oCodeListMetaData = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    If Not oCodeListMetaData Is Nothing Then
        GetCodeListStringValue = oCodeListMetaData.ShortStringValue(strCodelistTableName, lValue)
    End If

    Set oCodeListMetaData = Nothing
    
    Exit Function
    
ErrorHandler:
     Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

' ***********************************************************************************
' Private Function SetMarkinLineNameOnGeom3DObjects()
'
' Description:  Function gets the appropriate plate marking line AE from each Geom3d object
'               and set the name of the marking line on Geom3D object
'               Input arguments: GeomCol3d
' ***********************************************************************************
Public Function SetMarkinLineNameOnGeom3DObjects(oGeomCol3d As IJMfgGeomCol3d) As IJMfgGeomCol3d
Const METHOD = "SetMarkinLineNameOnGeom3DObjects"
On Error GoTo ErrorHandler

    Dim oGeom3d         As IJMfgGeom3D
    Dim lLoopCounter    As Long
    Dim oMarkingLineAE  As IJMfgMarkingLines_AE
    Dim lThicknessSide  As Long

    If Not oGeomCol3d Is Nothing Then
        For lLoopCounter = 1 To oGeomCol3d.GetCount
            Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
    
            'Get the Marking Line AE from Geom3D object
            Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
            
            Dim oMarkingNI As IJNamedItem
            Set oMarkingNI = oMarkingLineAE
            
            Dim oGeom3DNI As IJNamedItem
            Set oGeom3DNI = oGeom3d
            
            If Not oMarkingNI.Name = vbNullString Then
                oGeom3DNI.Name = oMarkingNI.Name
            End If
            
            Set oMarkingLineAE = Nothing
            Set oGeom3d = Nothing
        Next lLoopCounter
        
        Set SetMarkinLineNameOnGeom3DObjects = oGeomCol3d
    End If

    Exit Function
    
ErrorHandler:
     Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************************************
' Public Function CreateLapConnFittingMark()
' Description: Created Lap Connection Fitting Mark
' ***********************************************************************************
Public Function CreateLapConnFittingMark(oMarkPointPos As IJDPosition, oThisPartSuface As IUnknown, oConnPartSurface As IUnknown, dLocationFittingMarkLength As Double, oVector As IJDVector, Optional bMarkOnPlate As Boolean) As IJComplexString
    Const METHOD = "CreateLapConnFittingMark"
    On Error GoTo ErrorHandler

    Dim xStart As Double
    Dim yStart As Double
    Dim zStart As Double
    Dim xEnd As Double
    Dim yEnd As Double
    Dim zEnd As Double

    Dim oFaceSufaceBody As IJSurfaceBody
    Set oFaceSufaceBody = oConnPartSurface

    Dim oFaceNormalLineCS As IJComplexString
    Set oFaceNormalLineCS = New ComplexString3d

    'project oMarkPointPos on both surfaces
    Dim oVectorToThis As IJDVector
    Dim oVectorToConn As IJDVector
    Dim oMarkPointPosThisSurf As IJDPosition
    Dim oMarkPointPosConnSurf As IJDPosition

    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
   
    'Project mark point onto the surfaces
    Set oMarkPointPosThisSurf = oMfgRuleHelper.ProjectPointOnSurface(oMarkPointPos, oThisPartSuface, oVectorToThis)
    Set oMarkPointPosConnSurf = oMfgRuleHelper.ProjectPointOnSurface(oMarkPointPos, oConnPartSurface, oVectorToConn)

    ' Create normal to conn plate surface
    Dim oFaceNormalVector As IJDVector
    oFaceSufaceBody.GetNormalFromPosition oMarkPointPosConnSurf, oFaceNormalVector
    oMarkPointPos.Get xStart, yStart, zStart
    oFaceNormalVector.Length = 1
    
    Dim oFittingMarkVector As IJDVector
    
    If bMarkOnPlate = False Then
        ' If the fitting mark to be on the plate then get cross product of ref lap curve vector(oVector) and plate face normal
        Set oFittingMarkVector = oVector.Cross(oFaceNormalVector)
    Else
        ' If the fitting mark to be on the profile then input expected to be fitting mark vector
        Set oFittingMarkVector = oVector
    End If
    
    oFittingMarkVector.Length = dLocationFittingMarkLength
    oFittingMarkVector.Get xEnd, yEnd, zEnd
    
    Dim oFaceNormalLine As IngrGeom3D.Line3d
    Set oFaceNormalLine = New IngrGeom3D.Line3d
    
    If bMarkOnPlate = False Then
        oFaceNormalLine.DefineBy2Points xStart - xEnd, yStart - yEnd, zStart - zEnd, xStart + xEnd, yStart + yEnd, zStart + zEnd
    Else
        oFaceNormalLine.DefineBy2Points xStart, yStart, zStart, xStart - xEnd, yStart - yEnd, zStart - zEnd
    End If
    
    oFaceNormalLineCS.AddCurve oFaceNormalLine, False
    
    ' ComplexStringAlongVectorOnToSurface which projects CS on surface fails if the location
    ' fitting mark is outside the plate outer contour so "On Error Resume Next" is added below
    ' to skip those failed cases.
    
    On Error Resume Next
    
    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper
    
    Dim oCS As IJComplexString
    
    If bMarkOnPlate = True Then
        
        'Set oCS = m_oMfgRuleHelper.ComplexStringAlongVectorOnToSurface(oConnPartSurface, oFaceNormalLineCS, oVectorToConn)
    
        ' As the connected part(profile) surface is not proper(i.e., solid like or closed surface),
        ' the above routine is failing, so the work around is below.
    
        Dim oWireBody   As IJWireBody
        oMfgMGHelper.ComplexStringToWireBody oFaceNormalLineCS, oWireBody
        
        Dim oOverLapWire As IJWireBody
        Set oOverLapWire = m_oMfgRuleHelper.GetCommonGeometry(oConnPartSurface, oWireBody)
        
        oMfgMGHelper.WireBodyToComplexString oOverLapWire, oCS
        
        If oCS.CurveCount = 0 Then
            Set CreateLapConnFittingMark = Nothing
        Else
            Dim oMfgGeomUtilWrapper As New MfgGeomUtilWrapper
            oMfgGeomUtilWrapper.ExtendWire oCS, 0.08     ' 0.08 is the extension distance beyond profile overlap
            Set oMfgGeomUtilWrapper = Nothing
            
            ' Project the CS back on the oThisPartSuface surface
            Dim oProjCS     As IJComplexString
            Set oProjCS = m_oMfgRuleHelper.ComplexStringAlongVectorOnToSurface(oThisPartSuface, oCS, oVectorToThis)
            
            ' Return complex string
            If oProjCS Is Nothing Then
                Set CreateLapConnFittingMark = oCS
            Else
                Set CreateLapConnFittingMark = oProjCS
            End If
        End If
    Else
        Set oCS = m_oMfgRuleHelper.ComplexStringAlongVectorOnToSurface(oThisPartSuface, oFaceNormalLineCS, oVectorToThis)
        Set CreateLapConnFittingMark = oCS
    End If

CleanUp:
    Set oCS = Nothing
    Set oFaceSufaceBody = Nothing
    Set oFaceNormalLineCS = Nothing
    Set oVectorToThis = Nothing
    Set oVectorToConn = Nothing
    Set oMarkPointPosThisSurf = Nothing
    Set oMarkPointPosConnSurf = Nothing
    Set oFaceNormalVector = Nothing
    Set oFaceNormalLine = Nothing
    Set oMfgMGHelper = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function GetAssemblySequenceObject(oPart As Object) As Object
Const sMETHOD = "GetAssemblySequenceObject"
On Error GoTo ErrorHandler
    
    Dim oAssm As IJAssembly
    Dim oChild As IJAssemblyChild
    
    Set oChild = oPart
    Set oAssm = oChild.Parent
    
    Dim oAssySeq As GSCADPlanningInterfaces.IJAssemblySequence
    Set oAssySeq = oAssm
    
    ' Normalize the sequence, to get rid of gaps in the sequence
    oAssySeq.Normalize
    
    Set GetAssemblySequenceObject = oAssySeq

CleanUp:

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Public Function CheckIfPartIsBracket(oPart As Object) As Boolean
Const sMETHOD = "CheckIfPartIsBracket"
On Error GoTo ErrorHandler
    
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    
    Set oSDPartSupport.Part = oPart
    
    Dim oSys As IJSystem
    oSDPartSupport.IsSystemDerivedPart oSys, True
    
    Dim oPlateUtils As IJPlateAttributes
    Set oPlateUtils = New PlateUtils
    
    If Not oSys Is Nothing Then
        If (oPlateUtils.IsBracketByPlane(oSys) Or oPlateUtils.IsTrippingBracket(oSys)) Then
            CheckIfPartIsBracket = True
        End If
    End If

CleanUp:

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

'--------------------------------------------------------------------------------------------------
' Abstract :
'   GetManufacturingProductionName( ThisPart (input), Connected_Part (input), Condition (input) ) as String.
'    If Condition = 0   ==> return Connected.PartName.
'    if Condition = 1   ==> return Assembly condition for connected part.
'       If In same assembly ==> Connected part name.
'       If In later stage assembly ==> No name.
'       if In child-assembly of my-assembly ==> Assembly name.
'       If In <no assembly> ==> Connected part name.
'--------------------------------------------------------------------------------------------------
Public Function GetManufacturingProductionName(oThisPart As Object, oConnectedPart As Object, lCondition As Long) As String
Const METHOD = "GetManufacturingProductionName"
On Error GoTo ErrorHandler
    
    ' If lCondition = 0 then just return the connected part name
    Dim oNamedItem      As IJNamedItem
    If lCondition = 0 Then
        Set oNamedItem = oConnectedPart
        GetManufacturingProductionName = oNamedItem.Name
        Exit Function
    End If
    
    Dim oAssemblyChild As IJAssemblyChild
    Set oAssemblyChild = oThisPart
    
    Dim oThisPartAssy   As IJPlanningAssembly
    
    ' Check if ThisPartAssembly is Nothing
    If oAssemblyChild.Parent Is Nothing Then
        GetManufacturingProductionName = vbNullString
        Exit Function
    End If
    
    ' Check if ThisPartAssembly is Planning assembly
    If TypeOf oAssemblyChild.Parent Is IJPlanningAssembly Then
        Set oThisPartAssy = oAssemblyChild.Parent
        Set oAssemblyChild = Nothing
        Set oAssemblyChild = oConnectedPart
        
        ' Check if ConnPartAssembly is Nothing
        If oAssemblyChild.Parent Is Nothing Then
            GetManufacturingProductionName = vbNullString
            Exit Function
        End If
        
        ' Check if ConnPartAssembly is Planning assembly
        If TypeOf oAssemblyChild.Parent Is IJPlanningAssembly Then
            If oAssemblyChild.Parent Is oThisPartAssy Then
                Set oNamedItem = oConnectedPart
                GetManufacturingProductionName = oNamedItem.Name
                Exit Function
            End If
        Else
            If TypeOf oAssemblyChild.Parent Is IJConfigProjectRoot Then
                Set oNamedItem = oConnectedPart
                GetManufacturingProductionName = oNamedItem.Name
            Else
                GetManufacturingProductionName = vbNullString
            End If
            
            Exit Function
        End If
    Else
        GetManufacturingProductionName = vbNullString
        Exit Function
    End If
    
    
    Dim oPlnIntHelper As GSCADPlnIntHelper.IJDPlnIntHelper
    Set oPlnIntHelper = New CPlnIntHelper
    
    Dim oTempAllAssyChildren   As IJElements
    Set oTempAllAssyChildren = oPlnIntHelper.GetStoredProcAssemblyChildren(oThisPartAssy, "IJMfgParent", False, Nothing, False)
    
    ' If ThisPartAssembly doesnot have ConnPart then return NULL string
    If Not oTempAllAssyChildren.Contains(oConnectedPart) Then
        GetManufacturingProductionName = vbNullString
        Exit Function
    End If
    
    ' Get first hierarchy children assemblies of ThisAssembly
    Dim oThisAssemblyChildren   As IJElements
    Set oThisAssemblyChildren = oPlnIntHelper.GetAssemblyChildren(oThisPartAssy, "IJPlanningAssembly", False)
    
    Dim icount      As Long
    For icount = 1 To oThisAssemblyChildren.Count
        
        ' Get all the parts of first hierarchy children assembly
        Dim oTempAssemblyChildren   As IJElements
        Set oTempAssemblyChildren = oPlnIntHelper.GetStoredProcAssemblyChildren(oThisAssemblyChildren.Item(icount), "IJMfgParent", False, Nothing, False)
        
        ' If first hierarchy children assembly has the connected part then return the corresponding assembly name
        If oTempAssemblyChildren.Contains(oConnectedPart) Then
            Set oNamedItem = oThisAssemblyChildren.Item(icount)
            GetManufacturingProductionName = oNamedItem.Name
            Exit Function
        End If
    Next
    
 Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'**************************************************************************************
' Method       : GetProductionRoutingStageCode
' Abstract     : This Function gets the ProductionRouting stage code
'**************************************************************************************
Public Function GetProductionRoutingStageCode(oPart As Object) As String
Const METHOD = "GetProductionRoutingStageCode"
On Error GoTo ErrorHandler
    
    If oPart Is Nothing Then Exit Function
    
    Dim oPlnProdRouting As PlanningObjects.PlnProdRouting
    Set oPlnProdRouting = New PlanningObjects.PlnProdRouting
    Set oPlnProdRouting.object = oPart
    
    GetProductionRoutingStageCode = oPlnProdRouting.GetProductionRoutingCode
    
    Set oPlnProdRouting = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'**************************************************************************************
' Method       : CreateSpecialEndConnMarks
' Abstract     : This Function gets the Profile part and Plate/Profile Port object as input and gives the
'                EndConnectionMarks as a collection as output
'
'**************************************************************************************

Public Function CreateSpecialEndConnMarks(ByVal oProfilePart As IJProfilePart, ByVal oPort As IJPort) As IJElements
    Const METHOD = "CreateSpecialEndConnMarks"
    On Error GoTo ErrorHandler
    
    '** Declarations required for generating the mark**'
    Dim oMfgMGHelper As IJMfgMGHelper
    Dim oWebLeftPort As IJPort
    Dim oTopFacePort As IJPort
    
    Dim oWebLeftSurfaceBody As IJSurfaceBody
    Dim oTopSurfaceBody As IJSurfaceBody
    Dim oPartSurface As IJSurfaceBody
    Dim oWebLeftWB As IJWireBody
    Dim oWebLeftCS As IJComplexString
    Dim oSDProfileWrapper As StructDetailObjects.ProfilePart
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    Set oSDProfileWrapper.object = oProfilePart
    Dim oMfgGeomUtilWrapper As New MfgGeomUtilWrapper
    
    'Create an instance of the StrMfg Math Geom helper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    '**Start of Code for End Connection Mark SKDY Specific**
        
    '**Getting the WebLeft and TopFace Port**'
    Set oWebLeftPort = oSDProfileWrapper.SubPort(JXSEC_WEB_LEFT)
    Set oTopFacePort = oSDProfileWrapper.SubPort(JXSEC_TOP)
    
    '**Getting the Geometry of the Webleft and TopFace**'
    Set oWebLeftSurfaceBody = oWebLeftPort.Geometry
    Set oTopSurfaceBody = oTopFacePort.Geometry
    
    '**Getting the Plate Surface to which profiles are connected**'
    Set oPartSurface = oPort.Geometry
    
    Dim oGeomOffset As IJGeometryOffset
    Set oGeomOffset = New DGeomOpsOffset
    
    Dim dExtend As Double
    dExtend = 0.005
    
    On Error Resume Next ' As the Extend sheet body routine may fail
    
    'Extendind the SurfaceBody to ensure that intersection occurs between plate and WebLeft of Profile
    Dim oWebLeftExtendedSB As IJSurfaceBody
    oGeomOffset.CreateExtendedSheetBody Nothing, oWebLeftSurfaceBody, Nothing, dExtend, Nothing, oWebLeftExtendedSB
    
    If oWebLeftExtendedSB Is Nothing Then
        Set oWebLeftExtendedSB = oWebLeftSurfaceBody
    End If
    
    'Extendind the SurfaceBody to ensure that intersection occurs between plate and WebLeft of Profile
    Dim oExtendedPlateSB As IJSurfaceBody
    oGeomOffset.CreateExtendedSheetBody Nothing, oPartSurface, Nothing, dExtend, Nothing, oExtendedPlateSB
    
    If oExtendedPlateSB Is Nothing Then
        Set oExtendedPlateSB = oPartSurface
    End If
    
    '**Getting the intersections of WebLeft with Plate Surface and also TopFace with Plate Surface and assigning it to WireBody
    Dim oExtendedWB As IJWireBody
    Set oExtendedWB = GetIntersection(oWebLeftExtendedSB, oExtendedPlateSB)
    
    If oExtendedWB Is Nothing Then
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    'Trimming the WebLeft Geometry Generated as it will be more than the WebLeft Geometry
    oMfgMGHelper.WireBodyToComplexString oExtendedWB, oWebLeftCS
    m_oMfgRuleHelper.TrimCurveEnds oWebLeftCS, dExtend
    
    Set oWebLeftWB = m_oMfgRuleHelper.ComplexStringToWireBody(oWebLeftCS)
    
    Dim oPointOnLine As IJDPosition
    Dim oWebLeftLine As IJLine
    
    'Getting the Point for Creating the Horizontal line
    CreateLineByExtendingWireToInputSurface oWebLeftWB, oTopSurfaceBody, 0.001, oWebLeftLine, oPointOnLine
    
    '**Finding Vector Normal to the WebLeft Surface**'
    
    Dim oCOG As IJDPosition                 'Center of Gravity of Web Left Surface
    Dim oProjectionPoint As IJDPosition
    Dim oProjectionVector As IJDVector
    Dim oSurfaceNormal As IJDVector         'Vector Normal to Web Left Surface
    
    '***Setting the COG
    oWebLeftSurfaceBody.GetCenterOfGravity oCOG
    
    oMfgMGHelper.ProjectPointOnSurfaceBody oWebLeftSurfaceBody, oCOG, oProjectionPoint, oProjectionVector
    oWebLeftSurfaceBody.GetNormalFromPosition oProjectionPoint, oSurfaceNormal
    
    'Normalizing The Vector
    oSurfaceNormal.Length = 1
    
    Dim oTopLineLeftPoint As IJDPosition
    Dim oTopLineRightPoint As IJDPosition
    
    Set oTopLineLeftPoint = New DPosition
    Set oTopLineRightPoint = New DPosition
    
    'Generating the Horizontal Line which indicate the flange
    oTopLineLeftPoint.x = oPointOnLine.x + (LEFT_EXTENDED_LENGTH * oSurfaceNormal.x)
    oTopLineLeftPoint.y = oPointOnLine.y + (LEFT_EXTENDED_LENGTH * oSurfaceNormal.y)
    oTopLineLeftPoint.z = oPointOnLine.z + (LEFT_EXTENDED_LENGTH * oSurfaceNormal.z)
    
    oTopLineRightPoint.x = oPointOnLine.x + (RIGHT_EXTENDED_LENGTH * (-oSurfaceNormal.x))
    oTopLineRightPoint.y = oPointOnLine.y + (RIGHT_EXTENDED_LENGTH * (-oSurfaceNormal.y))
    oTopLineRightPoint.z = oPointOnLine.z + (RIGHT_EXTENDED_LENGTH * (-oSurfaceNormal.z))
    
    Dim oTopFaceLine As IJLine
    Set oTopFaceLine = New Line3d
    
    oTopFaceLine.DefineBy2Points oTopLineLeftPoint.x, oTopLineLeftPoint.y, oTopLineLeftPoint.z, oTopLineRightPoint.x, oTopLineRightPoint.y, oTopLineRightPoint.z
    
    Dim oTopFaceLineCS As IJComplexString
    Set oTopFaceLineCS = New ComplexString3d
    
    Dim oWebLeftLineCS As IJComplexString
    Set oWebLeftLineCS = New ComplexString3d
    
    oTopFaceLineCS.AddCurve oTopFaceLine, True
    oWebLeftLineCS.AddCurve oWebLeftLine, True
    
    'Extending the Web Left Geometry to Ensure That Even if Feature(scallop) is Present After Detailing the Total Web Left Length is Seen in Part Monitor
    oMfgGeomUtilWrapper.ExtendWire oWebLeftLineCS, 0.2, 1
    
    Dim oWebLeftLineProjCS As IJElements
    'Projecting The Web Left Line on the Plate Surface that will be manufactured
    oMfgMGHelper.ProjectCSToSurface oWebLeftLineCS, oPartSurface, Nothing, oWebLeftLineProjCS
    Dim oWebLeftCurve As IJCurve
    Set oWebLeftCurve = oWebLeftLineProjCS.Item(1)
    
    'Checking the Cases Where the Profiles may have fillet at the top and web left surface does not cover it
    If oSDProfileWrapper.WebLength - oWebLeftCurve.Length > 0.001 Then
        oMfgGeomUtilWrapper.ExtendWire oWebLeftCurve, (oSDProfileWrapper.WebLength - oWebLeftCurve.Length), 0
    End If
    
    Dim oCSColl As IJElements
    Set oCSColl = New JObjectCollection
    oCSColl.Add oTopFaceLineCS
    oCSColl.Add oWebLeftCurve
    
    Set CreateSpecialEndConnMarks = oCSColl
    
CleanUp:
    Set oMfgMGHelper = Nothing
    Set oWebLeftPort = Nothing
    Set oTopFacePort = Nothing
    Set oWebLeftSurfaceBody = Nothing
    Set oTopSurfaceBody = Nothing
    Set oWebLeftExtendedSB = Nothing
    Set oTopFaceLine = Nothing
    Set oTopLineLeftPoint = Nothing
    Set oTopLineRightPoint = Nothing
    Set oSurfaceNormal = Nothing
    Set oProjectionVector = Nothing
    Set oProjectionPoint = Nothing
    Set oCOG = Nothing
    Set oPointOnLine = Nothing
    Set oWebLeftLine = Nothing
    Set oPartSurface = Nothing
    Set oExtendedPlateSB = Nothing
    Set oWebLeftWB = Nothing
    Set oWebLeftCS = Nothing
    Set oGeomOffset = Nothing
    Set oSDProfileWrapper = Nothing
    Set oMfgGeomUtilWrapper = Nothing
    Set oWebLeftCurve = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'**************************************************************************************
' Method       : CreateLineByExtendingWireToInputSurface
' Abstract     : In this Function given a wire body and a surface,it tries to create a line
'                by extending wire to surface.
'
'       \  Wire
'        \
'         \
'          '  (Line created includes wire portion also)
'           '
'            '-------------  Surface
'**************************************************************************************

Public Function CreateLineByExtendingWireToInputSurface( _
            ByVal oWireBody As IJWireBody, _
            ByVal oAnotherSurface As Object, _
            ByVal dDistanceTolerance As Double, _
            ByRef oCreatedLine As IJLine, ByRef oTopPos As IJDPosition)
    
    On Error GoTo ErrorHandler
   
    Dim oModelBodyUtils As SGOModelBodyUtilities
    Dim oPointOnWire As IJDPosition
    Dim oPointOnSurface As IJDPosition
    Dim dDistance As Double
   
    Set oCreatedLine = Nothing
    Set oModelBodyUtils = New SGOModelBodyUtilities
    
    'Getting the minimum distance betwwen Profile Top Surface and Web Left Line
    oModelBodyUtils.GetClosestPointsBetweenTwoBodies _
                    oWireBody, oAnotherSurface, _
                     oPointOnWire, oPointOnSurface, dDistance
    
    Dim oTempCS As IJComplexString
    Dim oCurve As IJCurve
    Dim oMfgRuleHelper As MfgRuleHelpers.Helper
   
    Dim oCurveSP As IJDPosition
    Dim oCurveEP As IJDPosition
    Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
    Dim dOldStartX As Double, dOldStartY As Double, dOldStartZ As Double, dOldEndX As Double, dOldEndY As Double, dOldEndZ As Double
  
    Set oMfgRuleHelper = New MfgRuleHelpers.Helper
    Set oTempCS = oMfgRuleHelper.WireBodyToComplexString(oWireBody)
    
    Set oCurve = oTempCS
    oCurve.EndPoints dOldStartX, dOldStartY, dOldStartZ, dOldEndX, dOldEndY, dOldEndZ
    Set oCurve = Nothing
    
    'Ensuring that the Web Left Line is Offset by the Required Distance
    oMfgRuleHelper.ExtendWire oTempCS, dDistance + ENDCONN_EXTENSION
   
    Set oCurve = oTempCS
    oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
  
    Set oCurveSP = New DPosition
    Set oCurveEP = New DPosition
     
    oCurveSP.Set dStartX, dStartY, dStartZ
    oCurveEP.Set dEndX, dEndY, dEndZ
    
    'Creating the Required Line And Getting the Required point
    Set oCreatedLine = New Line3d
    Set oTopPos = New DPosition
    
    If oPointOnSurface.DistPt(oCurveEP) < oPointOnSurface.DistPt(oCurveSP) Then
        oTopPos.Set dEndX, dEndY, dEndZ
        oCreatedLine.DefineBy2Points _
                         oTopPos.x, oTopPos.y, oTopPos.z, _
                        dOldStartX, dOldStartY, dOldStartZ
    Else
        oTopPos.Set dStartX, dStartY, dStartZ
        oCreatedLine.DefineBy2Points _
                         oTopPos.x, oTopPos.y, oTopPos.z, _
                         dOldEndX, dOldEndY, dOldEndZ
    End If
    
CleanUp:
    Set oModelBodyUtils = Nothing
    Set oPointOnWire = Nothing
    Set oCurveSP = Nothing
    Set oPointOnSurface = Nothing
    Set oCurveEP = Nothing
    Set oCurve = Nothing

    Exit Function
   
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function

'--------------------------------------------------------------------------------------------------
' Abstract : The purpose of this routine is to calculate the intersection between any two given
'            objects. The call will be delegated to the G&T PlaceIntersectionObject routine
'--------------------------------------------------------------------------------------------------
Public Function GetIntersection(pIntersectedObject As Object, pIntersectingObject As Object) As Object
On Error GoTo ErrorHandler
Const METHOD = "GetIntersection"

    ' Find the intersection.
    Dim oGeometryIntersector    As IMSModelGeomOps.DGeomOpsIntersect
    Set oGeometryIntersector = New IMSModelGeomOps.DGeomOpsIntersect
    
    On Error Resume Next 'Needed for continuing with next skid mark if intersection fails
    Dim oIntersectionUnknown    As IUnknown        ' Resultant intersection.
    oGeometryIntersector.PlaceIntersectionObject Nothing, pIntersectedObject, pIntersectingObject, Nothing, oIntersectionUnknown
    
    On Error GoTo ErrorHandler
    Set GetIntersection = oIntersectionUnknown
    Set oGeometryIntersector = Nothing
    Set oIntersectionUnknown = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
