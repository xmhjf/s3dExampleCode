VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlgOutsideSnipeTubeParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : FlgOutsideSnipeTubeParm.cls
'
'  Description :
'       Sets parameters for the FlgOutsideSnipeTubeParm cut.
'
'  Author      : Alligators
'
'  History     :
'    08/Aug/2011 - Created
'    22/Aug/2011 svsmylav B-25298(CR-179926): corrected logic to identify top flange cut
'    and bottom flange cut properly. Also, used absolute value of dot product
'    to determine bounded is orthogonal to bounding case.
'    06/Sep/2011 - pnalugol for supporting Generic ACs for Tube standard(Bounding)Vs non-tube Standard member(Bounded)
'    30/Nov/2011 - svsmylav TR-205302: Replaced '.Subport' method call with 'GetLateralSubPortBeforeTrim'.
'    10/Feb/2011 - CM - TR210357 : Bounded Object was not set appropriately.
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
'*********************************************************************************************

Const m_sClassName As String = "FlgOutsideSnipeTubeParm"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Private Const TOLERANCE_VALUE = 0.0000001

Implements IJDUserSymbolServices
Implements IJEndCutDrawingType

'*********************************************************************************************
' Method      : ParameterRuleInputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleInputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleOutputs
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleOutputs(pOH As IJDOutputsHelper)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleOutputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    pOH.SetOutput "WebClearance"
    pOH.SetOutput "Nose"
    pOH.SetOutput "Slope"
    pOH.SetOutput "Radius"
    
    ' To control the Solid projection distance
    pOH.SetOutput "CutDepth"
    pOH.SetOutput "BottomFlange"
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ParameterRuleLogic
' Description :
'
'*********************************************************************************************
Public Sub ParameterRuleLogic(pPRL As IJDParameterLogic)
    Const METHOD = m_DefinitionProgid & "::ParameterRuleLogic"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ParameterRule Outputs"
    
    Dim oPortBounded As IJPort
    Dim oPortBounding As IJPort
    
    Set oPortBounded = pPRL.InputObject(INPUT_BOUNDED)
    Set oPortBounding = pPRL.InputObject(INPUT_BOUNDING)
    
    Dim oBoundingObject As StructDetailObjects.MemberPart
    Set oBoundingObject = New StructDetailObjects.MemberPart
    
    Dim bBuiltup As Boolean, oBUMember As ISPSMemberPartCommon
    Dim oDMember As ISPSDesignedMember
    
    If TypeOf oPortBounding.Connectable Is ISPSMemberPartCommon Then
        Set oBoundingObject.object = oPortBounding.Connectable
    Else
        IsFromBuiltUpMember oPortBounding.Connectable, bBuiltup, oBUMember
        If bBuiltup Then
            Set oBoundingObject.object = oBUMember
        Else
            GoTo ErrorHandler
        End If
    End If
    
    Dim oBoundingISPS As ISPSMemberPartCommon
    Set oBoundingISPS = oBoundingObject.object

    Dim dPI As Double
    dPI = Atn(1) * 4#
    
    'Set Web-clearance
    Dim dWebClearance As Double
    Get_AttributeValue pPRL.SmartOccurrence, "WebClearance", dWebClearance
    If dWebClearance < TOLERANCE_VALUE Then dWebClearance = 0.015 '15mm
    pPRL.Add "WebClearance", dWebClearance

    'Set Nose
    pPRL.Add "Nose", 0.015
        
    'Set Slope
    Dim sShapeOutSide As String
    Dim sParentItemName As String
    Dim oParentObj As Object
    Dim eCScode As JXSEC_CODE
    
    'Get answer to know if current case is bottom flange
    Dim sBottomFlange As String
    GetSelectorAnswer pPRL, "BottomFlange", sBottomFlange
    
    If StrComp(sBottomFlange, "No", vbTextCompare) <> 0 Then
        eCScode = JXSEC_TOP
    Else
        eCScode = JXSEC_BOTTOM
    End If
    
    'Determine whether this is a Top or Bottom flange cut
    Dim bIsTopFlgCut As Boolean
    If TypeOf pPRL.SmartItem Is IJSmartItem Then
        Dim oSmartItem As IJSmartItem
        Dim oSmartClass As IJSmartClass

        Set oSmartItem = pPRL.SmartItem
        Set oSmartClass = oSmartItem.Parent

        If oSmartClass.SCSubType = 313 Then
            bIsTopFlgCut = True 'Current smart item is Top-flange cut
        Else
            bIsTopFlgCut = False 'Current smart item is Bottom-flange cut
        End If
    End If
    
    
    'Set 'Radius' value: we need to consider one of the following two cases:
    'Case I: If the current case is such that the bounded flange-top/bottom plane (or surface) is perpendicular to the bounding axis at the bounded location,
    ' then the bounding cross section is a circle in this plane and Radius of the Tube member is to be used:
    
    'Get member connection data
    'Initialze endcut data
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    
    Parent_SmartItemName pPRL.SmartOccurrence, sParentItemName, oParentObj
    
    Do While Not TypeOf oParentObj Is IJAppConnection
        Parent_SmartItemName oParentObj, sParentItemName, oParentObj
    Loop

    Set oAppConnection = oParentObj
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If oBoundedData.MemberPart Is oBoundingData.MemberPart Then 'to support generic connections
        Set oBoundingData.MemberPart = oBoundingISPS
        Set oBoundingData.AxisCurve = oBoundingISPS.Axis
        
        'set default to Straight
        pPRL.Add "Slope", 90 * dPI / 180
        
    Else    'member Axis ACs
    
        GetSelectorAnswer oParentObj, "ShapeOutside", sShapeOutSide
        If sShapeOutSide = "" Then
            'Need to check if this is Top flange cut or else bottom flange cut
            If bIsTopFlgCut Then
                GetSelectorAnswer oParentObj, "TopShapeOutside", sShapeOutSide
            Else
                GetSelectorAnswer oParentObj, "BottomShapeOutside", sShapeOutSide
            End If
        End If
        
        If sShapeOutSide = gsStraight Then
            pPRL.Add "Slope", 90 * dPI / 180
        ElseIf sShapeOutSide = gsSniped Then
            pPRL.Add "Slope", 65 * dPI / 180
        Else
            pPRL.Add "Slope", 45 * dPI / 180
        End If
    End If
    
    'Get web cut position
    Dim oEndCutPos As IJDPosition
    Dim oBoundingPos As IJDPosition
    Dim oBoundedPos As IJDPosition
    Dim dDistance As Double

    GetDistanceFromBounding oPortBounding.Connectable, oPortBounded.Connectable, eCScode, _
                oBoundingPos, oBoundedPos
    
    'Get bounding axis curve to center
    Dim oCrv As IJCurve
    Set oCrv = GetAxisCurveAtTubeCenter(oPortBounding, oBoundingPos, oBoundingData)
    
    'Determine closest point on top/bottom port
    Dim oBounded As Object
    Dim oBoundedTopOrBotPort As IJPort
    
    Set oBounded = oPortBounded.Connectable
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        Dim oBoundedMbr As StructDetailObjects.MemberPart
        Set oBoundedMbr = New MemberPart
        Set oBoundedMbr.object = oBounded
        Set oBoundedTopOrBotPort = GetLateralSubPortBeforeTrim(oBounded, eCScode)
    ElseIf TypeOf oBounded Is IJProfile Then
        Dim oBoundedProfile As StructDetailObjects.ProfilePart
        Set oBoundedProfile = New ProfilePart
        Set oBoundedProfile.object = oBounded
        Set oBoundedTopOrBotPort = GetLateralSubPortBeforeTrim(oBounded, eCScode)
    Else
    
    End If
    
    Dim oSurface As IJSurface
    Set oSurface = oBoundedTopOrBotPort
    Dim dSrcX As Double, dSrcY As Double, dSrCZ As Double
    Dim dInX As Double, dInY As Double, dInZ As Double
    oCrv.DistanceBetween oSurface, dDistance, dSrcX, dSrcY, dSrCZ, dInX, dInY, dInZ
    Set oEndCutPos = New DPosition
    oEndCutPos.Set dInX, dInY, dInZ
    oBoundingPos.Set dSrcX, dSrcY, dSrCZ
    
    'Notes:-
    ' Matrix.IndexValue(0,1,2)   : U is direction Along Axis
    ' Matrix.IndexValue(4,5,6)   : V is direction normal to Web (from Web Right to Web Left)
    ' Matrix.IndexValue(8,9,10)  : W is direction normal to Flange (from Flange Bottom to Flange Top)
    ' Matrix.IndexValue(12,13,14): Root/Origin Point

    'Get bounding matrix
    oBoundingData.MemberPart.Rotation.GetTransformAtPosition oBoundingPos.x, oBoundingPos.y, oBoundingPos.z, oBoundingData.Matrix, Nothing

    'Get bounding U-direction vector
    Dim oBounding_UDir As IJDVector
    Set oBounding_UDir = New DVector
    oBounding_UDir.Set oBoundingData.Matrix.IndexValue(0), oBoundingData.Matrix.IndexValue(1), oBoundingData.Matrix.IndexValue(2)

    'Get bounded matrix
    oBoundedData.MemberPart.Rotation.GetTransformAtPosition oEndCutPos.x, oEndCutPos.y, oEndCutPos.z, oBoundedData.Matrix, Nothing

    'Get bounding W-direction vector
    Dim oBounded_WDir As IJDVector
    Set oBounded_WDir = New DVector
    oBounded_WDir.Set oBoundedData.Matrix.IndexValue(8), oBoundedData.Matrix.IndexValue(9), oBoundedData.Matrix.IndexValue(10)
    
    'For Case-I, the dot product of oBounding_UDir and oBounded_WDir would be zero
    Dim dRadius As Double
    dRadius = oBoundingObject.Height / 2 'Default value
    If Abs(1# - Abs(oBounding_UDir.Dot(oBounded_WDir))) < TOLERANCE_VALUE Then
        'Retain default value i.e. dRadius = oBoundingObject.Height / 2
    Else
        'Case II: Compute Radius by calculating distance between two points:
        ' i) First Point: intersection with bounding tube center line(or curve) with the Web-Center plane
        ' ii)Second Point is intersection with bounding Tube:
        '  Whenever bounded member is rotated about its axis such that web-plane is not perpendicular to the bounding axis,
        ' we see bounding cross section as an ellipse in the web-plane; web cut need to be created at 'WebClearnace' offset
        ' from center of the bounding tube (offset direction will be along axis of the bounded into the material). For
        ' this a straightline is created at 'WebClearance' distance and intersected with the bounding tube.

        ' i) First point:-
        Dim oGeometryFactory As IngrGeom3D.GeometryFactory
        Set oGeometryFactory = New IngrGeom3D.GeometryFactory
        
        Dim oBounded_VDir As IJDVector
        Set oBounded_VDir = New DVector
        oBounded_VDir.Set oBoundedData.Matrix.IndexValue(4), oBoundedData.Matrix.IndexValue(5), oBoundedData.Matrix.IndexValue(6)

        'Create infinite plane with bounded-flange normal as its normal
        Dim oPlane As IngrGeom3D.Plane3d
        Set oPlane = oGeometryFactory.Planes3d.CreateByPointNormal(Nothing, _
                        oEndCutPos.x, oEndCutPos.y, oEndCutPos.z, _
                        oBounded_WDir.x, oBounded_WDir.y, oBounded_WDir.z)
                        
        
        'Find intersection
        Dim lNumIsec As Long
        Dim lNumOverlaps As Long
        Dim dPts() As Double
        Dim Code As Geom3dIntersectConstants
        oCrv.Intersect oPlane, lNumIsec, dPts, lNumOverlaps, Code

        'If intersection exists, get center point of bounding tube
        If Code = ISECT_UNKNOWN Then
            'Get bounding center-position
            Dim oCenterPos As IJDPosition
            Set oCenterPos = New DPosition
            oCenterPos.Set dPts(0), dPts(1), dPts(2)
        Else
            sMsg = "No intersection with the web-cut plane"
            GoTo ErrorHandler
        End If
        
        ' ii) Second point:-
        'Get offset position
        Dim oWCoffsetPos As IJDPosition
        Dim oBounded_UDir As IJDVector
        Set oBounded_UDir = New DVector

        oBounded_UDir.Set oBoundedData.Matrix.IndexValue(0), oBoundedData.Matrix.IndexValue(1), oBoundedData.Matrix.IndexValue(2)

        oBounded_UDir.Length = dWebClearance
        Set oWCoffsetPos = oCenterPos.Offset(oBounded_UDir)

        Dim oLine As IngrGeom3D.Line3d
        Set oLine = oGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, _
            oWCoffsetPos.x, oWCoffsetPos.y, oWCoffsetPos.z, oBounded_VDir.x, oBounded_VDir.y, oBounded_VDir.z, 3) '3 meters line
        Set oCrv = oLine
        Dim oTubeOuterPort As IJPort
        Set oTubeOuterPort = GetLateralSubPortBeforeTrim(oBoundingObject.object, JXSEC_OUTER_TUBE)
      
        Set oSurface = oTubeOuterPort
        oCrv.Intersect oSurface, lNumIsec, dPts, lNumOverlaps, Code
        If Code = ISECT_UNKNOWN Then
            Dim oPosition As IJDPosition
            Set oPosition = New DPosition
            oPosition.Set dPts(0), dPts(1), dPts(2)
            dRadius = oPosition.DistPt(oCenterPos)
        Else
            sMsg = "No intersection with the tube outer-port"
            GoTo ErrorHandler
        End If
    End If
    pPRL.Add "Radius", dRadius
        
    ' define the cutting depth to completely cut the Member cross Section
    Dim dCuttingDepth As Double
    dCuttingDepth = EndCut_GetCutDepth(pPRL)
    pPRL.Add "CutDepth", dCuttingDepth

    'Set bottom-flange parameter
    If sBottomFlange = "No" Then
        pPRL.Add "BottomFlange", 0
    Else
        pPRL.Add "BottomFlange", 1
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub
  
' ** Start CM **
'*********************************************************************************************
' If needed Add Custom Method HERE
'*********************************************************************************************
' ** End CM **

'*********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
'*********************************************************************************************

'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pPR As IJDSymbolDefinition)
    On Error Resume Next
    
    ' Remove all existing defined Input and Output (Representations) before defining the current Inputs and Outputs
    pPR.IJDInputs.RemoveAllInput
    pPR.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    Dim pIH As IJDInputsHelper
    Dim pOH As IJDOutputsHelper
    
    pDFact.InitAbstractParameterRule pPR
    
    Set pIH = New InputHelper
    pIH.definition = pPR
    pIH.InitAs m_FamilyProgid
    ParameterRuleInputs pIH
    
    Set pOH = New OutputHelper
    pOH.Representation = pPR.IJDRepresentations.Item(1)
    pOH.InitAs m_FamilyProgid
    ParameterRuleOutputs pOH
    
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateParameterRule(m_DefinitionProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)

End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

'*********************************************************************************************
' Method      : CMParameterRule
' Description :
'
'*********************************************************************************************
Public Sub CMParameterRule(pRep As IJDRepresentation)
    Dim pPRL As IJDParameterLogic
    Set pPRL = New ParameterLogic
    pPRL.Representation = pRep
    ParameterRuleLogic pPRL
  
End Sub
'*****************************************************************************************************
' Method Name : IJEndCutDrawingType_UpdateDrawingType
' Inputs      : feature occurence object
' Output      : string, returns the end cut type
' Description : this method checks for the type of endcut, straight or sniped and offset or no-offset
'                and assigns the corresponding enum to a string variable.
'******************************************************************************************************
Public Function IJEndCutDrawingType_UpdateDrawingType(ByVal pSmartOccurrence As Object) As String
    On Error GoTo ErrorHandler
    IJEndCutDrawingType_UpdateDrawingType = Sniped_No_Offset_FlangeCuts
    
    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.TopOrLeft, IJEndCutDrawingType_UpdateDrawingType
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
End Function
'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************
