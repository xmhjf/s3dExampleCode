VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutsideSnipeTubeParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : OutsideSnipeTubeParm.cls
'
'  Description :
'       Sets parameters for the OutsideSnipeTubeParm cut.
'
'  Author      : Alligators
'
'  History     :
'    27/Jul/2011 - Created
'    22/Aug/2011 svsmylav B-25298(CR-179926): corrected code to make use of GetSelectorAnswer
'    method call to get "ShapeOutside" before its use.
'   06/Sep/2011 - pnalugol for supporting Generic ACs for Tube standard(Bounding)Vs non-tube Standard member(Bounded)
'    30/Nov/2011 - svsmylav TR-205302: Replaced '.Subport' method call with 'GetLateralSubPortBeforeTrim'.
'    25/Jan/2016 - Modified dsmamidi\mkonduri
'                  CR-273576 Set the StartEndCutData and EndEndCutData fields on profile parts from SD rules
'*********************************************************************************************

Const m_sClassName As String = "OutsideSnipeTubeParm"
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
    pOH.SetOutput "FlangeThickness"
    
    ' To control the Solid projection distance
    pOH.SetOutput "CutDepth"
  
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

    Dim oBoundedObject As Object
    Set oBoundedObject = oPortBounded.Connectable

    Dim dFlgThickness As Double
    
    If TypeOf oBoundedObject Is ISPSMemberPartCommon Then
        Dim oBoundedMemberPart As StructDetailObjects.MemberPart
        Set oBoundedMemberPart = New MemberPart
        Set oBoundedMemberPart.object = oBoundedObject
        dFlgThickness = oBoundedMemberPart.flangeThickness
    ElseIf TypeOf oBoundedObject Is IJProfile Then
        Dim oBoundedProfilePart As StructDetailObjects.ProfilePart
        Set oBoundedProfilePart = New ProfilePart
        Set oBoundedProfilePart.object = oBoundedObject
        dFlgThickness = oBoundedProfilePart.flangeThickness
    Else
        'Specail need to be handled
    End If

    Dim oBoundingObject As StructDetailObjects.MemberPart
    Set oBoundingObject = New StructDetailObjects.MemberPart
    Set oBoundingObject.object = oPortBounding.Connectable

    Dim bTFL As Boolean, bTFR As Boolean, bBFL As Boolean, bBFR As Boolean
    CrossSection_Flanges oBoundedObject, bTFL, bBFL, bTFR, bBFR

    Dim dPI As Double
    dPI = Atn(1) * 4#

    'Get parent object
    Dim sShapeOutSide As String
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName pPRL.SmartOccurrence, sParentItemName, oParentObj
    
    Do While Not TypeOf oParentObj Is IJAppConnection
        Parent_SmartItemName oParentObj, sParentItemName, oParentObj
    Loop
    
    'Initialze endcut data
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long

    Set oAppConnection = oParentObj
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    Dim oBoundingSPS As ISPSMemberPartCommon
    Set oBoundingSPS = oBoundingObject.object
    
     'Determine whether this is a Top or Bottom webcut
    Dim bIsTopWebCut As Boolean
    If TypeOf pPRL.SmartItem Is IJSmartItem Then
        Dim oSmartItem As IJSmartItem
        Dim oSmartClass As IJSmartClass

        Set oSmartItem = pPRL.SmartItem
        Set oSmartClass = oSmartItem.Parent

        If oSmartClass.SCSubType = 261 Then
            bIsTopWebCut = True 'Current smart item is Top-web cut
        Else
            bIsTopWebCut = False 'Current smart item is Bottom-web cut
        End If
    End If
    
    If oBoundedData.MemberPart Is oBoundingData.MemberPart Then 'Generic Acs
        Set oBoundingData.MemberPart = oBoundingSPS
        Set oBoundingData.AxisCurve = oBoundingSPS.Axis
        'set default to Straight
        pPRL.Add "Slope", 90 * dPI / 180
        
    Else    'member Axis ACs
            'Set Slope
        GetSelectorAnswer oParentObj, "ShapeOutside", sShapeOutSide
        If sShapeOutSide = "" Then
            'Need to check if this is Top web cut or bottom web cut and get corresponding selector answer
            If bIsTopWebCut Then
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
 
    'Set Web-clearance
    Dim dWebClearance As Double
    Get_AttributeValue pPRL.SmartOccurrence, "WebClearance", dWebClearance
    If dWebClearance < TOLERANCE_VALUE Then dWebClearance = 0.015 '15mm
    pPRL.Add "WebClearance", dWebClearance
    
    'Set Nose
    Dim dDistanceFromBounding As Double
    If bIsTopWebCut Then
        If bTFL Or bTFR Then
            'Get distance above the bounding and set nose value
            dDistanceFromBounding = GetDistanceFromBounding(oBoundingObject.object, oBoundedData.MemberPart, JXSEC_TOP) _
                    - dFlgThickness
        End If
    Else
        If bBFL Or bBFR Then
            'Get distance below the bounding and set nose value
            dDistanceFromBounding = GetDistanceFromBounding(oBoundingObject.object, oBoundedData.MemberPart, JXSEC_BOTTOM) _
                    - dFlgThickness
        End If
    End If
    If (dDistanceFromBounding - 0.015) > TOLERANCE_VALUE Then
        pPRL.Add "Nose", 0.015
    Else
        pPRL.Add "Nose", 0.0001 '0.1mm
    End If

    'Set 'Radius' value: we need to consider one of the following two cases:
    'Case I: If the current case is such that the bounded web-center plane (or surface) is perpendicular to the bounding axis at the bounded location,
    ' then the bounding cross section is a circle in this plane and Radius of the Tube member is to be used:
    'Get member connection data
    'Get web cut position
    Dim oWebCutPos As IJDPosition
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut

    Set oSDO_WebCut.object = pPRL.SmartOccurrence
    Set oWebCutPos = oSDO_WebCut.BoundedLocation

    'Notes:-
    ' Matrix.IndexValue(0,1,2)   : U is direction Along Axis
    ' Matrix.IndexValue(4,5,6)   : V is direction normal to Web (from Web Right to Web Left)
    ' Matrix.IndexValue(8,9,10)  : W is direction normal to Flange (from Flange Bottom to Flange Top)
    ' Matrix.IndexValue(12,13,14): Root/Origin Point

    'Get bounding matrix
    oBoundingData.MemberPart.Rotation.GetTransformAtPosition oWebCutPos.x, oWebCutPos.y, oWebCutPos.z, oBoundingData.Matrix, Nothing

    'Get bounding U-direction vector
    Dim oBounding_UDir As IJDVector
    Set oBounding_UDir = New DVector
    oBounding_UDir.Set oBoundingData.Matrix.IndexValue(0), oBoundingData.Matrix.IndexValue(1), oBoundingData.Matrix.IndexValue(2)

    'Get bounded matrix
    oBoundedData.MemberPart.Rotation.GetTransformAtPosition oWebCutPos.x, oWebCutPos.y, oWebCutPos.z, oBoundedData.Matrix, Nothing

    'Get bounding W-direction vector
    Dim oBounded_WDir As IJDVector
    Set oBounded_WDir = New DVector
    oBounded_WDir.Set oBoundedData.Matrix.IndexValue(8), oBoundedData.Matrix.IndexValue(9), oBoundedData.Matrix.IndexValue(10)
    
    'For Case-I, the dot product of oBounding_UDir and oBounded_WDir would be zero
    Dim dRadius As Double
    dRadius = oBoundingObject.Height / 2 'Default value
    If Abs(oBounding_UDir.Dot(oBounded_WDir)) < TOLERANCE_VALUE Then
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

        'Create infinite plane with bounded-web normal as its normal
        Dim oSurface As IJSurface
        Dim oPlane As IngrGeom3D.Plane3d
        Set oPlane = oGeometryFactory.Planes3d.CreateByPointNormal(Nothing, _
                        oWebCutPos.x, oWebCutPos.y, oWebCutPos.z, _
                        oBounded_VDir.x, oBounded_VDir.y, oBounded_VDir.z)
                        
        'Get bounding axis curve
        Dim oCrv As IJCurve
        Set oCrv = GetAxisCurveAtTubeCenter(oPortBounding, oWebCutPos, oBoundingData)
        
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
        Set oWCoffsetPos = oCenterPos.offset(oBounded_UDir)

        'Get second intersection point
        If bIsTopWebCut = False Then oBounded_WDir.Length = -1 'Reverse W vector

        Dim oLine As IngrGeom3D.Line3d
        Set oLine = oGeometryFactory.Lines3d.CreateByPtVectLength(Nothing, _
            oWCoffsetPos.x, oWCoffsetPos.y, oWCoffsetPos.z, oBounded_WDir.x, oBounded_WDir.y, oBounded_WDir.z, 3) '3 meters line
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

    'Set FlangeThickness
    Dim dFlangeThickness As Double
    dFlangeThickness = 0.00001 '0.01mm initial value
    If bIsTopWebCut Then
        If bTFL Or bTFR Then dFlangeThickness = dFlgThickness
    Else
        If bBFL Or bBFR Then dFlangeThickness = dFlgThickness
    End If
    pPRL.Add "FlangeThickness", dFlangeThickness

    ' define the cutting depth to completely cut the Member cross Section
    Dim dCuttingDepth As Double
    dCuttingDepth = EndCut_GetCutDepth(pPRL)
    pPRL.Add "CutDepth", dCuttingDepth

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
'*****************************************************************************************************
Public Function IJEndCutDrawingType_UpdateDrawingType(ByVal pSmartOccurrence As Object) As String
    On Error GoTo ErrorHandler
    IJEndCutDrawingType_UpdateDrawingType = Sniped_No_Offset_WebCuts
    
    AddFeatureEndCutData pSmartOccurrence, EndCutRelativePosition.TopOrLeft, IJEndCutDrawingType_UpdateDrawingType
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IJEndCutDrawingType_UpdateDrawingType").Number
End Function
'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************
