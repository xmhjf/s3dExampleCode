Attribute VB_Name = "BoundedCustomUtility"
Option Explicit
'-----------------------------------------------------------------------------------------
'  Copyright (C) 2001 - 2016, Intergraph Corporation.  All rights reserved.
'
'
'  File Info:
'      Folder:   SharedContent\Src\StructDetail\Rules\
'      Project:  SDBoundedUSSRule
'      Class:    BoundedCustomUtility
'
'  Abstract:
'       This class module contains the code that:
'       Provides utitlites that are used by the BoundedCustomRule class
'
'  Notes:
'
'
'  History:
'    12/Apr/2012 - svsmylav
'      TR-204570: Added new method 'GetAxisCurveAtCenter' and used this method in 'AdjustExtrusionForBoundingTube'
'        to translate bounded axis if cardinal point is not center of the section.
'      TR-250220: To avoid crash dump,  fix given if oSymbol obj is nothing in GetSymbolParameterValue, then accessing it
'                  from oSymbolOcc passed as an optional argument
'    27/May/2016 - svsmylav TR-295483: positive/negative extrusion distance is computed (instead of 1mm hard-coded value).
'-----------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------'

Private Const MODULE = "SharedContent\Src\StructDetail\Rules\SDBoundedUSSRule\BoundedCustomUtility.bas"
Const TOLERANCE_VALUE = 0.000011

'********************************************************************
' IsBoundedMemberTube
'
'   In:
'   Out:
'********************************************************************
Public Function IsMemberPartTubeType(oCheckObject As Object) As Boolean
Const MT = "IsMemberPartTubeType"
On Error GoTo ErrorHandler
    
    Dim sCStype As String
    
    Dim oport As IJPort
    Dim oMemberObject As Object
    
    Dim oCrossSection As IJCrossSection
    Dim oPartCommon As ISPSMemberPartCommon
    Dim oPartPrismatic As ISPSMemberPartPrismatic
    Dim oSPSCrossSection As ISPSCrossSection
    
    IsMemberPartTubeType = False
    
    ' Check if given Member Part Object is valid
    If TypeOf oCheckObject Is IJPort Then
        Set oport = oCheckObject
        Set oMemberObject = oport.Connectable
    
    ElseIf TypeOf oCheckObject Is ISPSMemberPartCommon Then
        Set oMemberObject = oCheckObject
    
    Else
        Exit Function
    End If
    
    ' Retreive the Member Part Cross Section
    If oMemberObject Is Nothing Then
        Exit Function
    ElseIf TypeOf oMemberObject Is ISPSMemberPartCommon Then
        Set oPartCommon = oMemberObject
        If Not oPartCommon.IsPrismatic Then
            Exit Function
        End If
        
        Set oPartPrismatic = oMemberObject
        Set oSPSCrossSection = oPartPrismatic.crossSection
    Else
        Exit Function
    End If
        
    ' Verify Bounded have valid Cross Section Type
    If oSPSCrossSection Is Nothing Then
        Exit Function
    ElseIf TypeOf oSPSCrossSection.Definition Is IJCrossSection Then
        Set oCrossSection = oSPSCrossSection.Definition
        sCStype = oCrossSection.Type
        
        ' Check for Known Tube Cross Section Types
        If Trim(LCase(sCStype)) = LCase("CS") Then
            IsMemberPartTubeType = True
        ElseIf Trim(LCase(sCStype)) = LCase("HSSC") Then
            IsMemberPartTubeType = True
        ElseIf Trim(LCase(sCStype)) = LCase("PIPE") Then
            IsMemberPartTubeType = True
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Function

'********************************************************************
' IsBoundingPlateLateralPort
'
'   In:
'   Out:
'********************************************************************
Public Function IsBoundingPlateLateralPort(oBoundingObject As Object, _
                                           oConnectable As Object) As Boolean
Const MT = "IsBoundingPlateLateralPort"
On Error GoTo ErrorHandler
    
    Dim sPortType As String
    Dim oport As IJPort
    
    IsBoundingPlateLateralPort = False
    
    ' Check if given Bounding Object is valid
    If TypeOf oBoundingObject Is IJPort Then
        Set oport = oBoundingObject
        Set oConnectable = oport.Connectable
    Else
        Exit Function
    End If
    
    ' Verify that bounding Coonnectable is PlatePart
    If oConnectable Is Nothing Then
        Exit Function
    ElseIf Not TypeOf oConnectable Is IJPlatePart Then
        Exit Function
    End If
        
    ' Verify that bounding Port is Lateral Face Port
    sPortType = Get_PortFaceType(oport)
    If sPortType = C_BaseSide Then
        Exit Function
    ElseIf sPortType = C_OffsetSide Then
        Exit Function
    End If
    
    IsBoundingPlateLateralPort = True

   Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Function

'********************************************************************
' GetMemberTubeRadius
'
'   In:
'   Out:
'********************************************************************
Public Function GetMemberTubeRadius(oMemberObject As ISPSMemberPartCommon) As Double
Const MT = "GetMemberTubeRadius"
On Error GoTo ErrorHandler
    
    Dim dXpnt2 As Double
    Dim dXpnt5 As Double
    Dim dYpnt2 As Double
    Dim dYpnt5 As Double
    
    Dim oSPSCrossSection As ISPSCrossSection
    Dim oMemberPartPrismatic As ISPSMemberPartPrismatic

    GetMemberTubeRadius = 0#
    If oMemberObject Is Nothing Then
        Exit Function
    ElseIf Not TypeOf oMemberObject Is ISPSMemberPartPrismatic Then
        Exit Function
    End If
    
    Set oMemberPartPrismatic = oMemberObject
    Set oSPSCrossSection = oMemberPartPrismatic.crossSection
    oSPSCrossSection.GetCardinalPointOffset 2, dXpnt2, dYpnt2
    oSPSCrossSection.GetCardinalPointOffset 5, dXpnt5, dYpnt5
    GetMemberTubeRadius = Abs(dYpnt5 - dYpnt2)
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Function

'********************************************************************
' GetEndCutCutdepth
'
'   In:
'   Out:
'********************************************************************
Public Function GetEndCutCutdepth(oSymbolDefinition As IJDSymbolDefinition) As Double
Const MT = "GetEndCutCutdepth"
On Error GoTo ErrorHandler

    'Get the grahic inputs to the EndCut symbol
    Dim lFound As Long
    Dim oEnumArg As IJDArgument
    Dim oEnumJDArgument As IEnumJDArgument
    Dim oParameterContent As IJDParameterContent
    
    Dim oSymbol As IJDSymbol
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Dim oSymbolInput As IMSSymbolEntities.IJDInput
    Dim oDefPlayerEx As IMSSymbolEntities.IJDDefinitionPlayerEx
    
    ' initialize the return value for Parameter Inputs
    GetEndCutCutdepth = 0#
    
    Set oInputs = oSymbolDefinition
    Set oDefPlayerEx = oSymbolDefinition
    Set oSymbol = oDefPlayerEx.PlayingSymbol
    
    'Get the enum of arguments set by reference on the symbol if any
    Set oEnumJDArgument = oSymbol.IJDInputsArg.GetInputs(igINPUT_ARGUMENTS_MERGE)
    If Not oEnumJDArgument Is Nothing Then
        Exit Function
    End If
    
    lFound = 1
    oEnumJDArgument.Reset
    While lFound <> 0
        oEnumJDArgument.Next 1, oEnumArg, lFound
        If lFound <> 0 Then
            Set oSymbolInput = oInputs.Item(oEnumArg.Index)
            If oSymbolInput.Properties = igINPUT_IS_A_PARAMETER Then
                If LCase(Trim(oSymbolInput.Name)) = LCase("CutDepth") Then
                    Set oParameterContent = oEnumArg.Entity
                    GetEndCutCutdepth = oParameterContent.UomValue
                    lFound = 0
                End If
            End If
            Set oSymbolInput = Nothing
        End If
    Wend

   Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number

End Function

'********************************************************************
' DoesOutputExists
'
'   In:
'   Out:
'********************************************************************
Public Function DoesOutputExists(oSymbolDefinition As IJDSymbolDefinition, _
                                 sRepName As String, _
                                 sOutput As String) As Boolean
Const MT = "DoesOutputExists"
On Error GoTo ErrorHandler

    Dim lIndex1 As Long
    Dim lIndex2 As Long
    
    Dim sTemp_Output As String
    Dim sTemp_RepName As String
    
    Dim oTemp_Output As IJDOutput
    Dim oTemp_Outputs As IJDOutputs
    Dim oTemp_SymbDefRep As IJDRepresentation
    Dim oTemp_SymbDefReps As IJDRepresentations
    
    DoesOutputExists = False
    
    Set oTemp_SymbDefReps = oSymbolDefinition.IJDRepresentations
'''zMsgBox "oTemp_SymbDefReps.Count: " & oTemp_SymbDefReps.count
    If oTemp_SymbDefReps Is Nothing Then
        Exit Function
    End If
    
    For lIndex1 = 1 To oTemp_SymbDefReps.Count
        Set oTemp_SymbDefRep = oTemp_SymbDefReps.Item(lIndex1)
        sTemp_RepName = oTemp_SymbDefRep.Name
        If Len(Trim(sRepName)) < 1 Then
            Set oTemp_Outputs = oTemp_SymbDefRep
        ElseIf Trim(sRepName) = Trim(sTemp_RepName) Then
            Set oTemp_Outputs = oTemp_SymbDefRep
        Else
            Set oTemp_Outputs = Nothing
        End If

        If Not oTemp_Outputs Is Nothing Then

            For lIndex2 = 1 To oTemp_Outputs.Count
                Set oTemp_Output = oTemp_Outputs.Item(lIndex2)
                sTemp_Output = oTemp_Output.Name
                
                If Trim(sOutput) = Trim(sTemp_Output) Then
                    DoesOutputExists = True
                    If Len(Trim(sRepName)) < 1 Then
                        sRepName = sTemp_RepName
                    End If
                    Exit Function
                End If
            Next lIndex2
        End If
        
    Next lIndex1

   Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number

End Function

'********************************************************************
' TranslateAxisCurve
'   Translates the Axis Curve of the Member Part
'    Firstly from the Current Postion to Center of the Tube
'    Then From Center of the Tube translates the Axis curve to the
'     if bTanslateToTop = true
'         +ve V direction of Sketching plane
'     else
'         to -ve V direction of Sketching plane
'    And then, if The Axis is Curve then after the translation
'    projects the translated Axis Curve on the Sketching plane in the
'    skecthing plane Normal Direction
'
'   In:
'   Out:
'
'********************************************************************
Public Sub TranslateAxisCurve(oBoundedObject As Object, _
                              oMemberPart As Object, _
                              oAxisCurve As Object, _
                              oTransform As DT4x4, _
                              bTanslateToTop As Boolean, _
                              oNewAxisCurve As Object, _
                              oEndCutObject As Object)

    Const MT = "TranslateAxisCurve"
    On Error GoTo ErrorHandler
    Dim dBoundedSize As Double
    
    Dim oTemp_Vvec As IJDVector
    Dim oTranslateMatrix As AutoMath.DT4x4
    
    Dim oCmplx As ComplexString3d
    Dim curveElms As IJElements
    Dim oGeometryFactory As GeometryFactory
    
    Set curveElms = New JObjectCollection
    If TypeOf oAxisCurve Is ComplexString3d Then
        Set oCmplx = oAxisCurve
        oCmplx.GetCurves curveElms
    Else
        curveElms.Add oAxisCurve
    End If

    Set oGeometryFactory = New GeometryFactory
    Set oCmplx = oGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, _
                                                                  curveElms)
    Set curveElms = Nothing

    Dim oConverter As MarineGenericSymbolLib.CSTGeomConverter
    
    Dim oBoundedPort As IJPort
    Dim oBoundedPart As ISPSMemberPartCommon
    Dim oStructPorfile As IJStructProfilePart
    Dim oPos As IJDPosition
    Dim oPoint As IJPoint
    Dim oXSecMatirix As IJDT4x4
    Dim oXSecUvec As IJDVector
    Dim oXSecVvec As IJDVector
    Dim oTranslateMatrixUDir As IJDT4x4
    Dim oTranslateMatrixVDir As IJDT4x4
    Dim oSPSCrossSec As ISPSCrossSection
    Dim dU As Double
    Dim dV As Double
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    Dim dCardinalPoint As Double
    Dim oSPSSplitAxisPort As ISPSSplitAxisPort
    Dim eSPSPortIndex As SPSMemberAxisPortIndex

    '-------------------------------------------------------------------
    'Translates the Axis Curve of the Member Part
    'from the Current Postion to Center of the Tube i.e Cardinal Point 5
    '-------------------------------------------------------------------
    
    If TypeOf oBoundedObject Is IJPort Then
        Set oBoundedPort = oBoundedObject
        Set oBoundedPart = oBoundedPort.Connectable

        If TypeOf oBoundedObject Is ISPSSplitAxisPort Then
             Set oSPSSplitAxisPort = oBoundedObject
             eSPSPortIndex = oSPSSplitAxisPort.PortIndex
        Else
             '(unknown case)need to know the type of End Port
             'so that Strat/End point of Member can be known
             'probably need to throw error here
        End If

        Set oPoint = oBoundedPart.PointAtEnd(eSPSPortIndex)
        oPoint.GetPoint dx, dy, dz
        Set oPos = New DPosition
        oPos.Set dx, dy, dz

        If TypeOf oBoundedPart Is IJStructProfilePart Then
            Set oStructPorfile = oBoundedPart
        Else
           'Unknown case
           'currently need to throw error
        End If

        Set oXSecMatirix = oStructPorfile.GetCrossSectionMatrixAtPoint(oPos)

        Set oXSecUvec = New DVector
        Set oXSecVvec = New DVector

        oXSecUvec.Set oXSecMatirix.IndexValue(0), oXSecMatirix.IndexValue(1), oXSecMatirix.IndexValue(2)
        oXSecVvec.Set oXSecMatirix.IndexValue(4), oXSecMatirix.IndexValue(5), oXSecMatirix.IndexValue(6)

        Set oSPSCrossSec = oBoundedPart.crossSection

        dCardinalPoint = oSPSCrossSec.CardinalPoint
        oSPSCrossSec.GetCardinalPointDelta Nothing, dCardinalPoint, 5, dU, dV

        oXSecUvec.Length = dU
        oXSecVvec.Length = dV

        Set oTranslateMatrixUDir = New DT4x4
        Set oTranslateMatrixVDir = New DT4x4
        oTranslateMatrixUDir.LoadIdentity
        oTranslateMatrixVDir.LoadIdentity

        oTranslateMatrixUDir.Translate oXSecUvec
        oCmplx.Transform oTranslateMatrixUDir

        oTranslateMatrixVDir.Translate oXSecVvec
        oCmplx.Transform oTranslateMatrixVDir
    Else
      ' Need to know the position of the Start/End Port
      ' currently probably need to throw error here
    End If
    
    '-------------------------------------------------------------------
    'From Center of the Tube translates the Axis curve to the Top/Btm
    ' of the Tube
    '-------------------------------------------------------------------
    dBoundedSize = GetMemberTubeRadius(oMemberPart)
    
    Set oTemp_Vvec = New AutoMath.DVector
    oTemp_Vvec.Set oTransform.IndexValue(4), _
                   oTransform.IndexValue(5), _
                   oTransform.IndexValue(6)
    
    If bTanslateToTop Then
        oTemp_Vvec.Length = dBoundedSize
    Else
        oTemp_Vvec.Length = -dBoundedSize
    End If
    
    Set oTranslateMatrix = New AutoMath.DT4x4
    oTranslateMatrix.LoadIdentity
    oTranslateMatrix.Translate oTemp_Vvec
    oCmplx.Transform oTranslateMatrix
    Set oNewAxisCurve = oCmplx
    
    Dim oWireBodyToProject As IJWireBody
    Dim oProjectedWireBody As IJWireBody
    Dim oProjectUtil As IMSModelGeomOps.Project
    Dim oSketchingPlane As IJPlane
    Dim oSketchingPlaneNormal As IJDVector
    
    '-------------------------------------------------------------------
    'if The Axis is Curve(i.e Non Linear) then after the translation
    'project the translated Axis Curve on the Sketching plane in the
    'sketching plane Normal Direction
    '-------------------------------------------------------------------
    If TypeOf oMemberPart Is ISPSMemberPartCurve Then
      'Continue
    Else
      Exit Sub
    End If
    
    Set oSketchingPlaneNormal = New DVector

    oSketchingPlaneNormal.Set oTransform.IndexValue(8), oTransform.IndexValue(9), oTransform.IndexValue(10)
    oSketchingPlaneNormal.Length = 1

    Set oSketchingPlane = New Plane3d
    oSketchingPlane.SetNormal oTransform.IndexValue(8), oTransform.IndexValue(9), oTransform.IndexValue(10)
    oSketchingPlane.SetRootPoint oTransform.IndexValue(12), oTransform.IndexValue(13), oTransform.IndexValue(14)
    oSketchingPlane.SetUDirection oTransform.IndexValue(0), oTransform.IndexValue(1), oTransform.IndexValue(2)

    Dim GeomOpr As IMSModelGeomOps.DGeomWireFrameBody
    Dim oElemCurves As IJElements
    oCmplx.GetCurves oElemCurves

    Set GeomOpr = New DGeomWireFrameBody
    Dim oObj As Object

    Set oObj = GeomOpr.CreateSmartWireBodyFromGTypedCurves(Nothing, oElemCurves)

    Set oWireBodyToProject = oObj
    
    Set oProjectUtil = New Project
    oProjectUtil.CurveAlongVectorOnToSurface Nothing, _
                                             oSketchingPlane, _
                                             oWireBodyToProject, _
                                             oSketchingPlaneNormal, _
                                             Nothing, _
                                             oProjectedWireBody

    Set oNewAxisCurve = oProjectedWireBody

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
    
End Sub

'********************************************************************
' AdjustTubeMatrixForGussetPlate
'
'   Given the case where a Tube is bounded by an Edge of a Gusset (Plate) Part
'   We want the View Transformation Matrix and the Sketching Plane such that
'   the Local U vector is defined along the Bounding Tube axis curve
'   the local V vector is defined by the Base Port of the Gusset (Plate) part
'   the local W vector is defined by the Local U and V
'   This will provide a consistent Sketching Plane such the Bounded Tube can
'   be snipe so that end caps can be applied
'
'   In:
'   Out:
'
'********************************************************************
Public Sub AdjustTubeMatrixForGussetPlate(oEndCutObject As Object, _
                                          oDefinitionPlayerEx As Object, _
                                          oBoundedObject As Object, _
                                          oBoundingObject As Object, _
                                          oBoundingConnectable As Object, _
                                          oConverterData As Object, _
                                          lEndCutType As Long, _
                                          lReturnCode As Long)
Const MT = "AdjustTubeMatrixForGussetPlate"
On Error GoTo ErrorHandler
    
    Dim dDot As Double
    Dim oEndCut_Pos As IJDPosition
    Dim oBounding_Pos As IJDPosition
    
    Dim oTemp_Uvec As IJDVector
    Dim oTemp_Vvec As IJDVector
    Dim oTemp_Wvec As IJDVector
    Dim oView_Uvec As IJDVector
    Dim oView_Vvec As IJDVector
    Dim oView_Wvec As IJDVector
    Dim oBounding_Uvec As IJDVector
    Dim oBounding_Vvec As IJDVector
    Dim oBounding_Wvec As IJDVector
    
    Dim oport As IJPort
    
    Dim oConverter As MarineGenericSymbolLib.CSTGeomConverter
    Dim oViewMatrix As DT4x4
    Dim oCrossMatrix As DT4x4
    Dim oSketchPlane As IJPlane
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    
    lReturnCode = 0

    'Use the Normal vector to set the Transformation matrix local V vector
    Set oConverter = oConverterData
    Set oViewMatrix = oConverter.ViewTransform
    Set oCrossMatrix = oConverter.CrossSectionMatrix
    Set oSketchPlane = oConverter.SketchingPlane

    'set the current View Transform local U,V,W vectors
    Set oView_Uvec = New AutoMath.DVector
    Set oView_Vvec = New AutoMath.DVector
    Set oView_Wvec = New AutoMath.DVector
    
    oView_Uvec.Set oViewMatrix.IndexValue(0), oViewMatrix.IndexValue(1), _
                   oViewMatrix.IndexValue(2)
    oView_Vvec.Set oViewMatrix.IndexValue(4), oViewMatrix.IndexValue(5), _
                   oViewMatrix.IndexValue(6)
    oView_Wvec.Set oViewMatrix.IndexValue(8), oViewMatrix.IndexValue(9), _
                   oViewMatrix.IndexValue(10)
    

    ' Assume Bounding is Gusset Type Plate
    ' Retreive the Gusset Plate Base Port and get its Normal
    Set oEndCut_Pos = New AutoMath.DPosition
    oEndCut_Pos.Set oViewMatrix.IndexValue(12), oViewMatrix.IndexValue(13), _
                    oViewMatrix.IndexValue(14)
    
    ' Get vector from Plate Lateral Face
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    Set oport = oBoundingObject
    oTopologyLocate.GetProjectedPointOnModelBody oport.Geometry, _
                                                 oEndCut_Pos, _
                                                 oBounding_Pos, _
                                                 oBounding_Uvec
    
    ' Get vector from Plate Base Face
    Set oport = SD_GetBasePort(oBoundingConnectable, CTX_BASE)
    oTopologyLocate.GetProjectedPointOnModelBody oport.Geometry, _
                                                 oEndCut_Pos, _
                                                 oBounding_Pos, _
                                                 oBounding_Vvec
    
    ' Set Projection Vector along Plate Lateral Edge
    Set oBounding_Wvec = oBounding_Uvec.Cross(oBounding_Vvec)
    
    ' Adjust the Bounded View Transform Matrix such that it is
    ' in the direction of the Bounding Gusset Plate Base Port
    Set oTemp_Vvec = oBounding_Wvec.Cross(oView_Uvec)
    
    Set oTemp_Wvec = oTemp_Vvec.Cross(oView_Uvec)
    dDot = oView_Wvec.Dot(oTemp_Wvec)
    If dDot < 0# Then
        oTemp_Wvec.Length = -1#
    Else
        oTemp_Wvec.Length = 1#
    End If
    
    Set oTemp_Vvec = oTemp_Wvec.Cross(oView_Uvec)
    dDot = oTemp_Vvec.Dot(oView_Vvec)
    If dDot < 0# Then
        oTemp_Vvec.Length = -1#
    Else
        oTemp_Vvec.Length = 1#
    End If

    dDot = oBounding_Wvec.Dot(oTemp_Wvec)
    If dDot < 0# Then
        oBounding_Wvec.Length = -1#
    Else
        oBounding_Wvec.Length = 1#
    End If

    ' Update the Transformation matrix V and W vectors
    oViewMatrix.IndexValue(4) = oTemp_Vvec.x
    oViewMatrix.IndexValue(5) = oTemp_Vvec.y
    oViewMatrix.IndexValue(6) = oTemp_Vvec.z
    
    oViewMatrix.IndexValue(8) = oTemp_Wvec.x
    oViewMatrix.IndexValue(9) = oTemp_Wvec.y
    oViewMatrix.IndexValue(10) = oTemp_Wvec.z
    oConverter.ViewTransform = oViewMatrix
    
    ' Update the Sketching Plane Normal vector
    oSketchPlane.SetNormal oTemp_Wvec.x, oTemp_Wvec.y, oTemp_Wvec.z
    oConverter.SketchingPlane = oSketchPlane
    
    ' Update the Cross Section matrix V and W vectors
    oCrossMatrix.IndexValue(4) = oBounding_Vvec.x
    oCrossMatrix.IndexValue(5) = oBounding_Vvec.y
    oCrossMatrix.IndexValue(6) = oBounding_Vvec.z
    
    oCrossMatrix.IndexValue(8) = oBounding_Wvec.x
    oCrossMatrix.IndexValue(9) = oBounding_Wvec.y
    oCrossMatrix.IndexValue(10) = oBounding_Wvec.z
    oConverter.CrossSectionMatrix = oCrossMatrix
    
    lReturnCode = 1

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Sub

'********************************************************************
' AdjustExtrusionForBoundingTube
'
'   Given the case where a Tube is bounded by an Tube
'   it is possible that the Bounded Tube will be bounded by multiple Tubes
'   (in an Nodal Connection type configuration)
'   In this case, we do not want the Bounded Tube to be completely cut by each
'   of the Bounding Tubes if the Bounding Tube does not completely bound
'
'   Know have Member Tube Part bounded by a Member Tube Along Axis
'   Based on the Bounded Member Tube Radius ...
'       Check if a point projected along both the Positive and Negative vectors
'       is on the actual Bounding Axis curve
'       if not, change the Extrusion distance to 0.001 for that direction
'
'   In:
'   Out:
'********************************************************************
Public Sub AdjustExtrusionForBoundingTube(oEndCutObject As Object, _
                                          oDefinitionPlayerEx As Object, _
                                          oBoundedObject As Object, _
                                          oBoundingObject As Object, _
                                          oConverterData As Object, _
                                          oGameOutputRep As Object, _
                                          lEndCutType As Long, _
                                          lReturnCode As Long)
Const MT = "AdjustExtrusionForBoundingTube"
On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim sName As String
    Dim sPrefix As String
    
    Dim dx1 As Double
    Dim dX2 As Double
    Dim dy1 As Double
    Dim dY2 As Double
    Dim dz1 As Double
    Dim dZ2 As Double
    Dim dMinDist As Double
    Dim dCutDepth As Double
    Dim dBoundedSize As Double
    Dim dBoundingSize As Double
    
    Dim oTemp_Pos As IJDPosition
    Dim oBound_Pos As IJDPosition
    Dim oEndCut_Pos As IJDPosition
    Dim oPointOnAxis As DPosition

    
    Dim oTemp_Wvec As IJDVector
    Dim oView_Wvec As IJDVector
    
    Dim oport As IJPort
    Dim bIslinear As Boolean

    
    Dim oBoundedAxis As IJCurve
    Dim oBoundedPart As ISPSMemberPartCommon
    
    Dim oBoundingAxis As IJCurve
    Dim oBoundingPart As ISPSMemberPartCommon
    
    Dim oConverter As MarineGenericSymbolLib.CSTGeomConverter
    Dim oSketchPlane As IJPlane
    Dim oParameterContent As IJDParameterContent
    Dim oOutputCollections As IJDOutputCollection
    Dim oGameRepresentation As IJDRepresentationDuringGame
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    Dim oShipGeomOps As GSCADShipGeomOps.SGOWireBodyUtilities
    Dim oPlane As IJPlane
    Dim oHelper As New StructDetailObjects.Helper
    Dim oFinitePlane As Object
    Dim bIsPointOnAxis As Boolean
    
    bIsPointOnAxis = False
    
    lReturnCode = 0
    sMsg = "BoundedSymbol.BoundedCustomUtility::AdjustExtrusionForBoundingTube"
    
    ' get EndCut location and projection direction from the Sketching Plane
    sPrefix = "WebCut:"
    Set oConverter = oConverterData
    Set oSketchPlane = oConverter.SketchingPlane
    oSketchPlane.GetRootPoint dx1, dy1, dz1
    Set oEndCut_Pos = New AutoMath.DPosition
    oEndCut_Pos.Set dx1, dy1, dz1
    
    oSketchPlane.GetNormal dx1, dy1, dz1
    Set oView_Wvec = New AutoMath.DVector
    oView_Wvec.Set dx1, dy1, dz1
    
    ' get Bounded Member Axis Curve and Tube Radius
    Set oport = oBoundedObject
    Set oBoundedPart = oport.Connectable
    Set oBoundedAxis = oBoundedPart.Axis
    dBoundedSize = GetMemberTubeRadius(oBoundedPart)

    ' We could trigger this check ONLY if the utDepth is less tehn radius
    ' or we could trigger the check based on if Bounding Tube axis
    dCutDepth = GetEndCutCutdepth(oDefinitionPlayerEx)

    ' get Bounding Member Axis Curve
    Set oport = oBoundingObject
    Set oBoundingPart = oport.Connectable
    Set oBoundingAxis = oBoundingPart.Axis
    dBoundingSize = GetMemberTubeRadius(oBoundingPart)
    
    ' Caclculate min. Distance between the Axis Curves
    Dim oSPSCrossSec As ISPSCrossSection
    Set oSPSCrossSec = oBoundedPart.crossSection
    If oSPSCrossSec.CardinalPoint <> 5 And oSPSCrossSec.CardinalPoint <> 10 And oSPSCrossSec.CardinalPoint <> 15 Then
        'Cardinal point of bounding is NOT tube center, so no need to transform curve
        Set oBoundedAxis = GetAxisCurveAtCenter(oBoundedPart, oEndCut_Pos)
    End If
    oBoundedAxis.DistanceBetween oBoundingAxis, dMinDist, dx1, dy1, dz1, dX2, dY2, dZ2
    
    Set oBound_Pos = New AutoMath.DPosition
    oBound_Pos.Set dX2, dY2, dZ2
    
    'Save position on bounded curve which closest to the bounding axis; compute its distance from the
    ' sketching plane and use it for extrusion as needed
    Dim oBddPos As IJDPosition
    Set oBddPos = New DPosition
    oBddPos.Set dx1, dy1, dz1
    
    Dim oSketchSurface As IJSurface
    Set oSketchSurface = oSketchPlane
    Dim numPts As Long
    Dim pnts1() As Double
    Dim pnts2() As Double
    Dim pars1() As Double
    Dim pars2() As Double

    oSketchSurface.DistanceBetween oBddPos, dMinDist, dx1, dy1, dz1, dX2, dY2, dZ2, numPts, pnts1, pnts2, pars1, pars2

    ' Check if the Bounding Member extends completely thru the Bounded Member
    ' based on the PositiveExtrusion direction
    Set oGameRepresentation = oGameOutputRep
    Set oOutputCollections = oGameRepresentation.outputCollection
    
    Set oTemp_Wvec = oView_Wvec.Clone
    oTemp_Wvec.Length = dBoundedSize
    Set oTemp_Pos = oBound_Pos.Offset(oTemp_Wvec)
    If Not oBoundingAxis.IsPointOn(oTemp_Pos.x, oTemp_Pos.y, oTemp_Pos.z) Then
        If TypeOf oBoundingAxis Is IJWireBody Then
            Set oShipGeomOps = New GSCADShipGeomOps.SGOWireBodyUtilities
            bIslinear = oShipGeomOps.IsLinear(oBoundingAxis)
            If Not bIslinear Then
                'Curved bounding member
                Set oPlane = New Plane3d
                oPlane.SetRootPoint oTemp_Pos.x, oTemp_Pos.y, oTemp_Pos.z
                oPlane.SetNormal oTemp_Wvec.x, oTemp_Wvec.y, oTemp_Wvec.z
                Set oFinitePlane = oHelper.CreateFinitePlaneFromPlane(oPlane, 4 * dBoundingSize, oTemp_Pos, Nothing)
                Set oTopologyLocate = New TopologyLocate
                Set oPointOnAxis = oTopologyLocate.FindIntersectionPoint(oFinitePlane, oBoundingAxis)
                If Not oPointOnAxis Is Nothing Then
                    bIsPointOnAxis = True
                End If
            End If
        Else
          'need to be hadled for non linear cases,
          'if type of Axis is not Wirebody
        End If
        If Not bIsPointOnAxis Then
            On Error Resume Next
            sName = sPrefix & "PositiveExtrusion"
            Set oParameterContent = oOutputCollections.GetOutput(sName)
            If Not oParameterContent Is Nothing Then
                If dMinDist - 0.001 > TOLERANCE_VALUE Then
                    oParameterContent.UomValue = dMinDist + TOLERANCE_VALUE 'tolerance is added to avoid sliver
                Else
                    oParameterContent.UomValue = 0.001
                End If
                Set oParameterContent = Nothing
                lReturnCode = lReturnCode + 1
                sMsg = sMsg & vbCrLf & _
                       sName & " ... oParameterContent.UomValue = 0.001"
            End If
        End If
    Else
        sMsg = sMsg & vbCrLf & _
               sName & " ... Test point is on Bounding Axis"
    
    End If
    
    ' Check if the Bounding Member extends completely thru the Bounded Member
    ' based on the NegativeExtrusion direction
    Set oTemp_Wvec = oView_Wvec.Clone
    oTemp_Wvec.Length = -dBoundedSize
    Set oTemp_Pos = oBound_Pos.Offset(oTemp_Wvec)
    If Not oBoundingAxis.IsPointOn(oTemp_Pos.x, oTemp_Pos.y, oTemp_Pos.z) Then
        If TypeOf oBoundingAxis Is IJWireBody Then
              Set oShipGeomOps = New GSCADShipGeomOps.SGOWireBodyUtilities
              bIslinear = oShipGeomOps.IsLinear(oBoundingAxis)
              If Not bIslinear Then
                Set oPlane = New Plane3d
                oPlane.SetRootPoint oTemp_Pos.x, oTemp_Pos.y, oTemp_Pos.z
                oPlane.SetNormal oTemp_Wvec.x, oTemp_Wvec.y, oTemp_Wvec.z
                Set oFinitePlane = oHelper.CreateFinitePlaneFromPlane(oPlane, 4 * dBoundingSize, oTemp_Pos, Nothing)
                Set oTopologyLocate = New TopologyLocate
                Set oPointOnAxis = oTopologyLocate.FindIntersectionPoint(oFinitePlane, oBoundingAxis)

                If Not oPointOnAxis Is Nothing Then
                   Exit Sub
                End If
              End If
        Else
          'need to be hadled for non linear cases,
          'if type of Axis is not Wirebody
        End If
        If Not bIsPointOnAxis Then
            On Error Resume Next
            sName = sPrefix & "NegativeExtrusion"
            Set oParameterContent = oOutputCollections.GetOutput(sName)
            If Not oParameterContent Is Nothing Then
                If dMinDist - 0.001 > TOLERANCE_VALUE Then
                    oParameterContent.UomValue = dMinDist + TOLERANCE_VALUE 'tolerance is added to avoid sliver
                Else
                    oParameterContent.UomValue = 0.001
                End If
                Set oParameterContent = Nothing
                lReturnCode = lReturnCode + 2
                sMsg = sMsg & vbCrLf & _
                       sName & " ... oParameterContent.UomValue = 0.001"
            End If
        End If
    Else
        sMsg = sMsg & vbCrLf & _
               sName & " ... Test point is on Bounding Axis"
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
    
End Sub

'**********************************************************************************************
' Method      : AdjustSkecthingPlaneForBoundedTube
' Description : Sets the Sketching plane
'               This method should be used only when the Tube is Bounded and Non-Tubular
'               members(Standard) is Bounding and when Bounding Port is Lateral Port
'
'**********************************************************************************************
Public Sub AdjustSketchingPlaneForBoundedTube(oEndCutObject As Object, _
                                          oDefinitionPlayerEx As Object, _
                                          oBoundedObject As Object, _
                                          oBoundingObject As Object, _
                                          oBoundingConnectable As Object, _
                                          oConverterData As Object, _
                                          lEndCutType As Long, _
                                          lReturnCode As Long)

    Const MT = "AdjustSkecthingPlaneForBoundedTube"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim oTubeCustomRule As Object

    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim oBoundingProfilePart As IJStructProfilePart
    Dim oOrigin As AutoMath.DPosition
    Dim oSketchPlane As IJPlane
    Dim oBoundingCrossSectionMatrix As IJDT4x4
    Dim oTubeSketchPlaneMatrix As IJDT4x4
    Dim oStructConverter As MarineGenericSymbolLib.CSTGeomConverter
    
    lReturnCode = 0
    
    '------------------------------------------------------
    ' Verify bounded object is a Profile/Member part port
    '------------------------------------------------------
    If TypeOf oBoundedObject Is IJPort Then
        Set oBoundedPort = oBoundedObject
    Else
        Exit Sub
    End If
    
    '------------------------------------------------------
    ' Verify/Retrieve that the Bounding Port is valid
    '------------------------------------------------------
    If TypeOf oBoundingObject Is IJPort Then
        Set oBoundingPort = oBoundingObject
        If TypeOf oBoundingPort.Connectable Is IJStructProfilePart Then
            Set oBoundingProfilePart = oBoundingPort.Connectable
        Else
            'Need to Handle if Bounding is Plate soon
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    '------------------------------------------------------
    'Get the Sketching Plane when Tubular Member is Bounded
    '------------------------------------------------------
    Set oSketchPlane = GetSketchPlaneForTube(oBoundingPort, oBoundedPort)
    
    '------------------------------------------------------
    'Get View Transform Matrix from the obtained Sketch Plane
    '------------------------------------------------------
    Set oTubeSketchPlaneMatrix = ConstructMatrixFromPlane(oSketchPlane)
    
    '------------------------------------------------------
    'Get the Root Point of the Sktech Plane
    '------------------------------------------------------
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    
    oSketchPlane.GetRootPoint dx, dy, dz
    Set oOrigin = New AutoMath.DPosition
    oOrigin.Set dx, dy, dz
    
    '------------------------------------------------------
    'Get the Bounding Cross Section Matrix
    '------------------------------------------------------
    If Not oBoundingProfilePart Is Nothing Then
        Set oBoundingCrossSectionMatrix = oBoundingProfilePart.GetCrossSectionMatrixAtPoint(oOrigin)
    Else
        'Need to Handle if Bounding is Plate
        Exit Sub
    End If
    
    Set oStructConverter = oConverterData
    
    oStructConverter.SketchingPlane = oSketchPlane
    oStructConverter.ViewTransform = oTubeSketchPlaneMatrix
    oStructConverter.CrossSectionMatrix = oBoundingCrossSectionMatrix

    lReturnCode = 1
    Set oTubeCustomRule = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
    
End Sub

'**********************************************************************************************
' Method      : ConstructMatrixFromPlane
' Description : Helper function to retrieve the Transformation Matrix given an IJPlane object
'
'**********************************************************************************************
Public Function ConstructMatrixFromPlane(oPlane As IJPlane) As AutoMath.IJDT4x4
    Const MT = "ConstructMatrixFromPlane"
    On Error GoTo ErrorHandler
    
    Dim dMatrix(16) As Double
    Dim oMatrix As AutoMath.IJDT4x4
    
    Set oMatrix = New AutoMath.DT4x4
    oMatrix.LoadIdentity
    oMatrix.Get dMatrix(0)

    oPlane.GetUDirection dMatrix(0), dMatrix(1), dMatrix(2)
    oPlane.GetVDirection dMatrix(4), dMatrix(5), dMatrix(6)
    oPlane.GetNormal dMatrix(8), dMatrix(9), dMatrix(10)
    oPlane.GetRootPoint dMatrix(12), dMatrix(13), dMatrix(14)

    oMatrix.Set dMatrix(0)
    Set ConstructMatrixFromPlane = oMatrix

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT, "").Number
   
End Function

'**********************************************************************************************
' Method      : EdgeMappingRuleProgID
' Description : This method tries to retrieve the End Cut Edge mapping Rule ProgID
'       from Catalog(if Bulkloaded).
'       If its not bulkloaded we just hard code the ProgID and return it as an output
'
'**********************************************************************************************
Public Function EdgeMappingRuleProgID() As String
    Const METHOD = "EdgeMappingRuleProgID"
    On Error GoTo ErrorHandler

    Dim oCatalogQuery As IJSRDQuery
    Dim oRule As IJSRDRule
    Dim sClassName As String
    Dim sAssemblyName As String
    Dim oRuleQuery As IJSRDRuleQuery
     
    Set oCatalogQuery = New SRDQuery
    Set oRuleQuery = oCatalogQuery.GetRulesQuery
    
    ' Check if a Rule Has been bulkloaded into the Catalog
    Dim oRuleUnk As Object
    On Error Resume Next
    Set oRuleUnk = oRuleQuery.GetRule("EndCutBoundingMapRule")
    
    If Not oRuleUnk Is Nothing Then
       ' EndCut Mapping Rule was found
       Set oRule = oRuleUnk
       
       EdgeMappingRuleProgID = oRule.ProgId
    End If
    
    Set oCatalogQuery = Nothing
    Set oRuleQuery = Nothing
    Set oRule = Nothing
    Set oRuleUnk = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
   
End Function

'**********************************************************************************************
' Method      : CreateSymbolInstance
' Description : Creates symbol instance for the passed in ProID
'
'**********************************************************************************************
Public Function CreateSymbolInstance(strProgId As String) As Object
    Const METHOD = "CreateSymbolInstance"
    On Error Resume Next
    
    Dim strCodeBase As String
    strCodeBase = Null
    
    Dim oCreateInstanceHelper As New CreateInstanceHelper
    Set CreateSymbolInstance = oCreateInstanceHelper.CreateInstance(strProgId, strCodeBase)
    Set oCreateInstanceHelper = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
    
End Function

'**********************************************************************************************
' Method      : RotateSketchingPlaneForFreeEndCutOnTube
' Description : Rotates the Sketching plane as per parameter defined on Symbol
'               This method should be used only when the Free End Cut is applied on Tube
'               End User is given the flexibility to roatate the Sketching plane
'               so it can orient the EndCut on Tubes per needs
'               "SketchPlaneRotation" parameter is exposed to the user through parameter
'               tab on PropertyPage
'**********************************************************************************************
Public Sub RotateSketchingPlaneForFreeEndCutOnTube(oEndCutObject As Object, _
                                                   oDefinitionPlayerEx As Object, _
                                                   oBoundedObject As Object, _
                                                   oBoundingObject As Object, _
                                                   oBoundingConnectable As Object, _
                                                   oConverterData As Object, _
                                                   lEndCutType As Long, _
                                                   lReturnCode As Long)
                                          
    Const METHOD = "CreateSymbolInstance"
    On Error GoTo ErrorHandler
    
    Dim RotateAxis As IJDVector
    Dim oStructConverter As MarineGenericSymbolLib.CSTGeomConverter
    Dim oRotationalMatrix As AutoMath.DT4x4
    Dim oTransform As DT4x4
    Dim dRotate As Double
    Dim oSketchPlane As IJPlane
    Dim oNewTubeSketchPlane As IJPlane
    Dim oNewTubeSketchPlaneMatrix As IJDT4x4
    Dim dNormalX As Double
    Dim dNormalY As Double
    Dim dNormalZ As Double
    Dim dUx As Double
    Dim dUy As Double
    Dim dUz As Double
    
    lReturnCode = 0
    
    Set oStructConverter = oConverterData
    Set oTransform = New DT4x4
    Set oTransform = oStructConverter.ViewTransform
    
    Set RotateAxis = New DVector
    RotateAxis.Set oTransform.IndexValue(0), oTransform.IndexValue(1), oTransform.IndexValue(2)
    
    dRotate = GetSymbolParameterValue(oDefinitionPlayerEx, "SketchPlaneRotation", oEndCutObject)
    
    Set oRotationalMatrix = New AutoMath.DT4x4
    oRotationalMatrix.LoadIdentity
    oRotationalMatrix.Rotate dRotate, RotateAxis
    
    Set oSketchPlane = oStructConverter.SketchingPlane
    oSketchPlane.GetUDirection dUx, dUy, dUz
    oSketchPlane.Transform oRotationalMatrix
    oSketchPlane.GetNormal dNormalX, dNormalY, dNormalZ
    
    Set oNewTubeSketchPlane = New Plane3d
    
    oNewTubeSketchPlane.SetNormal dNormalX, dNormalY, dNormalZ
    oNewTubeSketchPlane.SetRootPoint oTransform.IndexValue(12), oTransform.IndexValue(13), oTransform.IndexValue(14)
    oNewTubeSketchPlane.SetUDirection dUx, dUy, dUz
    
    Set oNewTubeSketchPlaneMatrix = ConstructMatrixFromPlane(oNewTubeSketchPlane)
    
    oStructConverter.SketchingPlane = oNewTubeSketchPlane
    oStructConverter.ViewTransform = oNewTubeSketchPlaneMatrix
        
    lReturnCode = 1
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    
End Sub
'********************************************************************
' GetSymbolParameterValue
'
'   In:
'   Out:
'********************************************************************
Public Function GetSymbolParameterValue(oSymbolDefinition As IJDSymbolDefinition, sParameterName As String, Optional oSymbolOcc As Object = Nothing) As Double
Const MT = "GetSymbolParameterValue"
On Error GoTo ErrorHandler

    'Get the grahic inputs to the EndCut symbol
    Dim lFound As Long
    Dim oEnumArg As IJDArgument
    Dim oEnumJDArgument As IEnumJDArgument
    Dim oParameterContent As IJDParameterContent
    
    Dim oSymbol As IJDSymbol
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Dim oSymbolInput As IMSSymbolEntities.IJDInput
    Dim oDefPlayerEx As IMSSymbolEntities.IJDDefinitionPlayerEx
    
    ' initialize the return value for Parameter Inputs
    GetSymbolParameterValue = 0#
    
    Set oInputs = oSymbolDefinition
    Set oDefPlayerEx = oSymbolDefinition
    
    'Check if the Optional argument passed is Nothing or not
    If Not oSymbolOcc Is Nothing Then
        Set oSymbol = oSymbolOcc
    Else
        Set oSymbol = oDefPlayerEx.PlayingSymbol
    End If
    
    'Check if oSymbol exist only then we can access any properties on it
    If oSymbol Is Nothing Then
        Exit Function
    End If
    
    'Get the enum of arguments set by reference on the symbol if any
    Set oEnumJDArgument = oSymbol.IJDInputsArg.GetInputs(igINPUT_ARGUMENTS_MERGE)
    If Not oEnumJDArgument Is Nothing Then
        'Exit Function
    End If
    
    lFound = 1
    oEnumJDArgument.Reset
    While lFound <> 0
        oEnumJDArgument.Next 1, oEnumArg, lFound
        If lFound <> 0 Then
            Set oSymbolInput = oInputs.Item(oEnumArg.Index)
            If oSymbolInput.Properties = igINPUT_IS_A_PARAMETER Then
                If LCase(Trim(oSymbolInput.Name)) = LCase(sParameterName) Then
                    Set oParameterContent = oEnumArg.Entity
                    GetSymbolParameterValue = oParameterContent.UomValue
                    lFound = 0
                End If
            End If
            Set oSymbolInput = Nothing
        End If
    Wend

   Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
    
End Function

'********************************************************************
' IsFreeEndCut
'
'   In:
'   Out:
'********************************************************************
Public Function IsFreeEndCut(oEndCutObject As Object) As Boolean
    Const MT = "IsFreeEndCut"
    On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    
    IsFreeEndCut = False
    
    If TypeOf oEndCutObject Is IJSmartOccurrence Then
       Set oSmartOcc = oEndCutObject
       Set oSDO_WebCut = New StructDetailObjects.WebCut
       Set oSDO_WebCut.object = oSmartOcc
    Else
       'need to handle such cases(right now unaware of such cases)
       Exit Function
    End If
   
    IsFreeEndCut = oSDO_WebCut.IsFreeEndCut
     
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
    
End Function

'***********************************************************************************************
'    Function      : GetAxisCurveAtCenter
'
'    Description   : Given a member part and end cut position returns axis curve transformed to
'                    the center.
'
'    Parameters    :
'          Input    Member part, EndCut position
'
'    Return        : Curve at Center as IJCurve
'
'***********************************************************************************************
Public Function GetAxisCurveAtCenter(oMemberPart As ISPSMemberPartCommon, oEndCutPos As IJDPosition) As IJCurve
    Const METHOD = "GetAxisCurveAtCenter"
    Dim sMsg As String
    
    'Declare variables
    Dim oMatrix As DT4x4
    Dim oXSecUvec As IJDVector
    Dim oXSecVvec As IJDVector
    
    'Get cardinal point information

    Dim oSPSCrossSec As ISPSCrossSection
    Dim lCardinalPoint As Long
    
    Set oSPSCrossSec = oMemberPart.crossSection
    lCardinalPoint = oSPSCrossSec.CardinalPoint

    'Prepare complex string
    Dim oCmplx As ComplexString3d
    Dim curveElms As IJElements
    Set curveElms = New JObjectCollection
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    
    If TypeOf oMemberPart.Axis Is ComplexString3d Then
        Set oCmplx = oMemberPart.Axis
        oCmplx.GetCurves curveElms
    Else
        curveElms.Add oMemberPart.Axis
    End If
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory
    Set oCmplx = oGeometryFactory.ComplexStrings3d.CreateByCurves(Nothing, curveElms)
    If oSPSCrossSec.CardinalPoint = 5 Or oSPSCrossSec.CardinalPoint = 10 Or oSPSCrossSec.CardinalPoint = 15 Then
        'Cardinal point is tube center, so no need to transform curve
    Else
        'Need to transform the curve to center of tube
        Dim dU As Double
        Dim dV As Double
        oSPSCrossSec.GetCardinalPointDelta Nothing, oSPSCrossSec.CardinalPoint, 5, dU, dV
        
        Dim oStructPorfile As IJStructProfilePart
        If TypeOf oMemberPart Is IJStructProfilePart Then
            Set oStructPorfile = oMemberPart
        Else
            sMsg = "Unknown case" 'Unknown case
            GoTo ErrorHandler
        End If

        Set oMatrix = oStructPorfile.GetCrossSectionMatrixAtPoint(oEndCutPos)

        Set oXSecUvec = New DVector
        Set oXSecVvec = New DVector

        oXSecUvec.Set oMatrix.IndexValue(0), oMatrix.IndexValue(1), oMatrix.IndexValue(2)
        oXSecVvec.Set oMatrix.IndexValue(4), oMatrix.IndexValue(5), oMatrix.IndexValue(6)
        
        'Compute U-vector
        oXSecUvec.Length = dU
    
        'Compute V-vector
        oXSecVvec.Length = dV
                
        'Compute resultant vector
        Dim oResVec As IJDVector
        Set oResVec = oXSecUvec.Add(oXSecVvec)
        
        'Prepare transformation matrix
        oMatrix.LoadIdentity
        oMatrix.Translate oResVec
                
        'Tranform complex string to tube-center
        oCmplx.Transform oMatrix
    End If

    Set GetAxisCurveAtCenter = oCmplx
    
    'Cleanup
    Set oGeometryFactory = Nothing
    Set oMemberPart = Nothing
    Set oSPSCrossSec = Nothing
    Set curveElms = Nothing
    Set oMatrix = Nothing
    Set oXSecUvec = Nothing
    Set oXSecVvec = Nothing

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

Public Function AdjustFlgCuttingDirection(oEndCutObject As Object, _
                                                ByRef oConverterData As Object, lEndCutType As Long, _
                                                oDefinitionPlayerEx As Object, Optional lReturnCode As Long = 0)


Const MT = "_AdjustFlgCuttingDirection"
On Error GoTo ErrorHandler
    lReturnCode = 0
    Dim oConverter As MarineGenericSymbolLib.CSTGeomConverter
    Set oConverter = oConverterData
    Dim oCrSecMatrix As New DT4x4
    Dim oFlangeCutPoints As Object
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object

    Dim iCount As Integer
    Dim IsTopFlangeCut As Boolean
    Dim oCommonGeom  As IMSModelGeomOps.DGeomOpsIntersect
    Set oCommonGeom = New IMSModelGeomOps.DGeomOpsIntersect
    

    Get_BoundedSymbolGrahicInputs oDefinitionPlayerEx, _
                                   oBoundedObject, _
                                  oBoundingObject, _
                                  oFlangeCutPoints

            Set oCrSecMatrix = oConverter.CrossSectionMatrix
         
                        
            Dim oFlgCut As StructDetailObjects.FlangeCut
            Dim oBndingTubeSurface As IJSurfaceBody
            Dim oBndedFlgSurface As IJSurfaceBody

            Dim oBoundingMemberPart As StructDetailObjects.MemberPart
            Set oBoundingMemberPart = New MemberPart
            Dim oBoundingPort As IJPort
            Dim oBoundedPort As IJPort
            Set oBoundingPort = oBoundingObject
            Set oBoundedPort = oBoundedObject
            Set oBoundingMemberPart.object = oBoundingPort.Connectable
            Set oBndingTubeSurface = oBoundingMemberPart.SubPort(JXSEC_OUTER_TUBE).Geometry

            Dim oPlane As IJPlane
            Dim oSurface As IJSurface
            Dim oBoundedMemberPart As StructDetailObjects.MemberPart
            Set oBoundedMemberPart = New StructDetailObjects.MemberPart
            Set oBoundedMemberPart.object = oBoundedPort.Connectable
            Dim oExtendedPort1 As IJSurfaceBody
            Dim oExtendedPort2 As IJSurfaceBody
            Dim IsTopPortIntr As Boolean
            Dim oGeomOperation As GSCADShipGeomOps.SGOModelBodyUtilities
            Set oGeomOperation = New SGOModelBodyUtilities
            Set oFlgCut = New FlangeCut
            Set oFlgCut.object = oEndCutObject
            IsTopFlangeCut = oFlgCut.IsTopFlange
            
            Dim IsBtmPortIntr As Boolean
            If oFlgCut.IsTopFlange Then
                ' Check whether the Bounded top port intersects bounding outer port
                Set oExtendedPort1 = GetExtendedPort(oBoundedMemberPart.SubPort(JXSEC_TOP))
                IsTopPortIntr = oGeomOperation.HasIntersectingGeometry(oExtendedPort1, oBoundingMemberPart.SubPort(JXSEC_OUTER_TUBE))
                If Not IsTopPortIntr Then
                    Exit Function
                End If
                
            Else
                'Check whether the Bounded bottom port intersects the bounding outer port
                Set oExtendedPort2 = GetExtendedPort(oBoundedMemberPart.SubPort(JXSEC_BOTTOM))
                IsBtmPortIntr = oGeomOperation.HasIntersectingGeometry(oExtendedPort2, oBoundingMemberPart.SubPort(JXSEC_OUTER_TUBE))
                If Not IsBtmPortIntr Then
                    Exit Function
                End If
            End If
            
            '*******************************************************
            'need to create a flange plane
            Dim oPlane1 As IJPlane
            Dim pTopoffsetSurface As IJSurfaceBody
            Dim pBtmoffsetSurface As IJSurfaceBody
            Dim oGeomOpr As GSCADShipGeomOps.SGOSurfaceBodyUtilities
            Set oGeomOpr = New SGOSurfaceBodyUtilities
            
            'creating the offset surface for Top and bottom
            Dim dhalfFlgThick As Double
            Dim dFlgThickness As Double
            dFlgThickness = oBoundedMemberPart.FlangeThickness
            dhalfFlgThick = dFlgThickness / 2

            Dim oBtmport As IJPort
            Dim oTopport As IJPort
            Set oBtmport = GetLateralSubPortBeforeTrim(oBoundedMemberPart.object, JXSEC_BOTTOM)
            Set oTopport = GetLateralSubPortBeforeTrim(oBoundedMemberPart.object, JXSEC_TOP)
            If oFlgCut.IsTopFlange Then
                oGeomOpr.CreateOffsetSurface oTopport.Geometry, -dhalfFlgThick, False, pTopoffsetSurface
            Else
                oGeomOpr.CreateOffsetSurface oBtmport.Geometry, -dhalfFlgThick, False, pBtmoffsetSurface
            End If
                
            Dim oOutputExtSB As Object
            'Getting Extended Sheet Body for WR
            
            Dim oExtWRWireTop As Object
            Dim oExtWRWireBtm As Object

            Dim oGeometryOffset As IJGeometryOffset
            Set oGeometryOffset = New DGeomOpsOffset
            'Create extended web-right plane with 1m for each side
            Dim oWRport As IJPort
            
            
            Set oWRport = GetLateralSubPortBeforeTrim(oBoundedMemberPart.object, JXSEC_WEB_RIGHT)
            oGeometryOffset.CreateExtendedSheetBody Nothing, oWRport.Geometry, Nothing, 0.5, Nothing, oOutputExtSB
     
            
            If oFlgCut.IsTopFlange Then
                ' Get the intersection object when top offset-surface and extended Web-right
                oCommonGeom.PlaceIntersectionObject Nothing, pTopoffsetSurface, oOutputExtSB, Nothing, oExtWRWireTop
                If oExtWRWireTop Is Nothing Then
                    Exit Function
                End If
            Else
                oCommonGeom.PlaceIntersectionObject Nothing, pBtmoffsetSurface, oOutputExtSB, Nothing, oExtWRWireBtm
                If oExtWRWireBtm Is Nothing Then
                    Exit Function
                End If
            End If
            

            Dim oTopExtWireBody As IJWireBody
            Dim oBtmExtWireBody As IJWireBody
            Dim outility As New IMSModelGeomOps.DGeomWireFrameBody
            
            If oFlgCut.IsTopFlange Then
                'Extend the wire body for finding points of intersection on bounding by projecting
                Set oTopExtWireBody = outility.CreateExtendedWire(Nothing, oExtWRWireTop, 1, 1, True, True, Nothing)
            Else
                Set oBtmExtWireBody = outility.CreateExtendedWire(Nothing, oExtWRWireBtm, 1, 1, True, True, Nothing)
            End If
            
            'Get the nearest intersection point from bounded

            Dim oSgoWirebodyUtil As New SGOWireBodyUtilities

            Dim oPointsGraph As IJPointsGraphBody

            ' Checking whether the passed ports has intersecting geometry
            Dim IsIntersectingGeomTop As Boolean
            Dim IsIntersectingGeomBtm As Boolean

            If oFlgCut.IsTopFlange Then
                IsIntersectingGeomTop = oGeomOperation.HasIntersectingGeometry(oTopExtWireBody, oBoundingMemberPart.SubPort(JXSEC_OUTER_TUBE))
                    If IsIntersectingGeomTop Then
                        Set oPointsGraph = oSgoWirebodyUtil.FindIntersectionOfWireAndOffsetSurface(oTopExtWireBody, oBoundingMemberPart.SubPort(JXSEC_OUTER_TUBE), 0)
                    Else
                        Exit Function
                    End If
            Else
                IsIntersectingGeomBtm = oGeomOperation.HasIntersectingGeometry(oBtmExtWireBody, oBoundingMemberPart.SubPort(JXSEC_OUTER_TUBE))
                    If IsIntersectingGeomBtm Then
                        Set oPointsGraph = oSgoWirebodyUtil.FindIntersectionOfWireAndOffsetSurface(oBtmExtWireBody, oBoundingMemberPart.SubPort(JXSEC_OUTER_TUBE), 0)
                    Else
                        Exit Function
                    End If
            End If
                      
            Dim oPointsGraphUtility As New SGOPointsGraphUtilities
            Dim oCollection As New Collection
            
            Set oCollection = oPointsGraphUtility.GetPositionsFromPointsGraph(oPointsGraph)

            Dim oPos1 As New DPosition
            Dim oPos2 As New DPosition
            Dim oFlangeUDir As IJDVector
            Set oFlangeUDir = New DVector
            Dim oVctr As IJDVector
            Dim oFinalPos As IJDPosition
            Set oFinalPos = New DPosition
            
            ' Getting the collection of points
            If oCollection.Count = 0 Then
                ' oCollection is empty
                Exit Function

            ElseIf oCollection.Count = 1 Then
                ' If the count is 1 then set that point to oFinalpos
                Set oPos1 = oCollection.Item(1)
                oFinalPos.Set oPos1.x, oPos1.y, oPos1.z

            ElseIf oCollection.Count > 1 Then
                ' If the count is greater than 1
                Set oPos1 = oCollection.Item(1)
                Set oPos2 = oCollection.Item(2)
                Set oVctr = New DVector
                oVctr.Set oPos2.x - oPos1.x, oPos2.y - oPos1.y, oPos2.z - oPos1.z
                oVctr.Length = 1
     
            oFlangeUDir.Set oConverter.ViewTransform.IndexValue(0), _
                            oConverter.ViewTransform.IndexValue(1), _
                            oConverter.ViewTransform.IndexValue(2)
            
                If oFlangeUDir.Dot(oVctr) >= 0 Then
                    oFinalPos.Set oPos2.x, oPos2.y, oPos2.z
                    
                Else
                   oFinalPos.Set oPos1.x, oPos1.y, oPos1.z
                End If
                
            End If

            Dim oSurfaceBody As IJSurfaceBody
            Set oSurfaceBody = oBndingTubeSurface

            Dim oSurfaceNormal As IJDVector
            oSurfaceBody.GetNormalFromPosition oFinalPos, oSurfaceNormal

            
            Dim oFlangeVDir As IJDVector
            Set oFlangeVDir = New DVector
            
            Dim oTangentailDir As IJDVector
            
            oSurfaceNormal.Length = 1

            oFlangeVDir.Set oConverter.ViewTransform.IndexValue(4), _
                            oConverter.ViewTransform.IndexValue(5), _
                            oConverter.ViewTransform.IndexValue(6)
            oFlangeVDir.Length = 1

            Set oTangentailDir = oSurfaceNormal.Cross(oFlangeVDir)
            If oFlgCut.IsTopFlange Then
                
                If IsTopPortIntr Then
                    oConverter.CrossSectionMatrix.IndexValue(0) = oSurfaceNormal.x
                    oConverter.CrossSectionMatrix.IndexValue(1) = oSurfaceNormal.y
                    oConverter.CrossSectionMatrix.IndexValue(2) = oSurfaceNormal.z
                    oConverter.CrossSectionMatrix.IndexValue(8) = oTangentailDir.x
                    oConverter.CrossSectionMatrix.IndexValue(9) = oTangentailDir.y
                    oConverter.CrossSectionMatrix.IndexValue(10) = oTangentailDir.z
                
                End If
            Else
                
                If IsBtmPortIntr Then
                    oConverter.CrossSectionMatrix.IndexValue(0) = oSurfaceNormal.x
                    oConverter.CrossSectionMatrix.IndexValue(1) = oSurfaceNormal.y
                    oConverter.CrossSectionMatrix.IndexValue(2) = oSurfaceNormal.z
                    oConverter.CrossSectionMatrix.IndexValue(8) = oTangentailDir.x
                    oConverter.CrossSectionMatrix.IndexValue(9) = oTangentailDir.y
                    oConverter.CrossSectionMatrix.IndexValue(10) = oTangentailDir.z
                End If
            End If
            
    
Exit Function



ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Function


'*************************************************************************
'Function
'HandleError
'
'Abstract
' called by other subs and fuctions during error. This method logs the error
' and returns success
'
'input
'Module(file) name, method name
'
'Return
'success
'
'Exceptions
'
'***************************************************************************
Public Sub HandleError(sModule As String, sMETHOD As String, Optional sExtraInfo As String = "")
    Dim oEditErrors As IJEditErrors
    
    Set oEditErrors = New JServerErrors
    If Not oEditErrors Is Nothing Then
        oEditErrors.AddFromErr Err, sExtraInfo, sMETHOD, sModule
    End If
    Set oEditErrors = Nothing
End Sub

'**********************************************************************************************
' Method      : IsWebPenetrated
' Description : Checks whether the Member is web penerated/ flange penetrated
'               This method Checks whether the test case is web penerated or flange penetrated
'               by returnig a boolean value. If the boolean value is true it is web penetrated case
'               or else it is flange penetrated case
'
'**********************************************************************************************
Public Function IsWebPenetrated(oBoundingPort As IJPort, oBoundedPort As IJPort) As Boolean
    Const MT = "IsWebPenetrated"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    Dim oBoundedPart As Object
    Dim oBoundedMemberPart As ISPSMemberPartCommon
    Dim oBoundingPart As IJStructProfilePart
    Dim oBoundedStart As IJPoint
    Dim oBoundedEnd As IJPoint
    Dim oBoundedPos As IJDPosition
    Dim dStartX As Double
    Dim dStartY As Double
    Dim dStartZ As Double
    Dim dEndX As Double
    Dim dEndY As Double
    Dim dEndZ As Double
    Dim oBoundedMemberPort As ISPSSplitAxisPort
    Dim oBoundingXSectionMatrix As IJDT4x4
    Dim oBoundedXSectionMatrix As IJDT4x4
    Dim oDummyMatrix As IJDT4x4
    Dim oBoundedStructPort As IJStructPort
    Dim bIsBuiltup As Boolean, oBUMember As ISPSDesignedMember
    
    ' by default web is penetrating
    IsWebPenetrated = True
    
    Set oBoundedPart = oBoundedPort.Connectable
    IsFromBuiltUpMember oBoundedPart, bIsBuiltup, oBUMember
    
    If bIsBuiltup Then
        Set oBoundedPart = oBUMember
    End If
    
    If Not oBoundedPart Is Nothing Then
        If TypeOf oBoundedPart Is IJStructProfilePart Then
            ' Valid case
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    If Not TypeOf oBoundingPort Is IJPort Or oBoundingPort Is Nothing Or oBoundingPort.Connectable Is Nothing Then
        Exit Function
    End If
    
    ' if Bounded is Tube and Bounding is Not Tube, consider as WebPenetrating
    If IsTubularMember(oBoundedPort) And Not IsTubularMember(oBoundingPort) Then
        IsWebPenetrated = True
        Exit Function
    End If
    
    ' if v-vector is more alligned with the bounding x-axis, must be penetrating flange
    Dim oBoundedU As IJDVector
    Dim oBoundedV As IJDVector
    Dim oBoundingW As IJDVector
    Dim oBdedLandingCurve As IJWireBody
    Dim posStart As IJDPosition
    Dim posEnd As IJDPosition
    
    sMsg = "Getting the bounded Position"
    
    Set oBoundedU = New DVector
    Set oBoundedV = New DVector
    Set oBoundingW = New DVector
    Set oBoundedPos = New DPosition
    
    ' Step 1: compute vectors for bounded
    If TypeOf oBoundedPart Is ISPSMemberPartCommon Then
        Set oBoundedMemberPart = oBoundedPart
        Set oBoundedStart = oBoundedMemberPart.PointAtEnd(SPSMemberAxisStart)
        Set oBoundedEnd = oBoundedMemberPart.PointAtEnd(SPSMemberAxisEnd)
        
        If TypeOf oBoundedPort Is ISPSSplitAxisPort Then
            Set oBoundedMemberPort = oBoundedPort
        If oBoundedMemberPort.PortIndex = SPSMemberAxisStart Then
            oBoundedStart.GetPoint dStartX, dStartY, dStartZ
            oBoundedPos.Set dStartX, dStartY, dStartZ
        ElseIf oBoundedMemberPort.PortIndex = SPSMemberAxisEnd Then
            oBoundedEnd.GetPoint dEndX, dEndY, dEndZ
            oBoundedPos.Set dEndX, dEndY, dEndZ
        End If
    End If

    sMsg = "computing the  vectors for bounded"
    
    oBoundedMemberPart.Rotation.GetTransformAtPosition oBoundedPos.x, oBoundedPos.y, oBoundedPos.z, oBoundedXSectionMatrix, oDummyMatrix
    oBoundedU.Set oBoundedXSectionMatrix.IndexValue(4), oBoundedXSectionMatrix.IndexValue(5), oBoundedXSectionMatrix.IndexValue(6)
    oBoundedV.Set oBoundedXSectionMatrix.IndexValue(8), oBoundedXSectionMatrix.IndexValue(9), oBoundedXSectionMatrix.IndexValue(10)
                  
    ElseIf TypeOf oBoundedPart Is IJProfile Then
        Set oBdedLandingCurve = GetProfilePartLandingCurve(oBoundedPart)
        oBdedLandingCurve.GetEndPoints posStart, posEnd
        If TypeOf oBoundedPort Is IJStructPort Then
            Set oBoundedStructPort = oBoundedPort
            If oBoundedStructPort.ContextID = CTX_BASE Then
                oBoundedPos.Set posStart.x, posStart.y, posStart.z
            ElseIf oBoundedStructPort.ContextID = CTX_OFFSET Then
                oBoundedPos.Set posEnd.x, posEnd.y, posEnd.z
            Else
                ' We dont expect the code to come here
                'Error Out ot Need to Modify the Logic
            End If
        End If
        Dim oProfile As IJProfileAttributes
        Dim oProfUVec As IJDVector
        Dim oProfVVec As IJDVector
        Dim oOriginPos As IJDPosition
        Set oProfile = New ProfileUtils
        oProfile.GetProfileOrientationAndLocation oBoundedPart, oBoundedPos, oProfUVec, oProfVVec, oOriginPos
        oBoundedU.Set oProfUVec.x, oProfUVec.y, oProfUVec.z
        oBoundedV.Set oProfVVec.x, oProfVVec.y, oProfVVec.z
    End If
    
    Set oBUMember = Nothing
    bIsBuiltup = False
    
'    Set oBoundingPart = oBoundingPort.Connectable
    IsFromBuiltUpMember oBoundingPort.Connectable, bIsBuiltup, oBUMember
    
    If bIsBuiltup Then
        Set oBoundingPart = oBUMember
    Else
        Set oBoundingPart = oBoundingPort.Connectable
    End If
    
    ' Step 2: compute vectors for bounding
    If TypeOf oBoundingPart Is IJStructProfilePart Then
        'Case for Members and Profiles
'        Set oBoundingPart = oBoundingPort.Connectable
        Set oBoundingXSectionMatrix = New DT4x4
        Set oBoundingXSectionMatrix = oBoundingPart.GetCrossSectionMatrixAtPoint(oBoundedPos)
        'Set BoundingW to the Z Axis of the Bounding cross section
        oBoundingW.Set oBoundingXSectionMatrix.IndexValue(8), oBoundingXSectionMatrix.IndexValue(9), oBoundingXSectionMatrix.IndexValue(10)
    End If
                    
    oBoundedU.Length = 1#
    oBoundedV.Length = 1#
    oBoundingW.Length = 1#

    Dim dot_wU As Double
    Dim dot_wV As Double
    
    sMsg = "Getting the dot product"
    
    dot_wU = Round(Abs(oBoundingW.Dot(oBoundedU)), 4)
    dot_wV = Round(Abs(oBoundingW.Dot(oBoundedV)), 4)

    If GreaterThan(dot_wV, dot_wU) Then
        IsWebPenetrated = False
    Else
        IsWebPenetrated = True
    End If
    
    Set oBoundingXSectionMatrix = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT, sMsg).Number
    
End Function


'***********************************************************************
' METHOD:  IsPortFromBuiltUpMember
'
' DESCRIPTION:  Compare the normal of the molded surface of two
'               connected plate parts. If they point to the same
'               direction, then return true. If not, return false.
'***********************************************************************
Public Sub IsFromBuiltUpMember(oObject As Object, _
                               bFromBuiltUp As Boolean, _
                               Optional oBuiltupMember As ISPSDesignedMember)
Const METHOD = "::IsFromBuiltUpMember"
On Error GoTo ErrorHandler
    
    bFromBuiltUp = False
    
    Dim oport As IJPort
    Dim oConnectable As Object
    Dim oParentObject As Object
    Dim oDesignChild As IJDesignChild
    Dim oPlateSystem As IJPlateSystem
    Dim oDesignParent As IJDesignParent
    
    ' Check if type of Object passed in
    If oObject Is Nothing Then
        Exit Sub
        
    ElseIf TypeOf oObject Is IJPort Then
        ' a IJPort was passed in: get its Connectable
        Set oport = oObject
        Set oConnectable = oport.Connectable
        Set oport = Nothing
        
        If Not TypeOf oConnectable Is IJPlatePart Then
            Exit Sub
        End If
        
    ElseIf TypeOf oObject Is IJPlatePart Then
        ' a IJPlatePart was passed in
        Set oConnectable = oObject
    
    ElseIf TypeOf oObject Is IJPlateSystem Then
        ' a IJPlateSystem was passed in (may be Root or Leaf System)
        Set oParentObject = oObject
    
    Else
        Exit Sub
    End If
    
    ' If Have a Connectable ( a IJPlatePart from Port or passed in)
    ' Get the Plate Part's Parent Object
    If Not oConnectable Is Nothing Then
        If TypeOf oConnectable Is IJDesignChild Then
            Set oDesignChild = oConnectable
            Set oParentObject = oDesignChild.GetParent
            Set oDesignChild = Nothing
        End If
    End If
    
    ' Verify have a valid IJPlateSystem either from the Port/IJPlatePate or passed in
    If oParentObject Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oParentObject Is IJPlateSystem Then
        Exit Sub
    
    ElseIf Not TypeOf oParentObject Is IJDesignChild Then
        Exit Sub
    End If
    
    ' Check if current IJPlateSystem is from a BuiltUP:
    ' (a Leaf Plate System is never)
    Set oPlateSystem = oParentObject
    bFromBuiltUp = oPlateSystem.IsBuiltupPlateSystem
    If bFromBuiltUp Then
        Set oBuiltupMember = oPlateSystem.ParentBuiltup
        Exit Sub
    End If
    
    ' The Above PlateSystem may be a Leaf Plate System
    ' Get the Above Plate Systems's Parent object
    Set oDesignChild = oParentObject
    Set oParentObject = oDesignChild.GetParent
    Set oDesignChild = Nothing
            
    ' Check if the Plate System's Parent object is IJPlateSystem
    If TypeOf oParentObject Is IJPlateSystem Then
        Set oPlateSystem = oParentObject
        bFromBuiltUp = oPlateSystem.IsBuiltupPlateSystem
                
        If bFromBuiltUp Then
            Set oBuiltupMember = oPlateSystem.ParentBuiltup
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub



'*************************************************************************
'Function
'IsTubularMember
'
'Abstract
'   Given the Member Part or member Axis Port (ISPSSplitAxisPort)
'   Check if the Cross Section type is Tubular
'
'input
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Function IsTubularMember(oMemberObject As Object) As Boolean
Const METHOD = "::IsTubularMember"
    
    Dim sMsg As String
    On Error GoTo ErrorHandler
    
    Dim sCStype As String
    
    Dim oport As IJPort
    Dim oMemberPart As ISPSMemberPartCommon
    Dim oSplitAxisPort As ISPSSplitAxisPort
    
    Dim oCrossSection As IJCrossSection
    Dim oPartDesigned As ISPSDesignedMember
    Dim oPartPrismatic As ISPSMemberPartPrismatic
    Dim oSPSCrossSection As ISPSCrossSection
    Dim oStiffenerPart As StructDetailObjects.ProfilePart
    
    Set oStiffenerPart = New StructDetailObjects.ProfilePart
    
    IsTubularMember = False
    
    Dim bIsBuiltup As Boolean, oBUMember As ISPSDesignedMember
    IsFromBuiltUpMember oMemberObject, bIsBuiltup, oBUMember
    Dim otempObj As Object
    
    Set otempObj = oMemberObject
    
    If bIsBuiltup Then
        Set oMemberObject = oBUMember
    End If
    
    If TypeOf oMemberObject Is ISPSSplitAxisPort Then
        Set oport = oMemberObject
        Set oSplitAxisPort = oMemberObject
        Set oMemberPart = oport.Connectable
    
    ElseIf TypeOf oMemberObject Is ISPSMemberPartCommon Then
        Set oMemberPart = oMemberObject
    ElseIf TypeOf oMemberObject Is IJProfile Then
        Set oStiffenerPart.object = oMemberObject
        sCStype = oStiffenerPart.SectionType
    Else
        IsTubularMember = False
        GoTo CleanUp
    End If
    
    If Not oMemberPart Is Nothing Then
    If oMemberPart.IsPrismatic Then
        Set oPartPrismatic = oMemberPart
        Set oSPSCrossSection = oPartPrismatic.crossSection
    
    ElseIf TypeOf oMemberPart Is ISPSDesignedMember Then
        Set oPartDesigned = oMemberPart
        Set oSPSCrossSection = oMemberPart
        ElseIf Not oMemberPart Is Nothing Then
        IsTubularMember = False
        GoTo CleanUp
    End If
    End If
        
    ' Verify Bounded have valid Cross Section Type
    If Not oSPSCrossSection Is Nothing Then
    If TypeOf oSPSCrossSection.Definition Is IJCrossSection Then
        Set oCrossSection = oSPSCrossSection.Definition
        sCStype = oCrossSection.Type
        Else
            IsTubularMember = False
        End If
    End If
        
    If Trim(LCase(sCStype)) = LCase("CS") Then
        IsTubularMember = True
    ElseIf Trim(LCase(sCStype)) = LCase("HSSC") Then
        IsTubularMember = True
    ElseIf Trim(LCase(sCStype)) = LCase("PIPE") Then
        IsTubularMember = True
    ElseIf Trim(LCase(sCStype)) = LCase("P") Or Trim(LCase(sCStype)) = LCase("R") Then
        IsTubularMember = True
    ElseIf Trim(LCase(sCStype)) = LCase("BUTube") Then
        IsTubularMember = True
    Else
        IsTubularMember = False
    End If
    
    
CleanUp:
    Set oMemberObject = otempObj
    Set otempObj = Nothing
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function



' ********************************************************************************
' Method:
'   LandingCurve
' Description:
'   Gets the landingcurve for the given Stiffener/ER/Beam
' ********************************************************************************
' This method copied from EndCutRules\Common.bas
' EndCuts should be made to use the MarineLibraryCommon.bas file, which is in a more public location
Public Function GetProfilePartLandingCurve(oProfilePart As Object) As IJWireBody
    On Error GoTo ErrorHandler

    Dim oStructDetailHelper As StructDetailHelper
    Dim oTopoLocate As IJTopologyLocate
    Dim oLandingCrv As IJWireBody

    Set oStructDetailHelper = New StructDetailHelper
    Set oTopoLocate = New TopologyLocate

    'Based on whether the part is derived from system or not,
    'we shall get the landing curve differently for performance

    Dim oIJStructGraph As IJStructGraph
    Set oIJStructGraph = oProfilePart

    If Not oIJStructGraph Is Nothing Then
        Dim oParentSystem As IJSystem
        oStructDetailHelper.IsPartDerivedFromSystem oIJStructGraph, oParentSystem

        'Derived from system?
        If Not oParentSystem Is Nothing Then
            Set oLandingCrv = oTopoLocate.GetProfileParentWireBody(oProfilePart)
        Else
            'Not derived from system
            Dim oPartSupport As GSCADSDPartSupport.IJPartSupport
            Dim oProfilePartSupport As GSCADSDPartSupport.IJProfilePartSupport

            Set oProfilePartSupport = New GSCADSDPartSupport.ProfilePartSupport
            Set oPartSupport = oProfilePartSupport
            Set oPartSupport.Part = oProfilePart

            ' Get the curve (default direction is SideUnspecified, which gives
            ' a curve through the load point)
            Dim oThicknessDir As IJDVector
            Dim bThicknessCentered As Boolean

            oProfilePartSupport.GetProfilePartLandingCurve oLandingCrv, _
                                                           oThicknessDir, _
                                                           bThicknessCentered

            Set oProfilePartSupport = Nothing
            Set oPartSupport = Nothing
            Set oThicknessDir = Nothing
        End If

        Set oParentSystem = Nothing
    End If

    Set GetProfilePartLandingCurve = oLandingCrv

    Set oLandingCrv = Nothing
    Set oStructDetailHelper = Nothing
    Set oTopoLocate = Nothing

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "GetProfilePartLandingCurve").Number
End Function


'************************************************************************
' Method : GreaterThan --- checkes whether left side double variable is
'          greater than right side double variable, uses given tolerance
'
'  Inputs --------- Two double variables and tolerance(optional)
'
'  Output --------- True or False
'
'  NOTE: GreaterThanZero method need to be used if right side variable
'        is zero '0#' for comparison
'************************************************************************
Public Function GreaterThan(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If (LeftVariable > (RightVariable - Tolerance)) Then
        GreaterThan = True
    Else
        GreaterThan = False
    End If

End Function

'------------------------------------------------------------------------------------------------------------
' METHOD:  GetExtendedPort
'
' DESCRIPTION:  Gets the ExtendedPort(Port which gets extended before
'               Trim is applied) of the Member Part(Standard(Rolled) Members)
'
'               Method can be enhanced for Built Up's(Designed Members)
'
' Inputs : oPort : Any existing Lateral Port of the member part
'
'
' Output : Returns Extend Port of the Argument passed
'          (Returned Output may or maynot be of Type IJPort)
'------------------------------------------------------------------------------------------------------------
Public Function GetExtendedPort(oport As IJPort) As Object

 Const METHOD = "GetExtendedPort"
 Dim sMsg As String
 
 On Error GoTo ErrorHandler
 
     Dim oConnectable As IJConnectable
     Dim strJXSEC_CODE As String
     Dim oMemberFactory As SPSMembers.SPSMemberFactory
     Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices
     Dim oStructProfPart As IJStructProfilePart
     Dim lCtxId As Long
     Dim lOptId As Long
     Dim lOprId As Long
     Dim oExtendedPortElem As IJElements
     Dim CollofFaces() As String
     Dim i As Long
     Dim ePortType As JS_TOPOLOGY_PROXY_TYPE
     
     Set oConnectable = oport.Connectable
     
     'SP3D Member Object Ports do not implement IJStructPort interface
     'if the port object is other than Member Object then the method needs to be enhanced
     If TypeOf oConnectable Is ISPSMemberPartPrismatic Then
     
            Set oMemberFactory = New SPSMembers.SPSMemberFactory
            Set oMemberConnectionServices = oMemberFactory.CreateConnectionServices
                           
            oMemberConnectionServices.GetStructPortInfo oport, ePortType, _
                                                        lCtxId, lOptId, lOprId
    ElseIf TypeOf oConnectable Is IJProfile Then
           
        Dim oStructPort As IJStructPort
        Set oStructPort = oport
        lOprId = oStructPort.OperatorID
        
    Else 'Could be Designed (built-up) member
       sMsg = METHOD & " not available for built-up members"
       GoTo ErrorHandler
    End If
            
            If lOprId = JXSEC_TOP Then
                strJXSEC_CODE = "514"
            ElseIf lOprId = JXSEC_BOTTOM Then
                strJXSEC_CODE = "513"
            ElseIf lOprId = JXSEC_WEB_LEFT Then
                strJXSEC_CODE = "257"
            ElseIf lOprId = JXSEC_WEB_RIGHT Then
                strJXSEC_CODE = "258"
            ElseIf lOprId = JXSEC_TOP_FLANGE_RIGHT Then
                strJXSEC_CODE = "1028"
            ElseIf lOprId = JXSEC_TOP_FLANGE_LEFT Then
                strJXSEC_CODE = "1026"
            ElseIf lOprId = JXSEC_BOTTOM_FLANGE_RIGHT Then
                strJXSEC_CODE = "1027"
            ElseIf lOprId = JXSEC_BOTTOM_FLANGE_LEFT Then
                strJXSEC_CODE = "1025"
            ElseIf lOprId = JXSEC_TOP_FLANGE_LEFT_BOTTOM Then
                strJXSEC_CODE = "770"
            ElseIf lOprId = JXSEC_TOP_FLANGE_RIGHT_BOTTOM Then
                strJXSEC_CODE = "772"
            ElseIf lOprId = JXSEC_BOTTOM_FLANGE_LEFT_TOP Then
                strJXSEC_CODE = "769"
            ElseIf lOprId = JXSEC_BOTTOM_FLANGE_RIGHT_TOP Then
                strJXSEC_CODE = "771"
            End If
            
            If TypeOf oConnectable Is IJStructProfilePart Then
                Set oStructProfPart = oConnectable
                oStructProfPart.GetSectionFaces True, oExtendedPortElem, CollofFaces()
                ' Verify returned List of Port(s) is valid
                If oExtendedPortElem Is Nothing Then
                    sMsg = "oExtendedPortElem Is Nothing"
                ElseIf oExtendedPortElem.Count < 1 Then
                    sMsg = "oExtendedPortElem.Count < 1"
                Else
                    For i = 1 To oExtendedPortElem.Count
                        If CollofFaces(i - 1) = strJXSEC_CODE Then
                           Set GetExtendedPort = oExtendedPortElem.Item(i)
                           Exit For
                        End If
                    Next i
                End If
            Else
                sMsg = METHOD & " : Object doesn't support IJStructProfilePart Interface"
                GoTo ErrorHandler
            End If
      
      Set oConnectable = Nothing
      Set oStructProfPart = Nothing
      Set oExtendedPortElem = Nothing
      Set oMemberFactory = Nothing
      Set oMemberConnectionServices = Nothing
      
 Exit Function
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function

'------------------------------------------------------------------------------------------------------------
' METHOD:  GetLateralSubPortBeforeTrim
'
' DESCRIPTION:  Gets the Lateral Sub Port of the Member Part(Standard(Rolled) Members) Before Trim/Cut
'               and returns it
'               Method can be enhanced for ProfileParts/Built Up's(Designed Members)
'
' Inputs : oMemberPart : Member Part on which the port exist
'          eSubPort : Enum of the port on the Member part which is to be retrieved(for e.g WEB_LEFT or WEB_RIGHT etc)
'
' Output : A Lateral SubPortBeforeTrim is returned(depending on eSubPort passed, if exists)
'          Or Nothing (if eSubPort passed doesnt exist on member part)
'------------------------------------------------------------------------------------------------------------
Public Function GetLateralSubPortBeforeTrim(oMemberPart As Object, ByVal eSubPort As JXSEC_CODE) As IJPort

    Const METHOD = "GetLateralSubPortBeforeTrim"
    
    Dim sMsg As String
    
    Dim lCtxId As Long
    Dim lOprId As Long
    Dim lOptId As Long
    
    Dim ePortType As JS_TOPOLOGY_PROXY_TYPE
    Dim eFilterType As JS_TOPOLOGY_FILTER_TYPE
    
    Dim oport As IJPort
    Dim oPortObject As Object
    Dim oPortElements As IJElements
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    Dim oPlatePart As New StructDetailObjects.PlatePart
    
    Dim oStructGraphConnectable As IJStructGraphConnectable
    
    Dim oMemberFactory As SPSMembers.SPSMemberFactory
    Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices
    
    Dim oSD_MemberPart As New StructDetailObjects.MemberPart
    
    On Error GoTo ErrorHandler
    
    Set GetLateralSubPortBeforeTrim = Nothing
    
    If TypeOf oMemberPart Is ISPSMemberPartPrismatic Then 'Standard (rolled) member
        
            eFilterType = JS_TOPOLOGY_FILTER_LCONNECT_PRT_SUB_LFACES
            
        ' Verify current Member Part
        If TypeOf oMemberPart Is IJStructGraphConnectable Then
            Set oStructGraphConnectable = oMemberPart
        Else
            sMsg = "Else ... Typeof(MemberPart):" & TypeName(oMemberPart)
            GoTo ErrorHandler
        End If
            
        ' Retreive list of Ports from the SPS Member Part's Solid Geometry before Member Cut Operation
        oStructGraphConnectable.enumPortsInGraphByTopologyFilter oPortElements, _
                                                                 eFilterType, _
                                                                 StableGeometry, _
                                                                 vbNull
        ' Verify returned List of Port(s) is valid
        If oPortElements Is Nothing Then
            sMsg = "oPortElements Is Nothing"
        ElseIf oPortElements.Count < 1 Then
            sMsg = "oPortElements.Count < 1"
        Else
            
            For Each oPortObject In oPortElements
            
                Set oMemberFactory = New SPSMembers.SPSMemberFactory
                Set oMemberConnectionServices = oMemberFactory.CreateConnectionServices
                
                oMemberConnectionServices.GetStructPortInfo oPortObject, ePortType, _
                                                                                     lCtxId, lOptId, lOprId
                                                            
                If TypeOf oPortObject Is IJPort Then
                    Set oport = oPortObject
                    If lOprId = eSubPort Then
                        Set GetLateralSubPortBeforeTrim = oport
                        Exit For
                    End If
                    
                Else
                    sMsg = "Else... TypeOf oPortObject Is IJPort"
                    GoTo ErrorHandler
                End If
                
             Next oPortObject
             
        End If
        
        Set oport = Nothing
        Set oPortObject = Nothing
        Set oPortElements = Nothing
        Set oStructGraphConnectable = Nothing
        Set oMemberFactory = Nothing
        Set oMemberConnectionServices = Nothing
                                                                                
    ElseIf TypeOf oMemberPart Is IJProfile Then
        Set oProfilePart.object = oMemberPart
        Set GetLateralSubPortBeforeTrim = oProfilePart.SubPortBeforeTrim(eSubPort)
        
    ElseIf TypeOf oMemberPart Is IJPlate Then
    'Designed (built-up) member
'        sMsg = METHOD & " not available for built-up members"
'        GoTo ErrorHandler
        
        Set oPlatePart.object = oMemberPart
        Set GetLateralSubPortBeforeTrim = oPlatePart.BasePortFromOperation(BPT_Lateral, "CreatePlatePart.GeneratePlatePart_AE.1", False)
    End If

    Set oProfilePart = Nothing
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
End Function


