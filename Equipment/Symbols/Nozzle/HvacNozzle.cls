VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HvacNozzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 1998-2005, Intergraph Corporation. All rights reserved.
'
'File
'    HvacNozzle.cls
'
'Author
'    RM
'
'Description
'    Definition of Hvac Nozzle Symbol
'
'Notes
'
'    <notes>
'
'   History
'      Mar/08/2000   SS
'                                               used JObjectCollection instead of UntransactedMiddleElems
'       14Feb2003     ALP           Modified to change the definition of the "FlatOval" cross section.
'                                              Treat the input width as "width including arcs".
'       09Sep2003      VS                   Changed the TOLERANCE to 10^-10.
'       15Nov2005     VS               Remove the maintenance aspect.
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const E_FAIL = -2147467259
Private Const TOLERANCE = 0.0000000001
Dim m_outputColl As IJDOutputCollection
Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
On Error GoTo ErrorHandler
 IJDUserSymbolServices_EditOccurence = False
 Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
On Error GoTo ErrorHandler
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    FeedNozzleDefinition pSymbolDefinition
  Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
On Error GoTo ErrorHandler
  IJDUserSymbolServices_GetDefinitionName = "GSCADNozzleVBSymb.HvacNozzle"
  Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler

  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
  Set oSymbolFactory = Nothing
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "GSCADNozzleVBSymb.HvacNozzle"
  oSymbolDefinition.CodeBase = CodeBase

  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId

  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  Set oSymbolDefinition = Nothing

  Exit Function

ErrorHandler:
    Err.Raise E_FAIL
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal RepName As String, ByVal OutputColl As Object, ByRef arrayOfInputs())
    On Error GoTo ErrorHandler
    Set m_outputColl = OutputColl
    If StrComp(RepName, "Physical", vbTextCompare) = 0 Then
        Symbolic arrayOfInputs
    ElseIf StrComp(RepName, "Insulation", vbTextCompare) = 0 Then
        InsulateHVACNozzle arrayOfInputs
    Else
        'Do nothing this case cannot come at all
    End If
    
Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

Private Sub Symbolic(ByRef arrayOfInputs())
    On Error GoTo ErrorHandler
    
    Dim Thickness As Double, FlangeWidth As Double
    Dim Width As Double, Depth As Double
    Dim CornerRadius As Double
    Dim DimensionBaseOuter As Boolean
    Dim HvacLength As Double
    Dim CrossSectionShape As OutfittingCrossSectionsFacelets.CrossSectionShapeTypes
    Dim CptOffset As Double, PortDepth As Double
    
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim projection As IngrGeom3D.Projection3d
    Dim projVecX As Double, projVecY As Double, projVecZ As Double
         
    Dim X As Double
    Dim Ymin As Double, Zmin As Double
    Dim Ymax As Double, Zmax As Double
    Dim Radius As Double
    
    Dim P1(0 To 2) As Double, P2(0 To 2) As Double, P3(0 To 2) As Double
    Dim P4(0 To 2) As Double, P5(0 To 2) As Double, P6(0 To 2) As Double
    Dim P7(0 To 2) As Double, P8(0 To 2) As Double
    Dim C1(0 To 2) As Double, C2(0 To 2) As Double
    Dim C3(0 To 2) As Double, C4(0 To 2) As Double
     
    Dim complexStr As IngrGeom3D.ComplexString3d
    Dim line As IngrGeom3D.Line3d
    Dim arc As IngrGeom3D.Arc3d
    Dim coll As IJElements
    Dim elem As Object
         
    CrossSectionShape = arrayOfInputs(1)
    Thickness = arrayOfInputs(2)
    FlangeWidth = arrayOfInputs(3)
    Width = arrayOfInputs(4)
    Depth = arrayOfInputs(5)
    CornerRadius = arrayOfInputs(6)
    DimensionBaseOuter = arrayOfInputs(7)
    HvacLength = arrayOfInputs(8)
    CptOffset = arrayOfInputs(9)
    PortDepth = arrayOfInputs(10)
    
    Dim oErrors As IJEditErrors
    Set oErrors = New IMSErrorLog.JServerErrors
'    oErrors.Clear
    
    If HvacLength < TOLERANCE Then
        oErrors.Add E_FAIL, "GSCADNozzleVBsymb.HvacNozzle", "NozzleLength must be greater than FlangeThickness", "Thickness > NozzleLength"
        GoTo ErrorHandler
    End If
    
    If (Width < TOLERANCE) Or (Depth < TOLERANCE) Then
       oErrors.Add E_FAIL, "GSCADNozzleVBsymb.HvacNozzle", "Shape Width/Depth can not be Zero", "WidthDepth Zero"
       GoTo ErrorHandler
    End If
    
    If (CornerRadius < 0) Then
       oErrors.Add E_FAIL, "GSCADNozzleVBsymb.HvacNozzle", "Corner Radius should be Greater than or Equal to zero", "CornerRadius > or = Zero"
       GoTo ErrorHandler
    End If
    
    If (CornerRadius > (Width / 2)) Or (CornerRadius > (Depth / 2)) Then
       oErrors.Add E_FAIL, "GSCADNozzleVBsymb.HvacNozzle", "Corner Radius Should not be greater than Half of Width/Depth", "CornerRadius > Width/Depth/2"
       GoTo ErrorHandler
    End If
    
    Select Case CrossSectionShape
    Case OutfittingCrossSectionsFacelets.Rectangular
        '*************************************************************************
        '           Rectangular
        '*************************************************************************
        ' Construction of Flange body
        '              P3               P4
        '               ------------------              Y
        '           P2| C2              C3 |P5          |
        '             |        .(0,0,0)    |            |
        '             |                    |             ------->Z
        '           P1| C1              C4 |P6
        '               ------------------
        '              P8               P7
        '
        ' Y axis is the radial direction
        
                
        X = CptOffset - PortDepth
        Ymin = -((Depth / 2) + FlangeWidth)
        Zmin = -((Width / 2) + FlangeWidth)
        Ymax = (Depth / 2) + FlangeWidth
        Zmax = (Width / 2) + FlangeWidth
        
        If CornerRadius = 0 Then
            Radius = 0#
        Else
            Radius = CornerRadius + FlangeWidth
        End If
        
        P1(0) = X
        P1(1) = Ymin + Radius
        P1(2) = Zmin
        P2(0) = X
        P2(1) = Ymax - Radius
        P2(2) = Zmin
        P3(0) = X
        P3(1) = Ymax
        P3(2) = Zmin + Radius
        P4(0) = X
        P4(1) = Ymax
        P4(2) = Zmax - Radius
        P5(0) = X
        P5(1) = Ymax - Radius
        P5(2) = Zmax
        P6(0) = X
        P6(1) = Ymin + Radius
        P6(2) = Zmax
        P7(0) = X
        P7(1) = Ymin
        P7(2) = Zmax - Radius
        P8(0) = X
        P8(1) = Ymin
        P8(2) = Zmin + Radius
        
        C1(0) = X
        C1(1) = Ymin + Radius
        C1(2) = Zmin + Radius
        C2(0) = X
        C2(1) = Ymax - Radius
        C2(2) = Zmin + Radius
        C3(0) = X
        C3(1) = Ymax - Radius
        C3(2) = Zmax - Radius
        C4(0) = X
        C4(1) = Ymin + Radius
        C4(2) = Zmax - Radius
        
        ' Create a complex string
        
        Set coll = New JObjectCollection
              
        'Construct individual lines and arcs and create a complex string
        
        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P1(0), P1(1), P1(2), P2(0), P2(1), P2(2))
        coll.Add line
        Set line = Nothing
        
        If Radius > 0 Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        C2(0), C2(1), C2(2), P2(0), P2(1), P2(2), _
                                        P3(0), P3(1), P3(2))
            coll.Add arc
            Set arc = Nothing
        End If

        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P3(0), P3(1), P3(2), P4(0), P4(1), P4(2))
        coll.Add line
        Set line = Nothing

        If Radius > 0 Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        C3(0), C3(1), C3(2), P4(0), P4(1), P4(2), _
                                        P5(0), P5(1), P5(2))
            coll.Add arc
            Set arc = Nothing
        End If

        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P5(0), P5(1), P5(2), P6(0), P6(1), P6(2))
        coll.Add line
        Set line = Nothing

        If Radius > 0 Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        C4(0), C4(1), C4(2), P6(0), P6(1), P6(2), _
                                        P7(0), P7(1), P7(2))
            coll.Add arc
            Set arc = Nothing
        End If

        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P7(0), P7(1), P7(2), P8(0), P8(1), P8(2))
        coll.Add line
        Set line = Nothing

        If Radius > 0 Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        C1(0), C1(1), C1(2), P8(0), P8(1), P8(2), _
                                        P1(0), P1(1), P1(2))
            coll.Add arc
            Set arc = Nothing
        End If
        
        ' Create complex string
        Set complexStr = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, coll)
        
        ' Create projection
        
        projVecX = 1#
        projVecY = 0#
        projVecZ = 0#
    
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, complexStr, _
                                        projVecX, projVecY, projVecZ, Thickness, True)
        m_outputColl.AddOutput "Flange", projection
        Set projection = Nothing
        
        ' Clear the lines and arcs in collection
        coll.Clear
        
'        complexStr.RemoveAll

        'Delete complexString which was created for creating projection.
        Set complexStr = Nothing

                
        'CONSTRUCTION OF  Hvac Cross Section
        X = X + Thickness
        Ymin = -(Depth / 2)
        Zmin = -(Width / 2)
        Ymax = (Depth / 2)
        Zmax = (Width / 2)
        Radius = CornerRadius
        
        P1(0) = X
        P1(1) = Ymin + Radius
        P1(2) = Zmin
        P2(0) = X
        P2(1) = Ymax - Radius
        P2(2) = Zmin
        P3(0) = X
        P3(1) = Ymax
        P3(2) = Zmin + Radius
        P4(0) = X
        P4(1) = Ymax
        P4(2) = Zmax - Radius
        P5(0) = X
        P5(1) = Ymax - Radius
        P5(2) = Zmax
        P6(0) = X
        P6(1) = Ymin + Radius
        P6(2) = Zmax
        P7(0) = X
        P7(1) = Ymin
        P7(2) = Zmax - Radius
        P8(0) = X
        P8(1) = Ymin
        P8(2) = Zmin + Radius
        
        C1(0) = X
        C1(1) = Ymin + Radius
        C1(2) = Zmin + Radius
        C2(0) = X
        C2(1) = Ymax - Radius
        C2(2) = Zmin + Radius
        C3(0) = X
        C3(1) = Ymax - Radius
        C3(2) = Zmax - Radius
        C4(0) = X
        C4(1) = Ymin + Radius
        C4(2) = Zmax - Radius
        
        'Construct individual lines and arcs and add to the complex string
        
        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P1(0), P1(1), P1(2), P2(0), P2(1), P2(2))
        coll.Add line
        Set line = Nothing
        
        If Radius > 0 Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        C2(0), C2(1), C2(2), P2(0), P2(1), P2(2), _
                                        P3(0), P3(1), P3(2))
            coll.Add arc
            Set arc = Nothing
        End If
        
        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P3(0), P3(1), P3(2), P4(0), P4(1), P4(2))
        coll.Add line
        Set line = Nothing
        
        If Radius > 0 Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        C3(0), C3(1), C3(2), P4(0), P4(1), P4(2), _
                                        P5(0), P5(1), P5(2))
            coll.Add arc
            Set arc = Nothing
        End If
        
        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P5(0), P5(1), P5(2), P6(0), P6(1), P6(2))
        coll.Add line
        Set line = Nothing
        
        If Radius > 0 Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        C4(0), C4(1), C4(2), P6(0), P6(1), P6(2), _
                                        P7(0), P7(1), P7(2))
            coll.Add arc
            Set arc = Nothing
        End If
        
        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P7(0), P7(1), P7(2), P8(0), P8(1), P8(2))
        coll.Add line
        Set line = Nothing
        
        If Radius > 0 Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        C1(0), C1(1), C1(2), P8(0), P8(1), P8(2), _
                                        P1(0), P1(1), P1(2))
            coll.Add arc
            Set arc = Nothing
        End If
        
        ' Create complex string
        Set complexStr = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, coll)
        
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, complexStr, _
                                        projVecX, projVecY, projVecZ, HvacLength, True)
        m_outputColl.AddOutput "Hvac", projection
        Set projection = Nothing
        
        ' clear the lines and arcs in collection
        coll.Clear
        
'        complexStr.RemoveAll

        Set complexStr = Nothing
        Set geomFactory = Nothing

    Case OutfittingCrossSectionsFacelets.FlatOval
    '*************************************************************************
    '           FlatOval
    '*************************************************************************
    '                                                   ^ Y
    '                 P2----------------P3              |
    '                 *                   *             |
    '            P5 *                       *P6          -------> Z
    '                 *                   *
    '                 P1----------------P4
    '
    ' Y axis is the radial direction
    
        ' The input width is the width including the arcs. We would need to deduce the
        ' flat width from the "input width"
        Dim dFlatWidth As Double
        dFlatWidth = Width - Depth
        If dFlatWidth < 0# Then dFlatWidth = 0.000001 ' use distance tolerance
    
        X = CptOffset - PortDepth
        Ymin = -((Depth / 2) + FlangeWidth)
        Zmin = -(dFlatWidth / 2)
        Ymax = (Depth / 2) + FlangeWidth
        Zmax = (dFlatWidth / 2)
        Radius = (Depth / 2) + FlangeWidth
        
        P1(0) = X
        P1(1) = Ymin
        P1(2) = Zmin
        P2(0) = X
        P2(1) = Ymax
        P2(2) = Zmin
        P3(0) = X
        P3(1) = Ymax
        P3(2) = Zmax
        P4(0) = X
        P4(1) = Ymin
        P4(2) = Zmax
        
        P5(0) = X
        P5(1) = 0#
        P5(2) = Zmin - Radius
        P6(0) = X
        P6(1) = 0#
        P6(2) = Zmax + Radius
                
        Set coll = New JObjectCollection
              
        'Construct individual lines and arcs and create a complex string
        
        Set arc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                    P1(0), P1(1), P1(2), P5(0), P5(1), P5(2), _
                                    P2(0), P2(1), P2(2))
        coll.Add arc
        Set arc = Nothing
        
        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P2(0), P2(1), P2(2), P3(0), P3(1), P3(2))
        coll.Add line
        Set line = Nothing
        
        Set arc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                    P3(0), P3(1), P3(2), P6(0), P6(1), P6(2), _
                                    P4(0), P4(1), P4(2))
        coll.Add arc
        Set arc = Nothing
        
        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P4(0), P4(1), P4(2), P1(0), P1(1), P1(2))
        coll.Add line
        Set line = Nothing

        
        ' Create complex string
        Set complexStr = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, coll)
        
        ' Create projection
        
        projVecX = 1#
        projVecY = 0#
        projVecZ = 0#
    
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, complexStr, _
                                        projVecX, projVecY, projVecZ, Thickness, True)
        m_outputColl.AddOutput "Flange", projection
        Set projection = Nothing
          
        ' Clear lines and arcs in collection
        coll.Clear
      
        Set complexStr = Nothing
                 
        'CONSTRUCTION OF  Hvac Cross Section
        X = X + Thickness
        Ymin = -(Depth / 2)
        Zmin = -(dFlatWidth / 2)
        Ymax = (Depth / 2)
        Zmax = (dFlatWidth / 2)
        Radius = Depth / 2
        
        P1(0) = X
        P1(1) = Ymin
        P1(2) = Zmin
        P2(0) = X
        P2(1) = Ymax
        P2(2) = Zmin
        P3(0) = X
        P3(1) = Ymax
        P3(2) = Zmax
        P4(0) = X
        P4(1) = Ymin
        P4(2) = Zmax
        
        P5(0) = X
        P5(1) = 0#
        P5(2) = Zmin - Radius
        P6(0) = X
        P6(1) = 0#
        P6(2) = Zmax + Radius
                
        'Construct individual lines and arcs and create a complex string
        
        Set arc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                    P1(0), P1(1), P1(2), P5(0), P5(1), P5(2), _
                                    P2(0), P2(1), P2(2))
        coll.Add arc
        Set arc = Nothing
        
        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P2(0), P2(1), P2(2), P3(0), P3(1), P3(2))
        coll.Add line
        Set line = Nothing
        
        Set arc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                    P3(0), P3(1), P3(2), P6(0), P6(1), P6(2), _
                                    P4(0), P4(1), P4(2))
        coll.Add arc
        Set arc = Nothing
        
        Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    P4(0), P4(1), P4(2), P1(0), P1(1), P1(2))
        coll.Add line
        Set line = Nothing

        
        ' Create complex string
        Set complexStr = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, coll)
        
        ' Create projection
        
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, complexStr, _
                                        projVecX, projVecY, projVecZ, HvacLength, True)
        m_outputColl.AddOutput "Hvac", projection
        Set projection = Nothing
        
        ' Clear lines and arcs in collection
        coll.Clear
        
        Set complexStr = Nothing
        Set geomFactory = Nothing
        
    Case OutfittingCrossSectionsFacelets.Round
    '*************************************************************************
    '           ROUND
    '*************************************************************************
        'CONSTRUCTION OF  Flange Body Cylinder
        Dim circleCenterX As Double, circleCenterY As Double, circleCenterZ As Double
        circleCenterX = CptOffset - PortDepth
        circleCenterY = 0#
        circleCenterZ = 0#
    
        Dim circleNormalX As Double, circleNormalY As Double, circleNormalZ As Double
        circleNormalX = 1#
        circleNormalY = 0#
        circleNormalZ = 0#
    
        Dim circlePointVecX As Double, circlePointVecY As Double, circlePointVecZ As Double
        circlePointVecX = 0#
        circlePointVecY = (Width * 0.5) + FlangeWidth
        circlePointVecZ = 0#
    
        Dim axesRatio As Double
        axesRatio = 1#
    
        Dim ellipse As IngrGeom3D.Ellipse3d
        Set ellipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                    circleCenterX, circleCenterY, circleCenterZ, _
                                    circleNormalX, circleNormalY, circleNormalZ, _
                                    circlePointVecX, circlePointVecY, circlePointVecZ, _
                                    axesRatio)
    
        projVecX = 1#
        projVecY = 0#
        projVecZ = 0#
    
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, ellipse, _
                                        projVecX, projVecY, projVecZ, Thickness, True)
        m_outputColl.AddOutput "Flange", projection
        Set projection = Nothing
        Set ellipse = Nothing
    
        'CONSTRUCTION OF  Hvac Cylinder
        circleCenterX = circleCenterX + Thickness
        circlePointVecY = Width * 0.5
        Set ellipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                    circleCenterX, circleCenterY, circleCenterZ, _
                                    circleNormalX, circleNormalY, circleNormalZ, _
                                    circlePointVecX, circlePointVecY, circlePointVecZ, _
                                    axesRatio)
    
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, ellipse, _
                                        projVecX, projVecY, projVecZ, HvacLength, True)
        m_outputColl.AddOutput "Hvac", projection
        Set projection = Nothing
        Set ellipse = Nothing
        Set geomFactory = Nothing
    Case Else
    
    End Select
    
    Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

' Feed Nozzle Definition
' Inputs:
'           "CrossSectionShape"
'           "Thickness"
'           "FlangeWidth"
'           "Width"
'           "Depth"
'           "CornerRadius"
'           "DimensionBaseOuter"
'           "HvacLength"
' The representation "Symbolic" is set to the definition
' Symbolic :

Sub FeedNozzleDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
On Error GoTo ErrorHandler
  
  ' Create a new input by new operator
  Dim Inputs(0 To 11) As IMSSymbolEntities.IJDInput
  
  ' Create a defaultValue
  Dim PC As IMSSymbolEntities.DParameterContent
  Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
  PC.Type = igValue

  Set Inputs(0) = New IMSSymbolEntities.DInput
  Inputs(0).Name = "CrossSectionShape"
  Inputs(0).Description = "Shape of the CrossSection"
  Inputs(0).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = OutfittingCrossSectionsFacelets.Round
  Inputs(0).DefaultParameterValue = PC
  
  Set Inputs(1) = New IMSSymbolEntities.DInput
  Inputs(1).Name = "Thickness"
  Inputs(1).Description = "Thickness of the Flange"
  Inputs(1).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.05
  Inputs(1).DefaultParameterValue = PC
  
  Set Inputs(2) = New IMSSymbolEntities.DInput
  Inputs(2).Name = "FlangeWidth"
  Inputs(2).Description = "Incremental width of the flange"
  Inputs(2).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.05
  Inputs(2).DefaultParameterValue = PC
  
  Set Inputs(3) = New IMSSymbolEntities.DInput
  Inputs(3).Name = "Width"
  Inputs(3).Description = "Width of the crosssection"
  Inputs(3).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.15
  Inputs(3).DefaultParameterValue = PC
  
  Set Inputs(4) = New IMSSymbolEntities.DInput
  Inputs(4).Name = "Depth"
  Inputs(4).Description = "Depth of the crosssection"
  Inputs(4).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.15
  Inputs(4).DefaultParameterValue = PC
  
  Set Inputs(5) = New IMSSymbolEntities.DInput
  Inputs(5).Name = "CornerRadius"
  Inputs(5).Description = "Corner Radius of the crosssection"
  Inputs(5).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0#
  Inputs(5).DefaultParameterValue = PC
  
  Set Inputs(6) = New IMSSymbolEntities.DInput
  Inputs(6).Name = "DimensionBaseOuter"
  Inputs(6).Description = "DimensionBaseOuter flag of the crosssection"
  Inputs(6).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = True
  Inputs(6).DefaultParameterValue = PC
  
  Set Inputs(7) = New IMSSymbolEntities.DInput
  Inputs(7).Name = "HvacLength"
  Inputs(7).Description = "Length of the Hvac"
  Inputs(7).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.25
  Inputs(7).DefaultParameterValue = PC
  
  Set Inputs(8) = New IMSSymbolEntities.DInput
  Inputs(8).Name = "CptOffset"
  Inputs(8).Description = "Connect Point Offset"
  Inputs(8).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0#
  Inputs(8).DefaultParameterValue = PC
  
  Set Inputs(9) = New IMSSymbolEntities.DInput
  Inputs(9).Name = "PortDepth"
  Inputs(9).Description = "Port Depth"
  Inputs(9).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0#
  Inputs(9).DefaultParameterValue = PC
  
  Set Inputs(10) = New IMSSymbolEntities.DInput
  Inputs(10).Name = "InsulationCriteria"
  Inputs(10).Description = "Insulation Criteria"
  Inputs(10).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0#
  Inputs(10).DefaultParameterValue = PC
  
  Set Inputs(11) = New IMSSymbolEntities.DInput
  Inputs(11).Name = "InsulationThickness"
  Inputs(11).Description = "Insulation Thickness"
  Inputs(11).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = -1#
  Inputs(11).DefaultParameterValue = PC
  
  'Set the input to the definition
  Dim InputsIf As IMSSymbolEntities.IJDInputs
  Set InputsIf = pSymbolDefinition

  Dim index As Integer
  For index = 0 To 11
    InputsIf.SetInput Inputs(index), index + 1
  Next
  For index = 0 To 11
    Set Inputs(index) = Nothing
  Next
  

  'Create the output
  Dim O1 As IMSSymbolEntities.IJDOutput
  Set O1 = New IMSSymbolEntities.DOutput
  
  O1.Name = "Flange"
  O1.Description = "Flange of Hvac Nozzle"
  O1.Properties = 0
  
  Dim O2 As IMSSymbolEntities.IJDOutput
  Set O2 = New IMSSymbolEntities.DOutput

  O2.Name = "Hvac"
  O2.Description = "Hvac of Nozzle"
  O2.Properties = 0

  'Define the representation "Symbolic"
  Dim rep1 As IMSSymbolEntities.IJDRepresentation
  Set rep1 = New IMSSymbolEntities.DRepresentation

  rep1.Name = "Physical"
  rep1.Description = "Representation of Nozzle is Symbolic"
  rep1.Properties = igREPRESENTATION_ISVBFUNCTION
  rep1.RepresentationId = SimplePhysical

  Dim oRep1Outputs As IMSSymbolEntities.IJDOutputs
  Set oRep1Outputs = rep1

  'Set the outputs
  oRep1Outputs.SetOutput O1
  oRep1Outputs.SetOutput O2
  Set O1 = Nothing
  Set O2 = Nothing

  'Set the 'symbolic' representation to definition
  Dim RepsIf As IMSSymbolEntities.IJDRepresentations
  Set RepsIf = pSymbolDefinition
  RepsIf.SetRepresentation rep1

  Set rep1 = Nothing
  Set RepsIf = Nothing
  Set oRep1Outputs = Nothing

  'Define the evaluation for the Symbolic representation
  Dim SymbolicRepEval As IJDRepresentationEvaluation
  Set SymbolicRepEval = New DRepresentationEvaluation
  SymbolicRepEval.Name = "Physical"
  SymbolicRepEval.Description = "Symbolic representation of flanged nozzle"
  SymbolicRepEval.Properties = igREPRESENTATION_HIDDEN
  SymbolicRepEval.Type = igREPRESENTATION_VBFUNCTION
  SymbolicRepEval.ProgId = "GSCADNozzleVBsymb.HvacNozzle"

  'Set the evaluations on symbol definition
  Dim RepsEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
  Set RepsEvalsIf = pSymbolDefinition

  RepsEvalsIf.AddRepresentationEvaluation SymbolicRepEval

  Set RepsEvalsIf = Nothing
  Set SymbolicRepEval = Nothing
  
'=============================================================================================
'  Insulation Representation of HVAC Nozzle
'
'==============================================================================================

    'Create the outputs for Insulation Aspect
    Dim O3 As IMSSymbolEntities.IJDOutput
    Set O3 = New IMSSymbolEntities.DOutput
    O3.Name = "InsulatedFlangeInner"
    O3.Description = "Insulated Flange inner projection"
    O3.Properties = 0
    
    Dim O30 As IMSSymbolEntities.IJDOutput
    Set O30 = New IMSSymbolEntities.DOutput
    O30.Name = "InsulatedFlangeOuter"
    O30.Description = "Insulated Flange outer projection"
    O30.Properties = 0
    
    Dim O31 As IMSSymbolEntities.IJDOutput
    Set O31 = New IMSSymbolEntities.DOutput
    O31.Name = "InsulatedFlangeCap1"
    O31.Description = "Insulated Flange Cap1"
    O31.Properties = 0
    
    Dim O32 As IMSSymbolEntities.IJDOutput
    Set O32 = New IMSSymbolEntities.DOutput
    O32.Name = "InsulatedFlangeCap2"
    O32.Description = "Insulated Flange Cap2"
    O32.Properties = 0
        
    Dim O4 As IMSSymbolEntities.DOutput
    Set O4 = New IMSSymbolEntities.DOutput
    O4.Name = "InsulatedNeckOuter"
    O4.Description = "Insulated Neck Of Nozzle outer Projection"
    O4.Properties = 0
    
    Dim O40 As IMSSymbolEntities.DOutput
    Set O40 = New IMSSymbolEntities.DOutput
    O40.Name = "InsulatedNeckInner"
    O40.Description = "Insulated Neck Of Nozzle inner Projection"
    O40.Properties = 0
    
    Dim O41 As IMSSymbolEntities.DOutput
    Set O41 = New IMSSymbolEntities.DOutput
    O41.Name = "InsulatedNeckCap1"
    O41.Description = "Insulated Neck Cap1"
    O41.Properties = 0
    
    Dim O42 As IMSSymbolEntities.DOutput
    Set O42 = New IMSSymbolEntities.DOutput
    O42.Name = "InsulatedNeckCap2"
    O42.Description = "Insulated Neck Cap2"
    O42.Properties = 0
    
    Dim oRep2 As IMSSymbolEntities.IJDRepresentation
    Set oRep2 = New IMSSymbolEntities.DRepresentation
    oRep2.Name = "Insulation"
    oRep2.Description = "Insulation Representation of HVAC Nozzle"
    oRep2.Properties = igREPRESENTATION_ISVBFUNCTION
    oRep2.RepresentationId = Insulation
    
    Dim oRep2Outputs As IMSSymbolEntities.IJDOutputs
    Set oRep2Outputs = oRep2
    ' Set the Outputs
    'Flange geometry details
    oRep2Outputs.SetOutput O3
    oRep2Outputs.SetOutput O30
    oRep2Outputs.SetOutput O31
    oRep2Outputs.SetOutput O32
    
    'Neck Geometry details
    oRep2Outputs.SetOutput O4 'neck insulation outer
    oRep2Outputs.SetOutput O40 'neck Insulation Inner
    oRep2Outputs.SetOutput O41 'Cap on the left side
    oRep2Outputs.SetOutput O42 'Cap on the right side
    
    Set O3 = Nothing
    Set O31 = Nothing
    Set O32 = Nothing
    Set O4 = Nothing
    Set O40 = Nothing
    Set O41 = Nothing
    Set O42 = Nothing
    'set the "Insulation" representation to the definition
    Set RepsIf = pSymbolDefinition
    RepsIf.SetRepresentation oRep2
    Set oRep2 = Nothing
    Set RepsIf = Nothing
    
    Set oRep2Outputs = Nothing
    'Define the evaluation for the Insulation representation
    Set SymbolicRepEval = New DRepresentationEvaluation
    SymbolicRepEval.Name = "Insulation"
    SymbolicRepEval.Description = "Insualtion representation of flanged nozzle"
    SymbolicRepEval.Properties = igREPRESENTATION_HIDDEN
    SymbolicRepEval.Type = igREPRESENTATION_VBFUNCTION
    SymbolicRepEval.ProgId = "GSCADNozzleVBsymb.HvacNozzle"
    'Set the evaluations on symbol definition
    Set RepsEvalsIf = pSymbolDefinition
    RepsEvalsIf.AddRepresentationEvaluation SymbolicRepEval
    Set RepsEvalsIf = Nothing

Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub

Private Sub InsulateHVACNozzle(ByRef arrayOfInputs())
On Error GoTo ErrorHandler
    Dim Thickness As Double
    Dim FlangeWidth As Double
    Dim Width As Double
    Dim Depth As Double
    Dim CornerRadius As Double
    Dim DimensionBaseOuter As Boolean
    Dim HvacLength As Double
    Dim CrossSectionShape As OutfittingCrossSectionsFacelets.CrossSectionShapeTypes
    Dim CptOffset As Double
    Dim PortDepth As Double
    Dim dT As Double 'Insulation Thickness
    Dim dInsulationCriteria As Long
    Dim oErrors As IJEditErrors
    Dim oLinestr As IngrGeom3D.LineString3d
    Dim PI As Double
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    
    CrossSectionShape = arrayOfInputs(1)
    Thickness = arrayOfInputs(2)
    FlangeWidth = arrayOfInputs(3)
    Width = arrayOfInputs(4)
    Depth = arrayOfInputs(5)
    CornerRadius = arrayOfInputs(6)
    DimensionBaseOuter = arrayOfInputs(7)
    HvacLength = arrayOfInputs(8)
    CptOffset = arrayOfInputs(9)
    PortDepth = arrayOfInputs(10)
    dInsulationCriteria = arrayOfInputs(11) 'Insulation Criteria
    dT = arrayOfInputs(12) 'Insulation Thickness
    
    'if there is no insulation thick ness then better we exit from here itself so that we
    'will remove all the existing outputs from the symbol machinery minimum of 0.1 mm insulation thickness
    'should exist
    If dT <= TOLERANCE Then
        CreateDummyOutput ((CptOffset - PortDepth + Thickness))
        Exit Sub
    End If
    'we donot insulate the nozzle if the insulation criteria is set to none
    If (dInsulationCriteria < 1) Then
        CreateDummyOutput ((CptOffset - PortDepth + Thickness))
        Exit Sub
    End If
    
    Set oErrors = New IMSErrorLog.JServerErrors
    'This is good to get the value of PI use Math call, donot use aTan_1 of 1 * 4
    PI = oGeomFactory.GeometryServices.GeomPI
    If (CornerRadius < 0#) Then
       oErrors.Add E_FAIL, "GSCADNozzleVBsymb.HvacNozzle", "Corner Radius should be Greater than or Equal to zero", "CornerRadius > or = Zero"
       GoTo ErrorHandler
    End If
    
    If (CornerRadius > (Width / 2)) Or (CornerRadius > (Depth / 2)) Then
       oErrors.Add E_FAIL, "GSCADNozzleVBsymb.HvacNozzle", "Corner Radius Should not be greater than Half of Width/Depth", "CornerRadius > Width/Depth/2"
       GoTo ErrorHandler
    End If
    'based on the cross section we try to insulate the nozzle
    
    Select Case CrossSectionShape
        Case OutfittingCrossSectionsFacelets.Rectangular
            'For the Time being create a projection and open on either sides.
            'Neck details
            If (HvacLength > TOLERANCE) Then
                Dim ONeckInsul As Object
                Dim oNeckBoundaryOuter As Object
                Dim oNeckBoundaryInner As Object
                Dim oNeckCapOuter As Object
                Dim oNeckCapInner As Object
                'Add the outer projection
                Set oNeckBoundaryOuter = CreateBoundaryRoundedRectangle((CptOffset - PortDepth + Thickness), _
                                                    (Width), (Depth), _
                                                    0#, dT, CornerRadius)
                Set ONeckInsul = CreateProjectionGivenBoundary(oNeckBoundaryOuter, HvacLength)
                m_outputColl.AddOutput "InsulatedNeckOuter", ONeckInsul
                Set ONeckInsul = Nothing
                'Add the inner Projection
                Set oNeckBoundaryInner = CreateBoundaryRoundedRectangle((CptOffset - PortDepth + Thickness), _
                                                    Width, Depth, _
                                                    0#, 0#, CornerRadius)
                Set ONeckInsul = CreateProjectionGivenBoundary(oNeckBoundaryInner, HvacLength)
                m_outputColl.AddOutput "InsulatedNeckInner", ONeckInsul
                'Create the Caps near Neck Start Cap1
                Set oNeckCapInner = CreateCapGivenboundaries(oNeckBoundaryOuter, oNeckBoundaryInner)
                m_outputColl.AddOutput "InsulatedNeckCap1", oNeckCapInner
                Set oNeckBoundaryOuter = Nothing
                Set oNeckBoundaryInner = Nothing
                'Create the cap near neck end Cap2
                Set oNeckBoundaryOuter = CreateBoundaryRoundedRectangle((CptOffset - PortDepth + Thickness + HvacLength), _
                                                    (Width), (Depth), _
                                                    0#, dT, CornerRadius)
                Set oNeckBoundaryInner = CreateBoundaryRoundedRectangle((CptOffset - PortDepth + Thickness + HvacLength), _
                                                    Width, Depth, _
                                                    0#, 0#, CornerRadius)
                Set oNeckCapOuter = CreateCapGivenboundaries(oNeckBoundaryOuter, oNeckBoundaryInner)
                m_outputColl.AddOutput "InsulatedNeckCap2", oNeckCapOuter
            End If
            
            If (dInsulationCriteria > 1) Then
                'Insulate the Flange
                'Flange Insulation happens in 2 ways
                '1) Insulation material cicumference on Flange Thickness.
                '2) Insulation material on the back side of the flange
                If (Thickness > TOLERANCE) Then
                    '1) Insulation material cicumference on Flange Thickness.
                    Dim OFlngInsul As Object
                    Dim oFlngBoundaryOuter As Object
                    Dim oFlngBoundaryInner As Object
                    Dim oFlngCapOuter As Object
                    Dim oFlngCapInner As Object
                    'Add the outer projection
                    Set oFlngBoundaryOuter = CreateBoundaryRoundedRectangle((CptOffset - PortDepth), _
                                                        (Width), (Depth), _
                                                        FlangeWidth, dT, CornerRadius)
                    Set OFlngInsul = CreateProjectionGivenBoundary(oFlngBoundaryOuter, (Thickness + dT))
                    m_outputColl.AddOutput "InsulatedFlangeOuter", OFlngInsul
                    Set OFlngInsul = Nothing
                    'Add the inner Projection
                    Set oFlngBoundaryInner = CreateBoundaryRoundedRectangle((CptOffset - PortDepth), _
                                                        Width, Depth, _
                                                        FlangeWidth, 0#, CornerRadius)
                    Set OFlngInsul = CreateProjectionGivenBoundary(oFlngBoundaryInner, Thickness)
                    m_outputColl.AddOutput "InsulatedFlangeInner", OFlngInsul
                    'Cap the flang on the front face
                    Set oFlngCapOuter = CreateCapGivenboundaries(oFlngBoundaryOuter, oFlngBoundaryInner)
                    m_outputColl.AddOutput "InsulatedFlangeCap1", oFlngCapOuter
                    Set oFlngBoundaryOuter = Nothing
                    'Cap the flange on the back face.
                    'This can be a single surface without any inner countour
                    Set oFlngBoundaryOuter = CreateBoundaryRoundedRectangle((CptOffset - PortDepth + Thickness + dT), _
                                                        (Width), (Depth), _
                                                        FlangeWidth, dT, CornerRadius)
                    Set oFlngCapInner = CreateCapGivenboundaries(oFlngBoundaryOuter, Nothing)
                    m_outputColl.AddOutput "InsulatedFlangeCap2", oFlngCapInner
                End If
            End If
        Case OutfittingCrossSectionsFacelets.FlatOval
            'Flat Oval follows similar pattern of round-rectangular.
            If (HvacLength > TOLERANCE) Then
                Dim oInsulProj As Object
                Dim oBaseNeckBoundary As Object
                Set oBaseNeckBoundary = CreateBoundaryFlatOval((CptOffset - PortDepth + Thickness), _
                                                            Width, Depth, _
                                                            0#, 0#, 0#)
                'Create a projection using the base boundary
                
                Set oInsulProj = CreateProjectionGivenBoundary(oBaseNeckBoundary, HvacLength)
                m_outputColl.AddOutput "InsulatedNeckInner", oInsulProj
                'Now create the outer boundary
                Dim oOuterNeckBoundary As Object
                Set oOuterNeckBoundary = CreateBoundaryFlatOval((CptOffset - PortDepth + Thickness), _
                                                            Width, Depth, _
                                                            0#, dT, 0#)
                Set oInsulProj = Nothing
                Set oInsulProj = CreateProjectionGivenBoundary(oOuterNeckBoundary, HvacLength)
                m_outputColl.AddOutput "InsulatedNeckOuter", oInsulProj
                'Cap the surfaces on first side
                Dim oCap1 As Object
                Set oCap1 = CreateCapGivenboundaries(oOuterNeckBoundary, oBaseNeckBoundary)
                m_outputColl.AddOutput "InsulatedNeckCap1", oCap1
                'Create Cap2
                Set oBaseNeckBoundary = Nothing
                Set oBaseNeckBoundary = CreateBoundaryFlatOval((CptOffset - PortDepth + Thickness + HvacLength), _
                                                            Width, Depth, _
                                                            0#, 0#, 0#)
                Set oOuterNeckBoundary = Nothing
                Set oOuterNeckBoundary = CreateBoundaryFlatOval((CptOffset - PortDepth + Thickness + HvacLength), _
                                                            Width, Depth, _
                                                            0#, dT, 0#)
                Dim oCap2 As Object
                Set oCap2 = CreateCapGivenboundaries(oOuterNeckBoundary, oBaseNeckBoundary)
                m_outputColl.AddOutput "InsulatedNeckCap2", oCap2
            End If
            If (dInsulationCriteria > 1) Then
                If (Thickness > TOLERANCE) Then
                    Dim obaseFlngBoundary As Object
                    Dim oFlngFTOvalInsul As Object
                    Set obaseFlngBoundary = CreateBoundaryFlatOval((CptOffset - PortDepth), _
                                                            Width, Depth, _
                                                            FlangeWidth, 0#, 0#)
                    'create a base porjection to the length of thickness
                    Set oFlngFTOvalInsul = CreateProjectionGivenBoundary(obaseFlngBoundary, Thickness)
                    m_outputColl.AddOutput "InsulatedFlangeInner", oFlngFTOvalInsul
                    Dim oOuterFlngBoundary As Object
                    Set oOuterFlngBoundary = CreateBoundaryFlatOval((CptOffset - PortDepth), _
                                                            Width, Depth, _
                                                            FlangeWidth, dT, 0#)
                    Set oFlngFTOvalInsul = Nothing
                    Set oFlngFTOvalInsul = CreateProjectionGivenBoundary(oOuterFlngBoundary, (Thickness + dT))
                    m_outputColl.AddOutput "InsulatedFlangeOuter", oFlngFTOvalInsul

                    Dim oCapFltOval As Object
                    Set oCapFltOval = CreateCapGivenboundaries(oOuterFlngBoundary, obaseFlngBoundary)
                    m_outputColl.AddOutput "InsulatedFlangeCap1", oCapFltOval
                    'Make the cap at the back face of flange. this cap need not have inner boundary
                    Set oOuterFlngBoundary = Nothing
                    Set oOuterFlngBoundary = CreateBoundaryFlatOval((CptOffset - PortDepth + Thickness + dT), _
                                                            Width, Depth, _
                                                            FlangeWidth, dT, 0#)
                    Set oCapFltOval = Nothing
                    Set oCapFltOval = CreateCapGivenboundaries(oOuterFlngBoundary, Nothing)
                    m_outputColl.AddOutput "InsulatedFlangeCap2", oCapFltOval
                End If
            End If
        Case OutfittingCrossSectionsFacelets.Round
            If (HvacLength > TOLERANCE) Then
                Dim oNeckInsulGeom As IngrGeom3D.Revolution3d
                Dim dCornerpts(0 To 14) As Double
                
                dCornerpts(0) = CptOffset - PortDepth + Thickness
                dCornerpts(1) = -1# * (Width * 0.5)
                dCornerpts(2) = 0#
                
                dCornerpts(3) = CptOffset - PortDepth + Thickness + HvacLength
                dCornerpts(4) = -1# * (Width * 0.5)
                dCornerpts(5) = 0#
                
                dCornerpts(6) = CptOffset - PortDepth + Thickness + HvacLength
                dCornerpts(7) = -1# * ((Width * 0.5) + dT)
                dCornerpts(8) = 0#
                
                dCornerpts(9) = CptOffset - PortDepth + Thickness
                dCornerpts(10) = -1# * ((Width * 0.5) + dT)
                dCornerpts(11) = 0#
                
                dCornerpts(12) = dCornerpts(0)
                dCornerpts(13) = dCornerpts(1)
                dCornerpts(14) = dCornerpts(2)
                
                Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dCornerpts)
                'Create the surface of Revolution using the same line string around Y axis
                Set oNeckInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, _
                                                                     oLinestr, _
                                                                     1#, 0#, 0#, _
                                                                     0#, _
                                                                     0#, 0#, (PI * 2#), True)
                m_outputColl.AddOutput "InsulatedNeckOuter", oNeckInsulGeom
            End If
            If (dInsulationCriteria > 1) Then
                If (Thickness > TOLERANCE) Then
                    'here the curve is a complicated curve of 6 sides there by making our points to 7
                    Dim dFlngCornerpts(0 To 20) As Double
                    Dim oFlngInsulGeom As IngrGeom3D.Revolution3d
                    
                    dFlngCornerpts(0) = CptOffset - PortDepth
                    dFlngCornerpts(1) = (Width * 0.5) + FlangeWidth
                    dFlngCornerpts(2) = 0#
                    
                    dFlngCornerpts(3) = CptOffset - PortDepth
                    dFlngCornerpts(4) = (Width * 0.5) + FlangeWidth + dT
                    dFlngCornerpts(5) = 0#
                    
                    dFlngCornerpts(6) = CptOffset - PortDepth + Thickness + dT
                    dFlngCornerpts(7) = (Width * 0.5) + FlangeWidth + dT
                    dFlngCornerpts(8) = 0#
                    
                    dFlngCornerpts(9) = CptOffset - PortDepth + Thickness + dT
                    dFlngCornerpts(10) = (Width * 0.5)
                    dFlngCornerpts(11) = 0#
                    
                    dFlngCornerpts(12) = CptOffset - PortDepth + Thickness
                    dFlngCornerpts(13) = (Width * 0.5)
                    dFlngCornerpts(14) = 0#
                    
                    dFlngCornerpts(15) = CptOffset - PortDepth + Thickness
                    dFlngCornerpts(16) = (Width * 0.5) + FlangeWidth
                    dFlngCornerpts(17) = 0#
                    
                    dFlngCornerpts(18) = dFlngCornerpts(0)
                    dFlngCornerpts(19) = dFlngCornerpts(1)
                    dFlngCornerpts(20) = dFlngCornerpts(2)
                    
                    'Construct the Line string with above points
                    Set oLinestr = Nothing
                    Set oLinestr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dFlngCornerpts)
                    'Create the surface of Revolution using the same line string around Y axis
                    Set oFlngInsulGeom = oGeomFactory.Revolutions3d.CreateByCurve(m_outputColl.ResourceManager, _
                                                                        oLinestr, _
                                                                        1#, 0#, 0#, _
                                                                        0#, 0#, 0#, (PI * 2#), True)
                    m_outputColl.AddOutput "InsulatedFlangeOuter", oFlngInsulGeom
                End If
            End If
        Case Else
        'we donot support any other criteria.
    End Select
Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub


Private Function CreateProjectionGivenBoundary(ByVal ObjCurve As Object, ByVal Length As Double) As Object
On Error GoTo ErrorHandler
Set CreateProjectionGivenBoundary = Nothing
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim projection As IngrGeom3D.Projection3d
    Dim projVecX As Double, projVecY As Double, projVecZ As Double
    projVecX = 1#
    projVecY = 0#
    projVecZ = 0#
    ' Create projection
    Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, ObjCurve, _
                                    projVecX, projVecY, projVecZ, Length, False)
    Set CreateProjectionGivenBoundary = projection
    
Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

Private Function CreateCapGivenboundaries(ByVal ObjCurveOuter As Object, ByVal ObjCurveInner As Object) As Object
On Error GoTo ErrorHandler
Set CreateCapGivenboundaries = Nothing
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim projection As IngrGeom3D.Projection3d
    Dim projVecX As Double, projVecY As Double, projVecZ As Double
    projVecX = 1#
    projVecY = 0#
    projVecZ = 0#
    'Create the Plane
    Dim oPlane As IngrGeom3D.Plane3d
    Dim cmplxStr As IngrGeom3D.ComplexString3d
    Set oPlane = geomFactory.Planes3d.CreateByOuterBdry(m_outputColl.ResourceManager, ObjCurveOuter)
    If Not oPlane Is Nothing Then
        If Not ObjCurveInner Is Nothing Then
            Set cmplxStr = ObjCurveInner
            oPlane.AddHole cmplxStr
        End If
    End If
    Set CreateCapGivenboundaries = oPlane
Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

Private Function CreateBoundaryRoundedRectangle(ByVal dInitPos As Double, _
                                ByVal Width As Double, _
                                ByVal Depth As Double, _
                                ByVal FlngWidth As Double, _
                                ByVal dInsulationThk As Double, _
                                ByVal dCornerRad As Double) As Object
On Error GoTo ErrorHandler
Set CreateBoundaryRoundedRectangle = Nothing
    'Dimensional Positional details
    Dim X As Double
    Dim Ymin As Double, Zmin As Double
    Dim Ymax As Double, Zmax As Double
    Dim Radius As Double

    Dim complexStr As IngrGeom3D.ComplexString3d
    Dim line As IngrGeom3D.Line3d
    Dim arc As IngrGeom3D.Arc3d
    Dim coll As IJElements
    Dim elem As Object
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    
    Dim P1(0 To 2) As Double, P2(0 To 2) As Double, P3(0 To 2) As Double
    Dim P4(0 To 2) As Double, P5(0 To 2) As Double, P6(0 To 2) As Double
    Dim P7(0 To 2) As Double, P8(0 To 2) As Double
    Dim C1(0 To 2) As Double, C2(0 To 2) As Double
    Dim C3(0 To 2) As Double, C4(0 To 2) As Double
    Set coll = New JObjectCollection


    X = dInitPos
    Ymin = -(Depth / 2) - dInsulationThk - FlngWidth
    Zmin = -(Width / 2) - dInsulationThk - FlngWidth
    Ymax = (Depth / 2) + dInsulationThk + FlngWidth
    Zmax = (Width / 2) + dInsulationThk + FlngWidth
    If dCornerRad > 0# Then
        Radius = dCornerRad + dInsulationThk + FlngWidth
    Else
        Radius = 0#
    End If
    'Construct the complex curve
    P1(0) = X
    P1(1) = Ymin + Radius
    P1(2) = Zmin
    P2(0) = X
    P2(1) = Ymax - Radius
    P2(2) = Zmin
    P3(0) = X
    P3(1) = Ymax
    P3(2) = Zmin + Radius
    P4(0) = X
    P4(1) = Ymax
    P4(2) = Zmax - Radius
    P5(0) = X
    P5(1) = Ymax - Radius
    P5(2) = Zmax
    P6(0) = X
    P6(1) = Ymin + Radius
    P6(2) = Zmax
    P7(0) = X
    P7(1) = Ymin
    P7(2) = Zmax - Radius
    P8(0) = X
    P8(1) = Ymin
    P8(2) = Zmin + Radius

    C1(0) = X
    C1(1) = Ymin + Radius
    C1(2) = Zmin + Radius
    C2(0) = X
    C2(1) = Ymax - Radius
    C2(2) = Zmin + Radius
    C3(0) = X
    C3(1) = Ymax - Radius
    C3(2) = Zmax - Radius
    C4(0) = X
    C4(1) = Ymin + Radius
    C4(2) = Zmax - Radius
    'Construct individual lines and arcs and add to the complex string
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, P1(0), P1(1), P1(2), P2(0), P2(1), P2(2))
    coll.Add line
    Set line = Nothing
    
    If Radius > 0# Then
        Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, C2(0), C2(1), C2(2), P2(0), P2(1), P2(2), P3(0), P3(1), P3(2))
        coll.Add arc
        Set arc = Nothing
    End If
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, P3(0), P3(1), P3(2), P4(0), P4(1), P4(2))
    coll.Add line
    Set line = Nothing
    
    If Radius > 0# Then
        Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, C3(0), C3(1), C3(2), P4(0), P4(1), P4(2), P5(0), P5(1), P5(2))
        coll.Add arc
        Set arc = Nothing
    End If
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, P5(0), P5(1), P5(2), P6(0), P6(1), P6(2))
    coll.Add line
    Set line = Nothing
    
    If Radius > 0# Then
        Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, C4(0), C4(1), C4(2), P6(0), P6(1), P6(2), P7(0), P7(1), P7(2))
        coll.Add arc
        Set arc = Nothing
    End If
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, P7(0), P7(1), P7(2), P8(0), P8(1), P8(2))
    coll.Add line
    Set line = Nothing
    
    If Radius > 0# Then
        Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, C1(0), C1(1), C1(2), P8(0), P8(1), P8(2), P1(0), P1(1), P1(2))
        coll.Add arc
        Set arc = Nothing
    End If
    
    ' Create complex string
    Set complexStr = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, coll)
    'Add the projection here
    Set CreateBoundaryRoundedRectangle = complexStr
    'clear the lines and arcs in collection
    coll.Clear
Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

Private Function CreateBoundaryFlatOval(ByVal dInitPos As Double, _
                                ByVal Width As Double, _
                                ByVal Depth As Double, _
                                ByVal FlngWidth As Double, _
                                ByVal dInsulationThk As Double, _
                                ByVal dCornerRad As Double) As Object
On Error GoTo ErrorHandler
Set CreateBoundaryFlatOval = Nothing
    'Dimensional Positional details
    Dim X As Double
    Dim Ymin As Double, Zmin As Double
    Dim Ymax As Double, Zmax As Double
    Dim Radius As Double

    Dim complexStr As IngrGeom3D.ComplexString3d
    Dim line As IngrGeom3D.Line3d
    Dim arc As IngrGeom3D.Arc3d
    Dim coll As IJElements
    Dim elem As Object
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    
    Dim P1(0 To 2) As Double, P2(0 To 2) As Double, P3(0 To 2) As Double
    Dim P4(0 To 2) As Double, P5(0 To 2) As Double, P6(0 To 2) As Double
    Dim P7(0 To 2) As Double, P8(0 To 2) As Double
    Dim C1(0 To 2) As Double, C2(0 To 2) As Double
    Dim C3(0 To 2) As Double, C4(0 To 2) As Double
    Set coll = New JObjectCollection
    
    Dim dFlatWidth As Double
    dFlatWidth = Width - Depth
    If dFlatWidth < 0# Then dFlatWidth = 0.000001 ' use distance tolerance

    X = dInitPos 'X + Thickness
    Ymin = -((Depth / 2) + dInsulationThk + FlngWidth)
    Zmin = -(dFlatWidth / 2)
    Ymax = (Depth / 2) + dInsulationThk + FlngWidth
    Zmax = (dFlatWidth / 2)
    Radius = (Depth / 2) + dInsulationThk + FlngWidth
    
    P1(0) = X
    P1(1) = Ymin
    P1(2) = Zmin
    P2(0) = X
    P2(1) = Ymax
    P2(2) = Zmin
    P3(0) = X
    P3(1) = Ymax
    P3(2) = Zmax
    P4(0) = X
    P4(1) = Ymin
    P4(2) = Zmax
    
    P5(0) = X
    P5(1) = 0#
    P5(2) = Zmin - Radius
    P6(0) = X
    P6(1) = 0#
    P6(2) = Zmax + Radius
            
    'Construct individual lines and arcs and create a complex string
    
    Set arc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                P1(0), P1(1), P1(2), P5(0), P5(1), P5(2), _
                                P2(0), P2(1), P2(2))
    coll.Add arc
    Set arc = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, P2(0), P2(1), P2(2), P3(0), P3(1), P3(2))
    coll.Add line
    Set line = Nothing
    
    Set arc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                P3(0), P3(1), P3(2), P6(0), P6(1), P6(2), _
                                P4(0), P4(1), P4(2))
    coll.Add arc
    Set arc = Nothing
    
    Set line = geomFactory.Lines3d.CreateBy2Points(Nothing, P4(0), P4(1), P4(2), P1(0), P1(1), P1(2))
    coll.Add line
    Set line = Nothing
    ' Create complex string
    Set complexStr = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, coll)
    Set CreateBoundaryFlatOval = complexStr
    'clear the lines and arcs in collection
    coll.Clear
Exit Function
ErrorHandler:
    Err.Raise E_FAIL
End Function

Private Sub CreateDummyOutput(ByVal dLoc As Double)
On Error GoTo ErrorHandler
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim oPoint As IngrGeom3D.Point3d
    Set oPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, dLoc, 0#, 0#)
    'Add the output to output Collection
    m_outputColl.AddOutput "InsulatedFlangeInner", oPoint
Exit Sub
ErrorHandler:
    Err.Raise E_FAIL
End Sub
