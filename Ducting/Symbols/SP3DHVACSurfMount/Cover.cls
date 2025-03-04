VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CleanBend.cls
'   ProgID:         SP3DHVACSurfMount.Cover
'   Author:         svsmylav
'   Creation Date:  Monday, Apr 24 2006
'   Description:
'       Hvac Surface Mount Cover
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   24.Apr.2006     SymbolTeam(India)       DI-94663  New HVAC symbol is prepared from existing
'                                           GSCAD symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_outputColl As IJDOutputCollection

Implements IJDUserSymbolServices

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 IJDUserSymbolServices_EditOccurence = False
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    FeedNozzleDefinition pSymbolDefinition
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

End Sub

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = "SP3DHVACSurfMount.Cover"
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  On Error GoTo ErrorHandler

  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition

  Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "SP3DHVACSurfMount.Cover"
  oSymbolDefinition.CodeBase = CodeBase

  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId

  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  Set oSymbolDefinition = Nothing
  Set oSymbolFactory = Nothing

  Exit Function

ErrorHandler:
    Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Debug.Assert False
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Set m_outputColl = outputcoll
    If StrComp(repName, "Detailed") = 0 Then
        Symbolic arrayOfInputs
    ElseIf StrComp(repName, "Symbolic") = 0 Then
        Symbolic arrayOfInputs
    End If
End Sub

'Draw a Detailed Nozzle
Private Sub Detailed(ByRef arrayOfInputs())
    
End Sub

Private Sub Symbolic(ByRef arrayOfInputs())
    Dim Thickness As Double, Cover As Double
    Dim holelength As Double, holebreadth As Double
    Dim covercornerradius As Double
    Dim gripdistance As Double
    Dim handlelength As Double
    Dim CrossSectionShape As OutfittingCrossSectionsFacelets.CrossSectionShapeTypes
    
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim projection As IngrGeom3D.Projection3d
    Dim projVecZ As Double, projVecY As Double, projVecX As Double
    Dim gelem As IJDObject
         
    Dim Z As Double
    Dim Ymin As Double, Xmin As Double
    Dim Ymax As Double, Xmax As Double
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
         
    CrossSectionShape = arrayOfInputs(2)
    Thickness = arrayOfInputs(3)
    Cover = arrayOfInputs(4)
    holelength = arrayOfInputs(5)
    holebreadth = arrayOfInputs(6)
    covercornerradius = arrayOfInputs(7)
    gripdistance = arrayOfInputs(8)
    handlelength = arrayOfInputs(9)

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
        
                
        Z = 0#
        Ymin = -((holebreadth / 2) + Cover)
        Xmin = -((holelength / 2) + Cover)
        Ymax = (holebreadth / 2) + Cover
        Xmax = (holelength / 2) + Cover
        
        If CmpDblEqual(covercornerradius, 0) Then
            Radius = 0#
        Else
            Radius = covercornerradius + Cover
        End If
        
        P1(0) = Xmin
        P1(1) = Ymin + Radius
        P1(2) = Z
        P2(0) = Xmin
        P2(1) = Ymax - Radius
        P2(2) = Z
        P3(0) = Xmin + Radius
        P3(1) = Ymax
        P3(2) = Z
        P4(0) = Xmax - Radius
        P4(1) = Ymax
        P4(2) = Z
        P5(0) = Xmax
        P5(1) = Ymax - Radius
        P5(2) = Z
        P6(0) = Xmax
        P6(1) = Ymin + Radius
        P6(2) = Z
        P7(0) = Xmax - Radius
        P7(1) = Ymin
        P7(2) = Z
        P8(0) = Xmin + Radius
        P8(1) = Ymin
        P8(2) = Z
        
        C1(0) = Xmin + Radius
        C1(1) = Ymin + Radius
        C1(2) = Z
        C2(0) = Xmin + Radius
        C2(1) = Ymax - Radius
        C2(2) = Z
        C3(0) = Xmax - Radius
        C3(1) = Ymax - Radius
        C3(2) = Z
        C4(0) = Xmax - Radius
        C4(1) = Ymin + Radius
        C4(2) = Z
        
        ' Create a complex string
        
        Set coll = New JObjectCollection
              
        'Construct individual lines and arcs and create a complex string
        
        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P1(0), P1(1), P1(2), P2(0), P2(1), P2(2))
        coll.Add line
        Set line = Nothing
        
        If CmpDblGreaterthan(Radius, 0) Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, _
                                        C2(0), C2(1), C2(2), P2(0), P2(1), P2(2), _
                                        P3(0), P3(1), P3(2))
            coll.Add arc
            Set arc = Nothing
        End If

        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P3(0), P3(1), P3(2), P4(0), P4(1), P4(2))
        coll.Add line
        Set line = Nothing

        If CmpDblGreaterthan(Radius, 0) Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, _
                                        C3(0), C3(1), C3(2), P4(0), P4(1), P4(2), _
                                        P5(0), P5(1), P5(2))
            coll.Add arc
            Set arc = Nothing
        End If

        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P5(0), P5(1), P5(2), P6(0), P6(1), P6(2))
        coll.Add line
        Set line = Nothing

        If CmpDblGreaterthan(Radius, 0) Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, _
                                        C4(0), C4(1), C4(2), P6(0), P6(1), P6(2), _
                                        P7(0), P7(1), P7(2))
            coll.Add arc
            Set arc = Nothing
        End If

        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P7(0), P7(1), P7(2), P8(0), P8(1), P8(2))
        coll.Add line
        Set line = Nothing

        If CmpDblGreaterthan(Radius, 0) Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, _
                                        C1(0), C1(1), C1(2), P8(0), P8(1), P8(2), _
                                        P1(0), P1(1), P1(2))
            coll.Add arc
            Set arc = Nothing
        End If
        
        ' Create complex string
        Set complexStr = geomFactory.ComplexStrings3d.CreateByCurves(m_outputColl.ResourceManager, coll)
        
        ' Create projection
        
        projVecZ = 1#
        projVecY = 0#
        projVecX = 0#
    
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, complexStr, _
                                        projVecX, projVecY, projVecZ, Thickness, True)
        m_outputColl.AddOutput "Flange", projection
        Set projection = Nothing
        
        ' Delete lines and arcs in collection
        
        For Each elem In coll
            Set gelem = elem
            Debug.Assert Not gelem Is Nothing
            Set elem = Nothing
            gelem.Remove
            Set gelem = Nothing
        Next elem

        coll.Clear
        
'        complexStr.RemoveAll

        'Delete complexString which was created for creating projection.
        Set gelem = complexStr
        Debug.Assert Not gelem Is Nothing
        Set complexStr = Nothing
        gelem.Remove

                
        'CONSTRUCTION OF  Hvac Cross Section
       Z = 0#
        Ymin = -(holebreadth / 2)
        Xmin = -(holelength / 2)
        Ymax = (holebreadth / 2)
        Xmax = (holelength / 2)
        Radius = covercornerradius
        
        P1(0) = Xmin
        P1(1) = Ymin + Radius
        P1(2) = Z
        P2(0) = Xmin
        P2(1) = Ymax - Radius
        P2(2) = Z
        P3(0) = Xmin + Radius
        P3(1) = Ymax
        P3(2) = Z
        P4(0) = Xmax - Radius
        P4(1) = Ymax
        P4(2) = Z
        P5(0) = Xmax
        P5(1) = Ymax - Radius
        P5(2) = Z
        P6(0) = Xmax
        P6(1) = Ymin + Radius
        P6(2) = Z
        P7(0) = Xmax - Radius
        P7(1) = Ymin
        P7(2) = Z
        P8(0) = Xmin + Radius
        P8(1) = Ymin
        P8(2) = Z
        
        C1(0) = Xmin + Radius
        C1(1) = Ymin + Radius
        C1(2) = Z
        C2(0) = Xmin + Radius
        C2(1) = Ymax - Radius
        C2(2) = Z
        C3(0) = Xmax - Radius
        C3(1) = Ymax - Radius
        C3(2) = Z
        C4(0) = Xmax - Radius
        C4(1) = Ymin + Radius
        C4(2) = Z
        
        'Construct individual lines and arcs and add to the complex string
        
        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P1(0), P1(1), P1(2), P2(0), P2(1), P2(2))
        coll.Add line
        Set line = Nothing
        
        If CmpDblGreaterthan(Radius, 0) Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, _
                                        C2(0), C2(1), C2(2), P2(0), P2(1), P2(2), _
                                        P3(0), P3(1), P3(2))
            coll.Add arc
            Set arc = Nothing
        End If
        
        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P3(0), P3(1), P3(2), P4(0), P4(1), P4(2))
        coll.Add line
        Set line = Nothing
        
        If CmpDblGreaterthan(Radius, 0) Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, _
                                        C3(0), C3(1), C3(2), P4(0), P4(1), P4(2), _
                                        P5(0), P5(1), P5(2))
            coll.Add arc
            Set arc = Nothing
        End If
        
        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P5(0), P5(1), P5(2), P6(0), P6(1), P6(2))
        coll.Add line
        Set line = Nothing
        
        If CmpDblGreaterthan(Radius, 0) Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, _
                                        C4(0), C4(1), C4(2), P6(0), P6(1), P6(2), _
                                        P7(0), P7(1), P7(2))
            coll.Add arc
            Set arc = Nothing
        End If
        
        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P7(0), P7(1), P7(2), P8(0), P8(1), P8(2))
        coll.Add line
        Set line = Nothing
        
        If CmpDblGreaterthan(Radius, 0) Then
            Set arc = geomFactory.Arcs3d.CreateByCenterStartEnd(m_outputColl.ResourceManager, _
                                        C1(0), C1(1), C1(2), P8(0), P8(1), P8(2), _
                                        P1(0), P1(1), P1(2))
            coll.Add arc
            Set arc = Nothing
        End If
        
        ' Create complex string
        Set complexStr = geomFactory.ComplexStrings3d.CreateByCurves(m_outputColl.ResourceManager, coll)
        
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, complexStr, _
                                        projVecX, projVecY, projVecZ, 0.0001, True)
        m_outputColl.AddOutput "Hvac", projection
        Set projection = Nothing
        
        ' Delete lines and arcs in collection
        For Each elem In coll
            Set gelem = elem
            Debug.Assert Not gelem Is Nothing
            Set elem = Nothing
            gelem.Remove
            Set gelem = Nothing
        Next elem
'        For i = coll.Count To 1 Step -1
'            Set gelem = coll.Item(i)
'            coll.Remove i
'            Debug.Assert Not gelem Is Nothing
'            gelem.Remove
'            Set gelem = Nothing
'        Next i
        
        coll.Clear
        
'        complexStr.RemoveAll
        'Delete complexStr which was created for creating projection.
        Set gelem = complexStr
        Debug.Assert Not gelem Is Nothing
        Set complexStr = Nothing
        gelem.Remove
    
        
        Set geomFactory = Nothing

    Case OutfittingCrossSectionsFacelets.FlatOval
    '*************************************************************************
    '           FlatOval
    '*************************************************************************
    '
    '                 P2----------------P3
    '                 *                   *
    '            P5 *                       *P6
    '                 *                   *
    '                 P1----------------P4
    '
       Z = 0#
       Radius = (holebreadth / 2) + Cover
        Ymin = -((holebreadth / 2) + Cover)
         Xmin = -((holelength / 2) - (holebreadth / 2))
        Ymax = (holebreadth / 2) + Cover
         Xmax = (holelength / 2) - (holebreadth / 2)
        
        
        P1(2) = Z
        P1(1) = Ymin
        P1(0) = Xmin
        P2(2) = Z
        P2(1) = Ymax
        P2(0) = Xmin
        P3(2) = Z
        P3(1) = Ymax
        P3(0) = Xmax
        P4(2) = Z
        P4(1) = Ymin
        P4(0) = Xmax
        
        P5(2) = Z
        P5(1) = 0#
        P5(0) = Xmin - Radius
        P6(2) = Z
        P6(1) = 0#
        P6(0) = Xmax + Radius
                
        Set coll = New JObjectCollection
              
        'Construct individual lines and arcs and create a complex string
        
        Set arc = geomFactory.Arcs3d.CreateBy3Points(m_outputColl.ResourceManager, _
                                    P1(0), P1(1), P1(2), P5(0), P5(1), P5(2), _
                                    P2(0), P2(1), P2(2))
        coll.Add arc
        Set arc = Nothing
        
        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P2(0), P2(1), P2(2), P3(0), P3(1), P3(2))
        coll.Add line
        Set line = Nothing
        
        Set arc = geomFactory.Arcs3d.CreateBy3Points(m_outputColl.ResourceManager, _
                                    P3(0), P3(1), P3(2), P6(0), P6(1), P6(2), _
                                    P4(0), P4(1), P4(2))
        coll.Add arc
        Set arc = Nothing
        
        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P4(0), P4(1), P4(2), P1(0), P1(1), P1(2))
        coll.Add line
        Set line = Nothing

        
        ' Create complex string
        Set complexStr = geomFactory.ComplexStrings3d.CreateByCurves(m_outputColl.ResourceManager, coll)
        
        ' Create projection
        
        projVecZ = 1#
        projVecY = 0#
        projVecX = 0#
    
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, complexStr, _
                                        projVecX, projVecY, projVecZ, Thickness, True)
        m_outputColl.AddOutput "Flange", projection
        Set projection = Nothing
          
        ' Delete lines and arcs in collection
        For Each elem In coll
            Set gelem = elem
            Debug.Assert Not gelem Is Nothing
            Set elem = Nothing
            gelem.Remove
            Set gelem = Nothing
        Next elem
        
        coll.Clear
      
        'Delete complexString which was created for creating projection.
        Set gelem = complexStr
        Debug.Assert Not gelem Is Nothing
        Set complexStr = Nothing
        gelem.Remove
        
                 
        'CONSTRUCTION OF  Hvac Cross Section
       Z = 0#
        Ymin = -(holebreadth / 2)
        Xmin = -((holelength / 2) - (holebreadth / 2))
        Ymax = (holebreadth / 2)
        Xmax = (holelength / 2) - (holebreadth / 2)
        Radius = holebreadth / 2
        
        P1(2) = Z
        P1(1) = Ymin
        P1(0) = Xmin
        P2(2) = Z
        P2(1) = Ymax
        P2(0) = Xmin
        P3(2) = Z
        P3(1) = Ymax
        P3(0) = Xmax
        P4(2) = Z
        P4(1) = Ymin
        P4(0) = Xmax
        
        P5(2) = Z
        P5(1) = 0#
        P5(0) = Xmin - Radius
        P6(2) = Z
        P6(1) = 0#
        P6(0) = Xmax + Radius
                
        'Construct individual lines and arcs and create a complex string
        
        Set arc = geomFactory.Arcs3d.CreateBy3Points(m_outputColl.ResourceManager, _
                                    P1(0), P1(1), P1(2), P5(0), P5(1), P5(2), _
                                    P2(0), P2(1), P2(2))
        coll.Add arc
        Set arc = Nothing
        
        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P2(0), P2(1), P2(2), P3(0), P3(1), P3(2))
        coll.Add line
        Set line = Nothing
        
        Set arc = geomFactory.Arcs3d.CreateBy3Points(m_outputColl.ResourceManager, _
                                    P3(0), P3(1), P3(2), P6(0), P6(1), P6(2), _
                                    P4(0), P4(1), P4(2))
        coll.Add arc
        Set arc = Nothing
        
        Set line = geomFactory.Lines3d.CreateBy2Points(m_outputColl.ResourceManager, _
                                    P4(0), P4(1), P4(2), P1(0), P1(1), P1(2))
        coll.Add line
        Set line = Nothing

        
        ' Create complex string
        Set complexStr = geomFactory.ComplexStrings3d.CreateByCurves(m_outputColl.ResourceManager, coll)
        
        ' Create projection
        
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, complexStr, _
                                        projVecX, projVecY, projVecZ, 0.0001, True)
        m_outputColl.AddOutput "Hvac", projection
        Set projection = Nothing
        
        ' Delete lines and arcs in collection
        For Each elem In coll
            Set gelem = elem
            Debug.Assert Not gelem Is Nothing
            Set elem = Nothing
            gelem.Remove
            Set gelem = Nothing
        Next elem
        
        coll.Clear
        
        'Delete complexString which was created for creating projection.
        Set gelem = complexStr
        Debug.Assert Not gelem Is Nothing
        Set complexStr = Nothing
        gelem.Remove
        
        Set geomFactory = Nothing
        
    Case OutfittingCrossSectionsFacelets.Round
    '*************************************************************************
    '           ROUND
    '*************************************************************************
        'CONSTRUCTION OF  Flange Body Cylinder
        Dim circleCenterZ As Double, circleCenterY As Double, circleCenterX As Double
        circleCenterZ = 0#
        circleCenterY = 0#
        circleCenterX = 0#
    
        Dim circleNormalZ As Double, circleNormalY As Double, circleNormalX As Double
        circleNormalZ = 1#
        circleNormalY = 0#
        circleNormalX = 0#
    
        Dim circlePointVecZ As Double, circlePointVecY As Double, circlePointVecX As Double
        circlePointVecZ = 0#
        circlePointVecY = (holelength * 0.5) + Cover
        circlePointVecX = 0#
    
        Dim axesRatio As Double
        axesRatio = 1#
    
        Dim ellipse As IngrGeom3D.Ellipse3d
        Set ellipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(m_outputColl.ResourceManager, _
                                    circleCenterX, circleCenterY, circleCenterZ, _
                                    circleNormalX, circleNormalY, circleNormalZ, _
                                    circlePointVecX, circlePointVecY, circlePointVecZ, _
                                    axesRatio)
    
        projVecZ = 1#
        projVecY = 0#
        projVecX = 0#
    
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, ellipse, _
                                        projVecX, projVecY, projVecZ, Thickness, True)
        m_outputColl.AddOutput "Flange", projection
        Set projection = Nothing
    
        'Delete ellipse which was created for creating projection.
        Set gelem = ellipse
        Debug.Assert Not gelem Is Nothing
        Set ellipse = Nothing
        gelem.Remove
    
        'CONSTRUCTION OF  Hvac Cylinder
        circleCenterZ = 0#
        circlePointVecY = holelength * 0.5
        Set ellipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(m_outputColl.ResourceManager, _
                                    circleCenterX, circleCenterY, circleCenterZ, _
                                    circleNormalX, circleNormalY, circleNormalZ, _
                                    circlePointVecX, circlePointVecY, circlePointVecZ, _
                                    axesRatio)
    
        Set projection = geomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, ellipse, _
                                        projVecX, projVecY, projVecZ, 0.0001, True)
        m_outputColl.AddOutput "Hvac", projection
        Set projection = Nothing
    
        'Delete ellipse which was created for creating projection.
        Set gelem = ellipse
        Debug.Assert Not gelem Is Nothing
        Set ellipse = Nothing
        gelem.Remove
    
        Set geomFactory = Nothing
    Case Else
    
    End Select
    
End Sub

' Feed Nozzle Definition
' Inputs:
'           "CrossSectionShape"
'           "Thickness"
'           "cover"
'           "holelength"
'           "holebreadth"
'           "covercornerradius"
'           "gripdistance"
'           "handlelength"
' The representations "Symbolic" and "Detailed" are set to the definition
' Symbolic :
' Detailed :

Sub FeedNozzleDefinition(pSymbolDefinition As IJDSymbolDefinition)
  
  On Error GoTo ErrorHandler
    'There should be one input defined for refData part.
    'This should be the first input
   
  ' Create a new input by new operator
      Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
   'Fix for TR 43660
   Dim oSymbolCache As New CustomCache
   oSymbolCache.SetupCustomCache pSymbolDefinition

  Dim Inputs(1 To 8) As IMSSymbolEntities.DInput
'    Dim oPart As New IMSSymbolEntities.DInput
'    oPart.Name = "Part"
'    oPart.Description = "Part"
'      InputsIf.SetInput oPart, 1
'    Set oPart = Nothing
 
  ' Create a defaultValue
  Dim PC As IMSSymbolEntities.DParameterContent
  Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
  
  PC.Type = igValue

  Set Inputs(1) = New IMSSymbolEntities.DInput
  Inputs(1).Name = "CrossSectionShape"
  Inputs(1).Description = "Shape of the CrossSection"
  Inputs(1).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = OutfittingCrossSectionsFacelets.Rectangular
  Inputs(1).DefaultParameterValue = PC
  
  Set Inputs(2) = New IMSSymbolEntities.DInput
  Inputs(2).Name = "Thickness"
  Inputs(2).Description = "Thickness of the Flange"
  Inputs(2).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.03
  Inputs(2).DefaultParameterValue = PC
  
  Set Inputs(3) = New IMSSymbolEntities.DInput
  Inputs(3).Name = "cover"
  Inputs(3).Description = "Incremental width of the flange"
  Inputs(3).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.075
  Inputs(3).DefaultParameterValue = PC
  
  Set Inputs(4) = New IMSSymbolEntities.DInput
  Inputs(4).Name = "holelength"
  Inputs(4).Description = "holelength of the crosssection"
  Inputs(4).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.56
  Inputs(4).DefaultParameterValue = PC
  
  Set Inputs(5) = New IMSSymbolEntities.DInput
  Inputs(5).Name = "holebreadth"
  Inputs(5).Description = "holebreadth of the crosssection"
  Inputs(5).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.36
  Inputs(5).DefaultParameterValue = PC
  
  Set Inputs(6) = New IMSSymbolEntities.DInput
  Inputs(6).Name = "covercornerradius"
  Inputs(6).Description = "Corner Radius of the crosssection"
  Inputs(6).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0#
  Inputs(6).DefaultParameterValue = PC
  
  Set Inputs(7) = New IMSSymbolEntities.DInput
  Inputs(7).Name = "gripdistance"
  Inputs(7).Description = "Height of the cover Handle "
  Inputs(7).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.4
  Inputs(7).DefaultParameterValue = PC
  
  Set Inputs(8) = New IMSSymbolEntities.DInput
  Inputs(8).Name = "handlelength"
  Inputs(8).Description = "Length of the Cover Handle"
  Inputs(8).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.2
  Inputs(8).DefaultParameterValue = PC
  
  'Set the input to the definition
  'Dim InputsIf As IMSSymbolEntities.IJDInputs
  'Set InputsIf = pSymbolDefinition

  Dim index As Integer
  For index = 1 To 8
    InputsIf.SetInput Inputs(index), index + 1
  Next

  'Create the output
  Dim O1 As IMSSymbolEntities.DOutput
  Set O1 = New IMSSymbolEntities.DOutput
  
  O1.Name = "Flange"
  O1.Description = "Flange of Hvac Nozzle"
  O1.Properties = 0
  
  Dim O2 As IMSSymbolEntities.DOutput
  Set O2 = New IMSSymbolEntities.DOutput

  O2.Name = "Hvac"
  O2.Description = "Hvac of Nozzle"
  O2.Properties = 0

  'Define the representation "Symbolic"
  Dim rep1 As IMSSymbolEntities.DRepresentation
  Set rep1 = New IMSSymbolEntities.DRepresentation

  rep1.Name = "Symbolic"
  rep1.Description = "Representation of Nozzle is Symbolic"
  rep1.Properties = igREPRESENTATION_ISVBFUNCTION
  'rep1.Properties = igREPRESENTATION_ISSCRIPTED

  Dim oRepPhysicalOutputs As IMSSymbolEntities.IJDOutputs
  Set oRepPhysicalOutputs = rep1 '
        
  'Set the outputs
  oRepPhysicalOutputs.SetOutput O1
  oRepPhysicalOutputs.SetOutput O2
  Set O1 = Nothing
  Set O2 = Nothing

  'Set the 'symbolic' representation to definition
  Dim RepsIf As IMSSymbolEntities.IJDRepresentations
  Set RepsIf = pSymbolDefinition
  RepsIf.SetRepresentation rep1

  'Define another representation "Detailed" and set it on the definition
  rep1.Name = "Detailed"
  rep1.Description = "Detailed Represntation of the Nozzle"
  rep1.Properties = igREPRESENTATION_ISVBFUNCTION

  RepsIf.SetRepresentation rep1

  Set rep1 = Nothing
  Set RepsIf = Nothing

  'Define the evaluation for the Symbolic representation
  Dim SymbolicRepEval As DRepresentationEvaluation
  Set SymbolicRepEval = New DRepresentationEvaluation
  SymbolicRepEval.Name = "Symbolic"
  SymbolicRepEval.Description = "Symbolic representation of flanged nozzle"
  SymbolicRepEval.Properties = igREPRESENTATION_HIDDEN
  SymbolicRepEval.Type = igREPRESENTATION_VBFUNCTION
  SymbolicRepEval.ProgId = "SP3DHVACSurfMount.Cover"
                                                

  'Define the evaluation for the Detailed representation
  Dim DetailedRepEval As DRepresentationEvaluation
  Set DetailedRepEval = New DRepresentationEvaluation
  DetailedRepEval.Name = "Detailed"
  DetailedRepEval.Description = "Detailed representation of flanged nozzle"
  DetailedRepEval.Properties = igREPRESENTATION_HIDDEN
  DetailedRepEval.Type = igREPRESENTATION_VBFUNCTION
  DetailedRepEval.ProgId = "SP3DHVACSurfMount.Cover"

  'Set the evaluations on symbol definition
  Dim RepsEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
  Set RepsEvalsIf = pSymbolDefinition

  RepsEvalsIf.AddRepresentationEvaluation SymbolicRepEval
  RepsEvalsIf.AddRepresentationEvaluation DetailedRepEval

  Set RepsEvalsIf = Nothing
  Set oRepPhysicalOutputs = Nothing
  Exit Sub

ErrorHandler:
  Debug.Print Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
  Debug.Assert False

End Sub


