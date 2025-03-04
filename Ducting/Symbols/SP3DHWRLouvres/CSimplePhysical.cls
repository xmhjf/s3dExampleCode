VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright(c)2005-08 Intergraph GmbH
'   All Rights Reserved
'
'  "SimplePhysical:".cls
'  Author:          svsmylav/KKK
'   Creation Date:  Dienstag, Jul 5 2005
'  Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The following Part data basis cases are addressed for the parameters specified:
'   Case A (Part data Basis value -105): Fixed Blade Weather Louver, Type A
'   Case B (Part data Basis value -110): Fixed Blade Weather Louver, Type B
'
'  Change History:
'  dd.mmm.yyyy    who    change description
'  -----------    ---    ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  08.AUG.2007     MA   CR-124221 Implemented part data basis for values 105 and 110.
'  29.AUG.2008     VKC  CR-149323 Implemented part data basis for the value 150.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
    Private Const MODULE = "CSimplePhysical" 'Used for error messages
    Private m_GeomFactory As IngrGeom3D.GeometryFactory
    Private Const E_FAIL = &H80004005
    Const NEGLIGIBLE_THICKNESS = 0.0001
    Private PI       As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
      PI = 4 * Atn(1)
      
    Set m_GeomFactory = New IngrGeom3D.GeometryFactory
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput             As Double
    Dim Width               As Double
    Dim Depth               As Double
    Dim FrameWidth          As Double
    Dim FrameDepth          As Double
    Dim LouvreLength        As Double
    Dim LouvreDist          As Double
    Dim Height1             As Double
    Dim FlangeThickness     As Double
    Dim Offset1             As Double
    Dim Offset2             As Double
    Dim parAngle            As Double
    
    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
   
    Dim dblToInner          As Double
    Dim dblToDown           As Double
    Dim dblNozzleLength     As Double
    
    Dim Dir1 As New AutoMath.DVector
    Dim Dir2 As New AutoMath.DVector
    Dim Dir3 As New AutoMath.DVector
       
    Dir1.Set 0, 0, 1
    Dir2.Set 0, 1, 0
    Dir3.Set 1, 0, 0
        
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
       
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2)
    Depth = arrayOfInputs(3)
    LouvreLength = arrayOfInputs(4)
    LouvreDist = arrayOfInputs(5)
    Height1 = arrayOfInputs(6)  'Corresponds to LouvreWidth in the WRLouvres.cls
    FlangeThickness = arrayOfInputs(7)
    'Offset1 = arrayOfInputs(8)
    'Offset2 = arrayOfInputs(9)
    'FrameWidth = arrayOfInputs(10)
    'FrameDepth = arrayOfInputs(11)
    'parAngle =arrayOfInputs(12)
      
    dblToInner = 0.005
    dblToDown = 0.01
    dblNozzleLength = 0.001
    
    iOutput = 0
    
    'Insert your code for output 1(Nozzle)

    'Place Nozzle 1
    Dim FlangeWidth As Double
    Dim NozzleFlangeThickness As Double
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim CptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim Dir As AutoMath.DVector
    Set Dir = New AutoMath.DVector
    Dim RadialDir As New AutoMath.DVector
    Dim CornerRadius As Double
    
    ' This symbol is based on the following Part data basis values that govern its geometry
    '   (Part data Basis value -105): Fixed Blade Weather Louver, Type A
    '   (Part data Basis value -110): Fixed Blade Weather Louver, Type B
    
    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    Select Case lPartDataBasis

    Case Is <= 1, 105
    'FlangeThickness = arrayOfInputs(7)
    
    '   Set HVAC nozzle parameters
        PortDepth = 0#
    
    '   FlangeWidth and Flange thickness assumed to be negigible thickness, Nozzle length to be
    '   greater than flange thickness in general.
        FlangeWidth = NEGLIGIBLE_THICKNESS
        NozzleFlangeThickness = NEGLIGIBLE_THICKNESS
        
        FlowDir = DistribFlow_UNDEFINED
        DimBaseOuter = True
        PortStatus = DistribPortStatus_BASE
        EndPrep = 301
        CptOffset = 0
        CornerRadius = 0
        
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(1, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                                NozzleFlangeThickness, FlangeWidth, FlowDir, Width, _
                                                Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                                "HNoz1", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)
        pos.Set -dblNozzleLength - LouvreLength, 0, 0
        Set iDistribPort = oHvacNozzle
        iDistribPort.SetPortLocation pos
    
        'Direction specified here of the nozzle should be the direction in which pipe will be routed.
        'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    
        Dir.Set -1, 0, 0
        iDistribPort.SetDirectionVector Dir
    
        Dir.Set 0, 1, 0
        iDistribPort.SetRadialOrient Dir
    
        Set iNozzle = oHvacNozzle
        iNozzle.Length = dblNozzleLength
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
        
        Dim objP1 As AutoMath.DPosition
        Dim objP2 As AutoMath.DPosition
        Dim objP3 As AutoMath.DPosition
        Dim objP4 As AutoMath.DPosition
        Dim stPoint   As New AutoMath.DPosition
       
        Dim lngIndex As Long
        Dim DirDown As New AutoMath.DVector
     
        Dim dblA As Double
        Dim dblHeight As Double
        
        lngIndex = 0
        
        stPoint.Set 0, 0, 0
      
        DirDown.Set -LouvreLength, 0, 0
    ' Add Code for the "Kasten": Depth LouvreLength, breite/Hoehe= Depth/Width
        Set objP1 = vecDir2(stPoint, Dir1, 0.5 * Width, Dir2, 0.5 * Depth)
        Set objP2 = vecDir2(stPoint, Dir1, -0.5 * Width, Dir2, 0.5 * Depth)
        Set objP3 = vecDir2(stPoint, Dir1, -0.5 * Width, Dir2, -0.5 * Depth)
        Set objP4 = vecDir2(stPoint, Dir1, 0.5 * Width, Dir2, -0.5 * Depth)
        Call createProjectedLines(m_OutputColl, "Feature", lngIndex, _
                objP1, objP2, objP3, objP4, DirDown, False)
        
        
        ' Add code for the Rahmen and Lamellen
        ' Rahmen:
        
        DirDown.Set -FlangeThickness, 0, 0
        
        Set objP1 = vecDir2(stPoint, Dir1, 0.5 * Width + Height1, Dir2, 0.5 * Depth + Height1)
        Set objP2 = vecDir2(stPoint, Dir1, -0.5 * Width - Height1, Dir2, 0.5 * Depth + Height1)
        Set objP3 = vecDir2(stPoint, Dir1, -0.5 * Width - Height1, Dir2, -0.5 * Depth - Height1)
        Set objP4 = vecDir2(stPoint, Dir1, 0.5 * Width + Height1, Dir2, -0.5 * Depth - Height1)
        Call createRect(m_OutputColl, "Feature", lngIndex, objP1, objP2, objP3, objP4, Height1, DirDown)
        
        
        ' Now create Lamellen
        DirDown.Set 0, 0, -Width
         
        dblA = Sqr(0.5 * FlangeThickness * FlangeThickness)
        dblHeight = 0
        
        While dblHeight < Depth - LouvreDist
        
            Set objP1 = vecDir2(stPoint, Dir1, 0.5 * Width, Dir2, 0.5 * Depth - dblHeight, Dir3, -LouvreLength + dblA)
            Set objP2 = vecDir2(stPoint, Dir1, 0.5 * Width, Dir2, 0.5 * Depth - dblA - dblHeight, Dir3, -LouvreLength)
            Set objP3 = vecDir2(stPoint, Dir1, 0.5 * Width, Dir2, 0.5 * Depth - LouvreDist - dblHeight, Dir3, -dblA)
            Set objP4 = vecDir2(stPoint, Dir1, 0.5 * Width, Dir2, 0.5 * Depth - LouvreDist + dblA - dblHeight, Dir3, 0#)
            Call createProjectedLines(m_OutputColl, "Feature", lngIndex, _
                objP1, objP2, objP3, objP4, DirDown)
                
            dblHeight = dblHeight + LouvreDist * 1.2
        
        Wend
        
        Debug.Print "Maximum index: " & lngIndex
        Set iDistribPort = Nothing
    
    Case 110

        Offset1 = arrayOfInputs(8)
        Offset2 = arrayOfInputs(9)
        LouvreLength = Offset1 + Offset2

    '   Set HVAC nozzle parameters
        PortDepth = 0#

    '   FlangeWidth and Flange thickness assumed to be negigible thickness, Nozzle length to be
    '   greater than flange thickness in general.
        
        FlangeWidth = FlangeThickness + 0.03
        NozzleFlangeThickness = FlangeThickness
    
        FlowDir = DistribFlow_UNDEFINED
        DimBaseOuter = True
        PortStatus = DistribPortStatus_BASE
        EndPrep = 11
        CptOffset = 0
        CornerRadius = 0
        dblNozzleLength = NozzleFlangeThickness + 0.001
    
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(1, "SymbDefn", GSCADNozzleEntities.Rectangular, EndPrep, _
                                                NozzleFlangeThickness, FlangeWidth, FlowDir, Width, _
                                                Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                                "HNoz1", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)
   
        pos.Set 0, 0, 0
        Set iDistribPort = oHvacNozzle
        iDistribPort.SetPortLocation pos
    
        'Direction specified here of the nozzle should be the direction in which pipe will be routed.
        'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    
        Dir.Set -1, 0, 0
        iDistribPort.SetDirectionVector Dir
    
        Dir.Set 0, 1, 0
        iDistribPort.SetRadialOrient Dir
    
        Set iNozzle = oHvacNozzle
        iNozzle.Length = dblNozzleLength
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
        
        'Create Rectangular Projection
        
        Dim oRectangle As Object
        Dim oAxisVec As AutoMath.DVector
        Set oAxisVec = New AutoMath.DVector

        Dim dLineStrPoints(0 To 14) As Double
        Dim oLineString As IngrGeom3D.LineString3d
        dLineStrPoints(0) = 0
        dLineStrPoints(1) = 0.5 * Depth
        dLineStrPoints(2) = 0.5 * Width

        dLineStrPoints(3) = 0
        dLineStrPoints(4) = 0.5 * Depth
        dLineStrPoints(5) = -0.5 * Width

        dLineStrPoints(6) = 0
        dLineStrPoints(7) = -0.5 * Depth
        dLineStrPoints(8) = -0.5 * Width

        dLineStrPoints(9) = 0
        dLineStrPoints(10) = -0.5 * Depth
        dLineStrPoints(11) = 0.5 * Width

        dLineStrPoints(12) = 0
        dLineStrPoints(13) = 0.5 * Depth
        dLineStrPoints(14) = 0.5 * Width

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set 1, 0, 0

        Set oRectangle = PlaceProjection(m_OutputColl, oLineString, oAxisVec, LouvreLength, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRectangle
        
        Set oRectangle = Nothing
        Set oLineString = Nothing
        
        'Create Vertical Flat Bar
        Dim iNumOfVanes As Long
        iNumOfVanes = (Depth * 1000) / 100
        
        Dim iSpaceBtwPlates As Double
        iSpaceBtwPlates = Depth * 1# / iNumOfVanes
        
        If CmpDblGreaterthanOrEqualTo(Width, 0.6) And CmpDblLessThan(Width, 1.2 - LINEAR_TOLERANCE) Then
                
            Dim oVertBar As Object
        
            oStPoint.Set 0.5 * Offset2, 0.5 * Depth, 0.003
            oEnPoint.Set 0.5 * Offset2 + iSpaceBtwPlates, -0.5 * Depth, -0.003
            
            Set oVertBar = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oVertBar
            Set oVertBar = Nothing
        
        End If
        
        If CmpDblGreaterthanOrEqualTo(Width, 1.2) Then
        
             Dim oVertBar1 As Object
             
             oStPoint.Set 0.5 * Offset2, 0.5 * Depth, 0.25 * Width + 0.006
             oEnPoint.Set 0.5 * Offset2 + iSpaceBtwPlates, -0.5 * Depth, 0.25 * Width
             
             Set oVertBar1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
             ' Set the output
             iOutput = iOutput + 1
             m_OutputColl.AddOutput arrayOfOutputs(iOutput), oVertBar1
             Set oVertBar1 = Nothing
             
             Dim oVertBar2 As Object
             
             oStPoint.Set 0.5 * Offset2, 0.5 * Depth, -0.25 * Width - 0.006
             oEnPoint.Set 0.5 * Offset2 + iSpaceBtwPlates, -0.5 * Depth, -0.25 * Width
             
             Set oVertBar2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
             ' Set the output
             iOutput = iOutput + 1
             m_OutputColl.AddOutput arrayOfOutputs(iOutput), oVertBar2
             Set oVertBar2 = Nothing
            
        End If
        
        'Create Vanes
        Dim oVane As Object
        
        dLineStrPoints(0) = 0.5 * Offset2
        dLineStrPoints(1) = iSpaceBtwPlates / 2
        dLineStrPoints(2) = 0.5 * Width

        dLineStrPoints(3) = (0.5 * Offset2) + iSpaceBtwPlates
        dLineStrPoints(4) = -iSpaceBtwPlates / 2
        dLineStrPoints(5) = 0.5 * Width

        dLineStrPoints(6) = (0.5 * Offset2) + iSpaceBtwPlates + (0.003 * Cos(PI / 4))
        dLineStrPoints(7) = -(iSpaceBtwPlates / 2 - (0.003 * Sin(PI / 4)))
        dLineStrPoints(8) = 0.5 * Width

        dLineStrPoints(9) = 0.5 * Offset2 + 0.003 * Cos(PI / 4)
        dLineStrPoints(10) = (iSpaceBtwPlates / 2) + (0.003 * Sin(PI / 4))
        dLineStrPoints(11) = 0.5 * Width

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set 0, 0, -1

        If iNumOfVanes = 0 Then iNumOfVanes = 1
        
        Dim iCount As Integer
        Dim oTransVector As AutoMath.DVector
        Dim oTransMatrix As DT4x4
        Set oTransMatrix = New DT4x4
        Dim dYCoor As Double
        For iCount = 1 To iNumOfVanes
            Set oVane = PlaceProjection(m_OutputColl, oLineString, oAxisVec, Width, False)
            Set oTransVector = New AutoMath.DVector
            If iCount = 1 Then
                dYCoor = (0.5 * (Depth - iSpaceBtwPlates)) - (iCount - 1) * iSpaceBtwPlates - 0.003 * Sin(PI / 4)
            Else
                dYCoor = (0.5 * (Depth - iSpaceBtwPlates)) - (iCount - 1) * iSpaceBtwPlates
            End If
             
            oTransVector.Set 0, dYCoor, 0
            oTransMatrix.LoadIdentity
            oTransMatrix.Translate oTransVector
            oVane.Transform oTransMatrix
            
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oVane
            Set oVane = Nothing
        Next

    Case 150
     
        FrameWidth = arrayOfInputs(10)
        FrameDepth = arrayOfInputs(11)
        parAngle = arrayOfInputs(12)
        
        Dim oHvacPort            As IJDHvacPort
        Dim oHvacColl            As IJDCollection
        Dim dThickness(1 To 2)   As Double
        Dim iPortIndex           As Integer
        Dim lEndPrep(1 To 2)     As Long
        Dim dFlangeWidth(1 To 2) As Double
        Dim lFlowDir(1 To 2)     As Long
        Dim dPortDepth(1 To 2)   As Double
        Dim dCptOffSet(1 To 2)   As Double
        Dim dNozzLength(1 To 2)  As Double
        Dim lCSType              As Long
          
        'Set HVAC nozzle parameters
        PortDepth = 0#

        'Nozzle length is assumed to be greater than flange thickness in general.
        
        DimBaseOuter = True
        PortStatus = DistribPortStatus_BASE
        
        iPortIndex = 1
        dNozzLength(1) = 0.001
        
        'Set HVAC nozzle parameters
        Set oHvacColl = oPartFclt.GetNozzles()
        For iPortIndex = 1 To oHvacColl.Size
            Set oHvacPort = oHvacColl.Item(iPortIndex)
            lEndPrep(iPortIndex) = oHvacPort.EndPrep
            dThickness(iPortIndex) = oHvacPort.Thickness
            dFlangeWidth(iPortIndex) = oHvacPort.FlangeWidth
            lFlowDir(iPortIndex) = oHvacPort.FlowDirection
            dPortDepth(iPortIndex) = oHvacPort.PortDepth
            dCptOffSet(iPortIndex) = oHvacPort.CptOffset
            
            'NozzleLength Has to be GREATER than NozzleFlangeThickness
            If CmpDblLessThanOrEqualTo(dThickness(iPortIndex), LINEAR_TOLERANCE) Then
                dThickness(iPortIndex) = 0.0001
            End If
            If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
                dNozzLength(iPortIndex) = dThickness(iPortIndex) + 0.001
            End If
        Next iPortIndex
        
        Set oHvacPort = Nothing
        oHvacColl.Clear
        Set oHvacColl = Nothing
    
        lCSType = GSCADNozzleEntities.Rectangular
        iPortIndex = 1
        Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", lCSType, lEndPrep(1), _
                                                dThickness(1), dFlangeWidth(1), lFlowDir(1), Width, _
                                                Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                                "HNoz1", PortDepth, CptOffset, False, m_OutputColl.ResourceManager)

        pos.Set 0, 0, 0
        Set iDistribPort = oHvacNozzle
        iDistribPort.SetPortLocation pos

        'Direction specified here of the nozzle should be the direction in which pipe will be routed.
        'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.

        Dir.Set -1, 0, 0
        iDistribPort.SetDirectionVector Dir

        Dir.Set 0, 1, 0
        iDistribPort.SetRadialOrient Dir

        Set iNozzle = oHvacNozzle
        iNozzle.Length = dNozzLength(iPortIndex)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
      
        'Add the code for the frame by using place box command.
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        
        Dim oBox1 As Object
        oStPoint.Set LouvreLength, 0.5 * Depth, 0.5 * FrameWidth
        oEnPoint.Set 0, 0.5 * FrameDepth, -0.5 * FrameWidth
        
        Set oBox1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox1
        Set oBox1 = Nothing
            
        Dim oBox2 As Object
        oStPoint.Set LouvreLength, 0.5 * Depth, -0.5 * Width
        oEnPoint.Set 0, -0.5 * Depth, -0.5 * FrameWidth
        
        Set oBox2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox2
        Set oBox2 = Nothing
            
        Dim oBox3 As Object
        oStPoint.Set LouvreLength, -0.5 * FrameDepth, 0.5 * FrameWidth
        oEnPoint.Set 0, -0.5 * Depth, -0.5 * FrameWidth
              
        Set oBox3 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox3
        Set oBox3 = Nothing
            
        Dim oBox4 As Object
        oStPoint.Set LouvreLength, 0.5 * Depth, 0.5 * FrameWidth
        oEnPoint.Set 0, -0.5 * Depth, 0.5 * Width
                 
        Set oBox4 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox4
        Set oBox4 = Nothing
                   
        'Add code for the creation of the louvers
         DirDown.Set 0, 0, -Width
         
        'dblA is the blade thickness
        'dblHeight is the spacing between the blades
        dblA = 0.01 * Depth
        dblHeight = 0
        

        While dblHeight < Depth - LouvreDist
                            
            Set objP1 = vecDir2(stPoint, Dir1, 0.5 * Width, Dir2, 0.5 * Depth - dblHeight, Dir3, 0#)
            Set objP2 = vecDir2(stPoint, Dir1, 0.5 * Width, Dir2, 0.5 * Depth - dblHeight - dblA, Dir3, 0#)
            Set objP3 = vecDir2(stPoint, Dir1, 0.5 * Width, Dir2, 0.5 * Depth - LouvreLength * Tan(parAngle) - dblHeight, Dir3, LouvreLength)
            Set objP4 = vecDir2(stPoint, Dir1, 0.5 * Width, Dir2, 0.5 * Depth - LouvreLength * Tan(parAngle) - dblHeight + dblA, Dir3, LouvreLength)
            
            Call createProjectedLines(m_OutputColl, "Feature", lngIndex, _
                                        objP1, objP2, objP3, objP4, DirDown)

            dblHeight = dblHeight + LouvreDist - dblA
        Wend
        
    Case Else
        GoTo ErrorLabel:

    End Select
    
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing

    Exit Sub
      
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
    Exit Sub
    Resume
    
End Sub
Private Sub Class_Terminate()
    Set m_GeomFactory = Nothing
End Sub

Private Function createRect(ByVal objOutputColl As Object, _
                    strName As String, _
                    lngIndex As Long, _
                    objP1 As AutoMath.DPosition, _
                    objP2 As AutoMath.DPosition, _
                    objP3 As AutoMath.DPosition, _
                    objP4 As AutoMath.DPosition, _
                    dblToInner As Double, _
                    DirDown As AutoMath.DVector) As Long
                    
' create 4 quader resulting in a rectangle:
'               1             2
'               +-------------+
'               |             |
'               |  +-------+  |
'               |  |       |  |
'               |  |       |  |
'               |  |       |  |
'               |  +-------+  |
'               |             |
'               +-------------+
'               4             3
'
'  P1 - P4 are the 4 outer corner points
'  dblToInner is the size from the outer points to the inner points
'  The size is measured for each side.
'  DirDown is the size and direction in down position
'


' Compute the 4 inner Points
Dim objVec1 As AutoMath.DVector
Dim objVec2 As AutoMath.DVector
 

Dim objP1I As AutoMath.DPosition
Dim objP2I As AutoMath.DPosition
Dim objP3I As AutoMath.DPosition
Dim objP4I As AutoMath.DPosition

 

Set objVec1 = objP2.Subtract(objP1)
Set objVec2 = objP4.Subtract(objP1)
Set objP1I = vecDir2(objP1, objVec1, dblToInner, objVec2, dblToInner)
 
Set objVec1 = objP1.Subtract(objP2)
Set objVec2 = objP3.Subtract(objP2)
Set objP2I = vecDir2(objP2, objVec1, dblToInner, objVec2, dblToInner)

Set objVec1 = objP2.Subtract(objP3)
Set objVec2 = objP4.Subtract(objP3)
Set objP3I = vecDir2(objP3, objVec1, dblToInner, objVec2, dblToInner)

Set objVec1 = objP3.Subtract(objP4)
Set objVec2 = objP1.Subtract(objP4)
Set objP4I = vecDir2(objP4, objVec1, dblToInner, objVec2, dblToInner)

' Compute the 4 curves (which will later be projected)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP1, objP2, objP2I, objP1I, DirDown)
        
         
        
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP2, objP3, objP3I, objP2I, DirDown)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP3, objP4, objP4I, objP3I, DirDown)
Call createProjectedLines(objOutputColl, strName, lngIndex, _
        objP4, objP1, objP1I, objP4I, DirDown)

End Function
Private Function createProjectedLines(ByVal objOutputColl As Object, _
                    strName As String, _
                    lngIndex As Long, _
                    objP1 As AutoMath.DPosition, _
                    objP2 As AutoMath.DPosition, _
                    objP3 As AutoMath.DPosition, _
                    objP4 As AutoMath.DPosition, _
                    objVec As AutoMath.DVector, _
                    Optional dblClosed As Boolean = True) As Long
                    
    Dim oLine As IngrGeom3D.Line3d
    Dim iElements As IJElements
    Dim complex As IngrGeom3D.ComplexString3d
    Dim Projection As IJProjection
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP1.x, objP1.y, objP1.z, objP2.x, objP2.y, objP2.z)
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    iElements.Add oLine
    Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP2.x, objP2.y, objP2.z, objP3.x, objP3.y, objP3.z)
    complex.AddCurve oLine, True
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP3.x, objP3.y, objP3.z, objP4.x, objP4.y, objP4.z)
    complex.AddCurve oLine, True
    
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, objP4.x, objP4.y, objP4.z, objP1.x, objP1.y, objP1.z)
    complex.AddCurve oLine, True
    
    Set Projection = m_GeomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, _
                                                     complex, objVec.x, objVec.y, objVec.z, objVec.Length, dblClosed)
                                                    
    lngIndex = lngIndex + 1
    objOutputColl.AddOutput strName & Trim$(Str$(lngIndex)), Projection
     
     
End Function
                    
                    
Private Function vecDir2(Pin As AutoMath.DPosition, _
                         Dir1 As AutoMath.DVector, dblSize1 As Double, _
                         Optional Dir2 As AutoMath.DVector = Nothing, _
                            Optional dblSize2 As Double = 0#, _
                         Optional Dir3 As AutoMath.DVector = Nothing, _
                            Optional dblSize3 As Double = 0#) As AutoMath.DPosition

    Dim Dir1X As AutoMath.DVector
    Dim Dir2X As AutoMath.DVector

    Set Dir1X = Dir1.Clone
    Dir1X.Length = dblSize1
    Set vecDir2 = Pin.Offset(Dir1X)

    If Not Dir2 Is Nothing Then
        Set Dir2X = Dir2.Clone
        Dir2X.Length = dblSize2
        Set vecDir2 = vecDir2.Offset(Dir2X)
    End If
    If Not Dir3 Is Nothing Then
        Set Dir2X = Dir3.Clone
        Dir2X.Length = dblSize3
        Set vecDir2 = vecDir2.Offset(Dir2X)
    End If


End Function
