VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MA
'   Creation Date:  Tuesday, Nov 27 2007
'   Description:
'   This symbol is created for CR-129440 (Rectangular Air Grille)
'   Source: Return Air Grille-Fresh Air Grille-Exhaust Air Grille.pdf.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   01.Sep.2008     GL      Added two new part data basis 148 and 149 for Grille-Horizontal Blades and Grille- Vertical Blades.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Integer
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parGrillLength As Double
    Dim parFrameWidth As Double
    Dim parFilterThickness As Double
    Dim parDamperLength As Double
    Dim parFrameThickness As Double
    Dim parOverallLength As Double
    Dim parBladeSpacing As Double
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oRectangle As Object
    Dim oAxisVec As AutoMath.DVector
    Dim dLineStrPoints(0 To 14) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim dSpaceBtwVanes As Double
    Dim dVaneThickness As Double
    Dim lNumOfVanes As Long
    Dim oHorVane As Object
    Dim iCount As Integer
    Dim oTransVector As AutoMath.DVector
    Dim oTransMatrix As DT4x4
    Dim oRotVector As AutoMath.DVector
    Dim dYCoord As Double
    Dim oVertVane As Object
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oBox As Object
    Dim dGrillHeight As Double
    Dim dGrillWidth As Double
    Dim dFrameThickness As Double
    Dim dFilterThickness As Double
    
    'Assume Frame Thickness to be 0.005meter
    dFrameThickness = 0.005
 
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parWidth = arrayOfInputs(2)
'    parDepth = arrayOfInputs(3)
'    parGrillLength = arrayOfInputs(4)
'    parFrameWidth = arrayOfInputs(5)
'    parFilterThickness = arrayOfInputs(6)
'    parDamperLength = arrayOfInputs(7)
    
    'Origin is taken at the Port location
    'For Horizontally routed Duct X - Along the flow and Y - towards UP
    iOutput = 0
    
    ' This symbol is based on the following Part data basis values that govern its geometry
    '   (Part data Basis value -112): Rectangular Air Grille Without Damper
    '   (Part data Basis value -113): Rectangular Air Grille With Damper
    
    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing

    Select Case lPartDataBasis

    Case Is <= 1, 112
    
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parGrillLength = arrayOfInputs(4)
        parFrameWidth = arrayOfInputs(5)
        parFilterThickness = arrayOfInputs(6)
        
        dGrillHeight = parDepth + 2 * parFrameWidth
        dGrillWidth = parWidth + 2 * parFrameWidth
        
        'Create Rectangular Projection
        dLineStrPoints(0) = 0
        dLineStrPoints(1) = 0.5 * parDepth
        dLineStrPoints(2) = 0.5 * parWidth

        dLineStrPoints(3) = 0
        dLineStrPoints(4) = 0.5 * parDepth
        dLineStrPoints(5) = -0.5 * parWidth

        dLineStrPoints(6) = 0
        dLineStrPoints(7) = -0.5 * parDepth
        dLineStrPoints(8) = -0.5 * parWidth

        dLineStrPoints(9) = 0
        dLineStrPoints(10) = -0.5 * parDepth
        dLineStrPoints(11) = 0.5 * parWidth

        dLineStrPoints(12) = 0
        dLineStrPoints(13) = 0.5 * parDepth
        dLineStrPoints(14) = 0.5 * parWidth

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 1, 0, 0

        Set oRectangle = PlaceProjection(m_OutputColl, oLineString, oAxisVec, _
                                                                parGrillLength, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRectangle
        Set oRectangle = Nothing
        Set oLineString = Nothing
        
        'Create Horizontal Vanes
        'Assume Space between Vanes is equal to Filter Thickness.
        dSpaceBtwVanes = parFilterThickness
        'Assume Vane thickness to be 0.002meter
        dVaneThickness = 0.002
        lNumOfVanes = (parDepth - dSpaceBtwVanes) / dSpaceBtwVanes

        dLineStrPoints(0) = parGrillLength - dFrameThickness - parFilterThickness
        dLineStrPoints(1) = dVaneThickness / 2
        dLineStrPoints(2) = 0.5 * parWidth

        dLineStrPoints(3) = parGrillLength - dFrameThickness - parFilterThickness
        dLineStrPoints(4) = -dVaneThickness / 2
        dLineStrPoints(5) = 0.5 * parWidth

        dLineStrPoints(6) = parGrillLength - dFrameThickness
        dLineStrPoints(7) = -dVaneThickness / 2
        dLineStrPoints(8) = 0.5 * parWidth

        dLineStrPoints(9) = parGrillLength - dFrameThickness
        dLineStrPoints(10) = dVaneThickness / 2
        dLineStrPoints(11) = 0.5 * parWidth

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set 0, 0, -1

        If lNumOfVanes = 0 Then lNumOfVanes = 1

        Set oTransMatrix = New DT4x4
        For iCount = 1 To lNumOfVanes
            Set oHorVane = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parWidth, True)
            Set oTransVector = New AutoMath.DVector
            dYCoord = (0.5 * parDepth) - (iCount) * dSpaceBtwVanes
            oTransVector.Set 0, dYCoord, 0
            oTransMatrix.LoadIdentity
            oTransMatrix.Translate oTransVector
            oHorVane.Transform oTransMatrix
            
            'Set the output
            m_OutputColl.AddOutput "HorVane", oHorVane
            Set oHorVane = Nothing
            Set oTransVector = Nothing
        Next
        Set oTransMatrix = Nothing
      
        'Create Vertical Vanes
        lNumOfVanes = (parWidth - dSpaceBtwVanes) / dSpaceBtwVanes

        dLineStrPoints(0) = parGrillLength - dFrameThickness - parFilterThickness
        dLineStrPoints(1) = parDepth / 2
        dLineStrPoints(2) = dVaneThickness / 2

        dLineStrPoints(3) = parGrillLength - dFrameThickness - parFilterThickness
        dLineStrPoints(4) = parDepth / 2
        dLineStrPoints(5) = -dVaneThickness / 2

        dLineStrPoints(6) = parGrillLength - dFrameThickness
        dLineStrPoints(7) = parDepth / 2
        dLineStrPoints(8) = -dVaneThickness / 2

        dLineStrPoints(9) = parGrillLength - dFrameThickness
        dLineStrPoints(10) = parDepth / 2
        dLineStrPoints(11) = dVaneThickness / 2

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set 0, -1, 0

        If lNumOfVanes = 0 Then lNumOfVanes = 1

        Set oTransMatrix = New DT4x4
        For iCount = 1 To lNumOfVanes
            Set oVertVane = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parDepth, True)
            Set oTransVector = New AutoMath.DVector
            dYCoord = (0.5 * parWidth) - (iCount) * dSpaceBtwVanes
            oTransVector.Set 0, 0, dYCoord
            oTransMatrix.LoadIdentity
            oTransMatrix.Translate oTransVector
            oVertVane.Transform oTransMatrix
            
            'Set the output
            m_OutputColl.AddOutput "VertVane", oVertVane
            Set oVertVane = Nothing
            Set oTransVector = Nothing
        Next
        Set oTransMatrix = Nothing
        Set oAxisVec = Nothing
   
        'Create Box 1
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        oStPoint.Set parGrillLength, dGrillHeight / 2, dGrillWidth / 2
        oEnPoint.Set parGrillLength - dFrameThickness, parDepth / 2, -dGrillWidth / 2
       
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
     
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing

        'Create Box 2
        oStPoint.Set parGrillLength, -dGrillHeight / 2, dGrillWidth / 2
        oEnPoint.Set parGrillLength - dFrameThickness, -parDepth / 2, -dGrillWidth / 2

        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
        
        'Create Box 3
        oStPoint.Set parGrillLength, parDepth / 2, parWidth / 2
        oEnPoint.Set parGrillLength - dFrameThickness, -parDepth / 2, dGrillWidth / 2

        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
    
        'Create Box 4
        oStPoint.Set parGrillLength, parDepth / 2, -parWidth / 2
        oEnPoint.Set parGrillLength - dFrameThickness, -parDepth / 2, -dGrillWidth / 2

        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
    
    Case 113
    
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parGrillLength = arrayOfInputs(4)
        parFrameWidth = arrayOfInputs(5)
        parFilterThickness = arrayOfInputs(6)
        parDamperLength = arrayOfInputs(7)
        
        dGrillHeight = parDepth + 2 * parFrameWidth
        dGrillWidth = parWidth + 2 * parFrameWidth
        
        'Create Rectangular Projection 1
        dLineStrPoints(0) = 0
        dLineStrPoints(1) = 0.5 * parDepth
        dLineStrPoints(2) = 0.5 * parWidth

        dLineStrPoints(3) = 0
        dLineStrPoints(4) = 0.5 * parDepth
        dLineStrPoints(5) = -0.5 * parWidth

        dLineStrPoints(6) = 0
        dLineStrPoints(7) = -0.5 * parDepth
        dLineStrPoints(8) = -0.5 * parWidth

        dLineStrPoints(9) = 0
        dLineStrPoints(10) = -0.5 * parDepth
        dLineStrPoints(11) = 0.5 * parWidth

        dLineStrPoints(12) = 0
        dLineStrPoints(13) = 0.5 * parDepth
        dLineStrPoints(14) = 0.5 * parWidth

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 1, 0, 0

        Set oRectangle = PlaceProjection(m_OutputColl, oLineString, oAxisVec, _
                                                 parDamperLength, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRectangle
        Set oRectangle = Nothing
        Set oLineString = Nothing
        Set oAxisVec = Nothing
        
        'Create Rectangular Projection 2
        dLineStrPoints(0) = parDamperLength
        dLineStrPoints(1) = 0.5 * parDepth
        dLineStrPoints(2) = 0.5 * parWidth

        dLineStrPoints(3) = parDamperLength
        dLineStrPoints(4) = 0.5 * parDepth
        dLineStrPoints(5) = -0.5 * parWidth

        dLineStrPoints(6) = parDamperLength
        dLineStrPoints(7) = -0.5 * parDepth
        dLineStrPoints(8) = -0.5 * parWidth

        dLineStrPoints(9) = parDamperLength
        dLineStrPoints(10) = -0.5 * parDepth
        dLineStrPoints(11) = 0.5 * parWidth

        dLineStrPoints(12) = parDamperLength
        dLineStrPoints(13) = 0.5 * parDepth
        dLineStrPoints(14) = 0.5 * parWidth

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 1, 0, 0

        Set oRectangle = PlaceProjection(m_OutputColl, oLineString, oAxisVec, _
                                                parGrillLength, False)

        'Set the output
        m_OutputColl.AddOutput "RectProj", oRectangle
        Set oRectangle = Nothing
        Set oLineString = Nothing
        Set oAxisVec = Nothing
        
        'Create Horizontal Vanes
        'Assume Space between Vanes is equal to Filter Thickness.
        dSpaceBtwVanes = parFilterThickness
        'Assume Vane thickness to be 0.002meter
        dVaneThickness = 0.002
        lNumOfVanes = (parDepth - dSpaceBtwVanes) / dSpaceBtwVanes

        dLineStrPoints(0) = parGrillLength + parDamperLength - dFrameThickness - parFilterThickness
        dLineStrPoints(1) = dVaneThickness / 2
        dLineStrPoints(2) = 0.5 * parWidth

        dLineStrPoints(3) = parGrillLength + parDamperLength - dFrameThickness - parFilterThickness
        dLineStrPoints(4) = -dVaneThickness / 2
        dLineStrPoints(5) = 0.5 * parWidth

        dLineStrPoints(6) = parGrillLength + parDamperLength - dFrameThickness
        dLineStrPoints(7) = -dVaneThickness / 2
        dLineStrPoints(8) = 0.5 * parWidth

        dLineStrPoints(9) = parGrillLength + parDamperLength - dFrameThickness
        dLineStrPoints(10) = dVaneThickness / 2
        dLineStrPoints(11) = 0.5 * parWidth

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 0, 0, -1

        If lNumOfVanes = 0 Then lNumOfVanes = 1

        Set oTransMatrix = New DT4x4
        For iCount = 1 To lNumOfVanes
            Set oHorVane = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parWidth, True)
            Set oTransVector = New AutoMath.DVector
            dYCoord = (0.5 * parDepth) - (iCount) * dSpaceBtwVanes
            oTransVector.Set 0, dYCoord, 0
            oTransMatrix.LoadIdentity
            oTransMatrix.Translate oTransVector
            oHorVane.Transform oTransMatrix
            
            'Set the output
            m_OutputColl.AddOutput "HorVane", oHorVane
            Set oHorVane = Nothing
            Set oTransVector = Nothing
        Next
        Set oTransMatrix = Nothing
        
        'Create Vertical Vanes
        lNumOfVanes = (parWidth - dSpaceBtwVanes) / dSpaceBtwVanes

        dLineStrPoints(0) = parGrillLength + parDamperLength - dFrameThickness - parFilterThickness
        dLineStrPoints(1) = parDepth / 2
        dLineStrPoints(2) = dVaneThickness / 2

        dLineStrPoints(3) = parGrillLength + parDamperLength - dFrameThickness - parFilterThickness
        dLineStrPoints(4) = parDepth / 2
        dLineStrPoints(5) = -dVaneThickness / 2

        dLineStrPoints(6) = parGrillLength + parDamperLength - dFrameThickness
        dLineStrPoints(7) = parDepth / 2
        dLineStrPoints(8) = -dVaneThickness / 2

        dLineStrPoints(9) = parGrillLength + parDamperLength - dFrameThickness
        dLineStrPoints(10) = parDepth / 2
        dLineStrPoints(11) = dVaneThickness / 2

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set 0, -1, 0

        If lNumOfVanes = 0 Then lNumOfVanes = 1

        Set oTransMatrix = New DT4x4
        For iCount = 1 To lNumOfVanes
            Set oVertVane = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parDepth, True)
            Set oTransVector = New AutoMath.DVector
            dYCoord = (0.5 * parWidth) - (iCount) * dSpaceBtwVanes
            oTransVector.Set 0, 0, dYCoord
            oTransMatrix.LoadIdentity
            oTransMatrix.Translate oTransVector
            oVertVane.Transform oTransMatrix
            
            'Set the output
            m_OutputColl.AddOutput "VertVane", oVertVane
            Set oVertVane = Nothing
            Set oTransVector = Nothing
        Next
        Set oTransMatrix = Nothing
        
        'Create Box 1
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        
        oStPoint.Set parGrillLength + parDamperLength, dGrillHeight / 2, dGrillWidth / 2
        oEnPoint.Set parGrillLength + parDamperLength - dFrameThickness, parDepth / 2, _
                                                                    -dGrillWidth / 2
        
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
      
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing

        'Create Box 2
        oStPoint.Set parGrillLength + parDamperLength, -dGrillHeight / 2, _
                                                                    dGrillWidth / 2
        oEnPoint.Set parGrillLength + parDamperLength - dFrameThickness, _
                                                        -parDepth / 2, -dGrillWidth / 2

        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
        
        'Create Box 3
        oStPoint.Set parGrillLength + parDamperLength, parDepth / 2, parWidth / 2
        oEnPoint.Set parGrillLength + parDamperLength - dFrameThickness, _
                                                        -parDepth / 2, dGrillWidth / 2

        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
    
        'Create Box 4
        oStPoint.Set parGrillLength + parDamperLength, parDepth / 2, -parWidth / 2
        oEnPoint.Set parGrillLength + parDamperLength - dFrameThickness, _
                                                        -parDepth / 2, -dGrillWidth / 2

        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
      
        'Create Damper Vanes
        Dim oDamperVane As Object
        'Assume Space between Vanes is equal to Filter Thickness.
        dSpaceBtwVanes = parFilterThickness
        'Assume Vane thickness to be 0.002meter
        dVaneThickness = 0.002
        lNumOfVanes = (parDepth + dSpaceBtwVanes) / _
                                (0.8 * parDamperLength - 0.002 * Cos(PI / 4) + _
                                0.002 * Sin(PI / 4) + dSpaceBtwVanes)

        dLineStrPoints(0) = 0.1 * parDamperLength + 0.002 * Cos(PI / 4)
        dLineStrPoints(1) = parDepth / 2
        dLineStrPoints(2) = 0.5 * parWidth

        dLineStrPoints(3) = 0.1 * parDamperLength
        dLineStrPoints(4) = parDepth / 2 - 0.002 * Sin(PI / 4)
        dLineStrPoints(5) = 0.5 * parWidth

        dLineStrPoints(6) = 0.1 * parDamperLength + (0.8 * parDamperLength - 0.002 * Cos(PI / 4))
        dLineStrPoints(7) = (parDepth / 2) - 0.002 * Sin(PI / 4) - (0.8 * parDamperLength - (0.002 * Cos(PI / 4)))
        dLineStrPoints(8) = 0.5 * parWidth

        dLineStrPoints(9) = 0.9 * parDamperLength
        dLineStrPoints(10) = dLineStrPoints(7) + (0.002 * Sin(PI / 4))
        dLineStrPoints(11) = 0.5 * parWidth

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set 0, 0, -1
        
        'Translate the linestring to origin
        Set oTransMatrix = New DT4x4
        Set oTransVector = New AutoMath.DVector
        Dim dConst As Double
        dConst = 0.5 * (0.8 * parDamperLength - 0.002 * (Cos(PI / 4) - Sin(PI / 4)))
        
        oTransVector.Set -0.5 * parDamperLength, -(parDepth / 2 - dConst), 0
        oTransMatrix.LoadIdentity
        oTransMatrix.Translate oTransVector
        oLineString.Transform oTransMatrix
        
        Set oTransVector = Nothing
        Set oTransMatrix = Nothing
        
        If lNumOfVanes = 0 Then lNumOfVanes = 1
        
        Set oTransMatrix = New DT4x4
        Set oRotVector = New AutoMath.DVector
        Set oTransVector = New AutoMath.DVector
        
        For iCount = 0 To lNumOfVanes - 1
            Set oDamperVane = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parWidth, True)
            dYCoord = (parDepth / 2) - ((2 * iCount + 1) * dConst + (iCount * dSpaceBtwVanes))
            If (iCount Mod 2) = 0 Then
                oTransVector.Set 0.5 * parDamperLength, dYCoord, 0
                oTransMatrix.LoadIdentity
                oTransMatrix.Translate oTransVector
                oDamperVane.Transform oTransMatrix
                
                'Set the output
                m_OutputColl.AddOutput "DamperVane", oDamperVane
                Set oDamperVane = Nothing
            Else
                oRotVector.Set 0, 0, 1
                oTransMatrix.LoadIdentity
                oTransMatrix.Rotate PI / 2, oRotVector
                oDamperVane.Transform oTransMatrix
                
                oTransVector.Set 0.5 * parDamperLength, dYCoord, 0
                oTransMatrix.LoadIdentity
                oTransMatrix.Translate oTransVector
                oDamperVane.Transform oTransMatrix
                
                'Set the output
                m_OutputColl.AddOutput "DamperVane", oDamperVane
                Set oDamperVane = Nothing
            End If
        Next
        
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        

    Case 148
    
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parGrillLength = arrayOfInputs(4)
        parFrameWidth = arrayOfInputs(5)
        parFrameThickness = arrayOfInputs(8)
        parOverallLength = arrayOfInputs(9)
        parBladeSpacing = arrayOfInputs(10)
        
        dFilterThickness = parGrillLength - parFrameThickness
        
        dGrillHeight = parDepth + 2 * parFrameWidth
        dGrillWidth = parWidth + 2 * parFrameWidth
        dSpaceBtwVanes = parBladeSpacing
        
        'Create Rectangular Projection
        dLineStrPoints(0) = parOverallLength - parGrillLength
        dLineStrPoints(1) = 0.5 * parDepth
        dLineStrPoints(2) = 0.5 * parWidth

        dLineStrPoints(3) = parOverallLength - parGrillLength
        dLineStrPoints(4) = 0.5 * parDepth
        dLineStrPoints(5) = -0.5 * parWidth

        dLineStrPoints(6) = parOverallLength - parGrillLength
        dLineStrPoints(7) = -0.5 * parDepth
        dLineStrPoints(8) = -0.5 * parWidth

        dLineStrPoints(9) = parOverallLength - parGrillLength
        dLineStrPoints(10) = -0.5 * parDepth
        dLineStrPoints(11) = 0.5 * parWidth

        dLineStrPoints(12) = parOverallLength - parGrillLength
        dLineStrPoints(13) = 0.5 * parDepth
        dLineStrPoints(14) = 0.5 * parWidth

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 1, 0, 0

        Set oRectangle = PlaceProjection(m_OutputColl, oLineString, oAxisVec, _
                                                                parGrillLength, False)
                                                        

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRectangle
        Set oRectangle = Nothing
        Set oLineString = Nothing
        
        'Create DamperBox
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        
        oStPoint.Set 0, parDepth / 2, parWidth / 2
        oEnPoint.Set parOverallLength - parGrillLength, -parDepth / 2, _
                                                                    -parWidth / 2
        
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        m_OutputColl.AddOutput "Damper", oBox
         Set oBox = Nothing
        
        'Assume Vane thickness to be 0.002meter
        'Create Horizontal Vanes
        dVaneThickness = 0.002
        lNumOfVanes = (parDepth - dSpaceBtwVanes) / dSpaceBtwVanes

        dLineStrPoints(0) = parOverallLength - dFrameThickness - dFilterThickness
        dLineStrPoints(1) = dVaneThickness / 2
        dLineStrPoints(2) = 0.5 * parWidth

        dLineStrPoints(3) = parOverallLength - dFrameThickness - dFilterThickness
        dLineStrPoints(4) = -dVaneThickness / 2
        dLineStrPoints(5) = 0.5 * parWidth

        dLineStrPoints(6) = parOverallLength - dFrameThickness
        dLineStrPoints(7) = -dVaneThickness / 2
        dLineStrPoints(8) = 0.5 * parWidth

        dLineStrPoints(9) = parOverallLength - dFrameThickness
        dLineStrPoints(10) = dVaneThickness / 2
        dLineStrPoints(11) = 0.5 * parWidth

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set 0, 0, -1

        If lNumOfVanes = 0 Then lNumOfVanes = 1

        Set oTransMatrix = New DT4x4
        For iCount = 1 To lNumOfVanes
            Set oHorVane = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parWidth, True)
            Set oTransVector = New AutoMath.DVector
            dYCoord = (0.5 * parDepth) - (iCount) * dSpaceBtwVanes
            oTransVector.Set 0, dYCoord, 0
            oTransMatrix.LoadIdentity
            oTransMatrix.Translate oTransVector
            oHorVane.Transform oTransMatrix
            
            'Set the output
            m_OutputColl.AddOutput "HorVane", oHorVane
            Set oHorVane = Nothing
            Set oTransVector = Nothing
        Next
        Set oTransMatrix = Nothing
        'Create Box 1
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        
        oStPoint.Set parOverallLength, dGrillHeight / 2, dGrillWidth / 2
        oEnPoint.Set parOverallLength - dFrameThickness, parDepth / 2, _
                                                                    -dGrillWidth / 2
        
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
      
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing

        'Create Box 2
        oStPoint.Set parOverallLength, -dGrillHeight / 2, dGrillWidth / 2
        oEnPoint.Set parOverallLength - dFrameThickness, -parDepth / 2, -dGrillWidth / 2
        
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
        
        'Create Box 3
        oStPoint.Set parOverallLength, parDepth / 2, parWidth / 2
        oEnPoint.Set parOverallLength - dFrameThickness, _
                                                        -parDepth / 2, dGrillWidth / 2

        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
    
        'Create Box 4
        oStPoint.Set parOverallLength, parDepth / 2, -parWidth / 2
        oEnPoint.Set parOverallLength - dFrameThickness, _
                                                        -parDepth / 2, -dGrillWidth / 2

        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        
    Case 149
    
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parGrillLength = arrayOfInputs(4)
        parFrameWidth = arrayOfInputs(5)
        parFrameThickness = arrayOfInputs(8)
        parOverallLength = arrayOfInputs(9)
        parBladeSpacing = arrayOfInputs(10)
        
        dFilterThickness = parGrillLength - parFrameThickness
        
        dGrillHeight = parDepth + 2 * parFrameWidth
        dGrillWidth = parWidth + 2 * parFrameWidth
        dSpaceBtwVanes = parBladeSpacing
        
        'Create Rectangular Projection
        dLineStrPoints(0) = parOverallLength - parGrillLength
        dLineStrPoints(1) = 0.5 * parDepth
        dLineStrPoints(2) = 0.5 * parWidth

        dLineStrPoints(3) = parOverallLength - parGrillLength
        dLineStrPoints(4) = 0.5 * parDepth
        dLineStrPoints(5) = -0.5 * parWidth

        dLineStrPoints(6) = parOverallLength - parGrillLength
        dLineStrPoints(7) = -0.5 * parDepth
        dLineStrPoints(8) = -0.5 * parWidth

        dLineStrPoints(9) = parOverallLength - parGrillLength
        dLineStrPoints(10) = -0.5 * parDepth
        dLineStrPoints(11) = 0.5 * parWidth

        dLineStrPoints(12) = parOverallLength - parGrillLength
        dLineStrPoints(13) = 0.5 * parDepth
        dLineStrPoints(14) = 0.5 * parWidth

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        Set oAxisVec = New AutoMath.DVector
        oAxisVec.Set 1, 0, 0

        Set oRectangle = PlaceProjection(m_OutputColl, oLineString, oAxisVec, _
                                                                parGrillLength, False)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRectangle
        Set oRectangle = Nothing
        Set oLineString = Nothing
        
        'Create DamperBox
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        
        oStPoint.Set 0, parDepth / 2, parWidth / 2
        oEnPoint.Set parOverallLength - parGrillLength, -parDepth / 2, _
                                                                    -parWidth / 2
        
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        m_OutputColl.AddOutput "Damper", oBox
        Set oBox = Nothing
         
        'Create Vertical Vanes
        'Assume Vane thickness to be 0.002meter
        dVaneThickness = 0.002
        lNumOfVanes = (parWidth - dSpaceBtwVanes) / dSpaceBtwVanes

        dLineStrPoints(0) = parOverallLength - dFrameThickness - dFilterThickness
        dLineStrPoints(1) = parDepth / 2
        dLineStrPoints(2) = dVaneThickness / 2

        dLineStrPoints(3) = parOverallLength - dFrameThickness - dFilterThickness
        dLineStrPoints(4) = parDepth / 2
        dLineStrPoints(5) = -dVaneThickness / 2

        dLineStrPoints(6) = parOverallLength - dFrameThickness
        dLineStrPoints(7) = parDepth / 2
        dLineStrPoints(8) = -dVaneThickness / 2

        dLineStrPoints(9) = parOverallLength - dFrameThickness
        dLineStrPoints(10) = parDepth / 2
        dLineStrPoints(11) = dVaneThickness / 2

        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)

        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        oAxisVec.Set 0, -1, 0

        If lNumOfVanes = 0 Then lNumOfVanes = 1

        Set oTransMatrix = New DT4x4
        For iCount = 1 To lNumOfVanes
            Set oVertVane = PlaceProjection(m_OutputColl, oLineString, oAxisVec, parDepth, True)
            Set oTransVector = New AutoMath.DVector
            dYCoord = (0.5 * parWidth) - (iCount) * dSpaceBtwVanes
            oTransVector.Set 0, 0, dYCoord
            oTransMatrix.LoadIdentity
            oTransMatrix.Translate oTransVector
            oVertVane.Transform oTransMatrix
            
            'Set the output
            m_OutputColl.AddOutput "VertVane", oVertVane
            Set oVertVane = Nothing
            Set oTransVector = Nothing
        Next
        Set oTransMatrix = Nothing
        
        'Create Box 1
        Set oStPoint = New AutoMath.DPosition
        Set oEnPoint = New AutoMath.DPosition
        
        oStPoint.Set parOverallLength, dGrillHeight / 2, dGrillWidth / 2
        oEnPoint.Set parOverallLength - dFrameThickness, parDepth / 2, _
                                                                    -dGrillWidth / 2
        
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
      
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing

        'Create Box 2
        oStPoint.Set parOverallLength, -dGrillHeight / 2, dGrillWidth / 2
        oEnPoint.Set parOverallLength - dFrameThickness, -parDepth / 2, -dGrillWidth / 2
        
        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
        
        'Create Box 3
        oStPoint.Set parOverallLength, parDepth / 2, parWidth / 2
        oEnPoint.Set parOverallLength - dFrameThickness, _
                                                        -parDepth / 2, dGrillWidth / 2

        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
    
        'Create Box 4
        oStPoint.Set parOverallLength, parDepth / 2, -parWidth / 2
        oEnPoint.Set parOverallLength - dFrameThickness, _
                                                        -parDepth / 2, -dGrillWidth / 2

        Set oBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oBox
        Set oBox = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        
        Set oTransVector = Nothing
        Set oRotVector = Nothing
        Set oTransMatrix = Nothing
        Set oAxisVec = Nothing

    Case Else
        GoTo ErrorLabel:

    End Select

    '==================================================
    'BUILD HVACNOZZLE 1
    '==================================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oDistribPort As GSCADNozzleEntities.IJDistribPort

    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim oPos As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim CornerRadius As Double

    Dim oHvacPort As IJDHvacPort
    Dim oHvacColl As IJDCollection

    Dim lEndPrep(1) As Long
    Dim dThickness(1) As Double
    Dim dFlangeWidth(1) As Double
    Dim lFlowDir(1) As Long
    Dim dPortDepth(1) As Double
    Dim dCptOffSet(1) As Double
    Dim dNozzLength(1) As Double
    
    CornerRadius = 0#
    dNozzLength(1) = 0.001
    
    iPortIndex = 1
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
            dThickness(iPortIndex) = NEGLIGIBLE_THICKNESS
        End If
        If CmpDblLessThan(dNozzLength(iPortIndex), dThickness(iPortIndex)) Then
            dNozzLength(iPortIndex) = dThickness(iPortIndex) + NEGLIGIBLE_THICKNESS
        End If
    Next iPortIndex
        
    Set oHvacPort = Nothing
    oHvacColl.Clear
    Set oHvacColl = Nothing
    
    iPortIndex = 1
    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", _
                                GSCADNozzleEntities.Rectangular, lEndPrep(1), _
                                dThickness(1), dFlangeWidth(1), lFlowDir(1), parWidth, _
                                parDepth, CornerRadius, DimBaseOuter, PortStatus, _
                                "HNoz1", dPortDepth(1), dCptOffSet(1), False, _
                                m_OutputColl.ResourceManager)
 
    'Position of the Nozzle
    Dim x As Double, y As Double, z As Double
    x = 0#
    y = 0#
    z = 0#
    oPos.Set x, y, z
    Set oDistribPort = oHvacNozzle
    oDistribPort.SetPortLocation oPos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    oDir.Set -1, 0, 0
    oDistribPort.SetDirectionVector oDir

    oDir.Set 0, 1, 0
    oDistribPort.SetRadialOrient oDir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = dNozzLength(1)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set oDistribPort = Nothing
    Set NozzleFactory = Nothing
    Set oPos = Nothing
    Set oDir = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
