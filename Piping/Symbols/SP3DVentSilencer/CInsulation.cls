VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Tuesday, Jun 4 2002
'   Description:
'   The symbol supports the following part data basis values:
'   1.)Default
'   2.)Silencer, vent or blowoff(958)
'   3.)Silencer, gas diffuser(959)

'   For the default case,this class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Symbol Model No. is: S4AC or S4AZ, Page No. D-100 PDS Piping Component Data Reference Guide.
'   Five Insulation aspect outputs are created using PlaceCylinder function.

'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------    -----                ------------------
'   09.Jul.2003    SymbolTeam(India)    Copyright Information, Header  is added.
'   22.Aug.2005    svsmylav             TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'  08.SEP.2006     KKC                  DI-95670  Replace names with initials in all revision history sheets and symbols
'  23.Apr.2008     RRK                  CR-136581 Enhanced the symbol to support new part data basis values-
'                                                  'Silencer, vent or blowoff(958)' and 'Silencer, gas diffuser(959)'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Const NEGLIGIBLE_VALUE = 0.0001
Private PI       As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    Dim ObjInsVentSilencerBody As Object
    Dim ObjInsPort1 As Object
    Dim ObjInsBody1 As Object
    Dim ObjInsPort2 As Object
    Dim ObjInsBody2 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset1 As Double
    Dim parOffset2 As Double
    Dim parVentSilencerBodyDiameter As Double
    Dim parVentSilencerBodyLength1 As Double
    Dim parVentSilencerBodyLength2 As Double
    Dim parInsulationThickness As Double
    Dim parHeight As Double
    Dim parLength As Double
    Dim parDiameter As Double
    
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    Dim oPort1CylStPoint   As AutoMath.DPosition
    Dim oPort1CylEnPoint   As AutoMath.DPosition
    Set oPort1CylStPoint = New AutoMath.DPosition
    Set oPort1CylEnPoint = New AutoMath.DPosition
    
    Dim parInsulationDiameter   As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFace1toCenter = arrayOfInputs(2)
    'parFace2toCenter = arrayOfInputs(3)
    'parOffset1 = arrayOfInputs(4)
    'parOffset2 = arrayOfInputs(5)
    'parVentSilencerBodyDiameter = arrayOfInputs(6)
    'parVentSilencerBodyLength1 = arrayOfInputs(7)
    'parVentSilencerBodyLength2 = arrayOfInputs(8)
    parInsulationThickness = arrayOfInputs(9)
    'parHeight = arrayOfInputs(10)
    'parLength = arrayOfInputs(11)
    'parDiameter = arrayOfInputs(12)

    'Retrieving part data basis value
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Select Case lPartDataBasis
    Case Is <= 1    'Default
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        parOffset1 = arrayOfInputs(4)
        parOffset2 = arrayOfInputs(5)
        parVentSilencerBodyDiameter = arrayOfInputs(6)
        parVentSilencerBodyLength1 = arrayOfInputs(7)
        parVentSilencerBodyLength2 = arrayOfInputs(8)
        
        'Setting start and end points of port 1 insulation cylinder
        oPort1CylStPoint.Set -parFace1toCenter, 0, -parOffset1
        If CmpDblEqual(flangeThick, 0) Then
            oPort1CylEnPoint.Set -parFace1toCenter + NEGLIGIBLE_VALUE, 0, -parOffset1
        Else
            oPort1CylEnPoint.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, -parOffset1
        End If
        
    Case 958    'Silencer, vent or blowoff
        parHeight = arrayOfInputs(10)
        parDiameter = arrayOfInputs(12)
        
        'Setting start and end points of port 1 insulation cylinder
        oPort1CylStPoint.Set 0, 0, 0
        If CmpDblEqual(flangeThick, 0) Then
            oPort1CylEnPoint.Set NEGLIGIBLE_VALUE, 0, 0
        Else
            oPort1CylEnPoint.Set flangeThick + parInsulationThickness, 0, 0
        End If
            
    Case 959    'Silencer, gas diffuser
        parLength = arrayOfInputs(11)
        parDiameter = arrayOfInputs(12)
        
        'Setting start and end points of port 1 insulation cylinder
        oPort1CylStPoint.Set 0, 0, 0
        If CmpDblEqual(flangeThick, 0) Then
            oPort1CylEnPoint.Set NEGLIGIBLE_VALUE, 0, 0
        Else
            oPort1CylEnPoint.Set flangeThick + parInsulationThickness, 0, 0
        End If
                    
    Case Else
        GoTo ErrorLabel:    ' Invalid Part data basis.
    End Select

    iOutput = 0

    Select Case lPartDataBasis
    Case Is <= 1    'Default
    ' Insert your code for output 1(Insulation for Filetr  Body)
        stPoint.Set -parVentSilencerBodyLength1 - parInsulationThickness, 0, 0
        enPoint.Set parVentSilencerBodyLength2 + parInsulationThickness, 0, 0
        parInsulationDiameter = parVentSilencerBodyDiameter + 2 * parInsulationThickness
        Set ObjInsVentSilencerBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
        
    ' Insert your code for output 3(Insulation for body Port1 Side)
        stPoint.Set -parFace1toCenter + flangeThick, 0, -parOffset1
        enPoint.Set -parVentSilencerBodyLength1, 0, -parOffset1
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        Set ObjInsBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
    '   Set the output
        m_OutputColl.AddOutput "InsBody1", ObjInsBody1
        Set ObjInsBody1 = Nothing
    
    ' Insert your code for output 4(Insulation Port2 Side)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        
        stPoint.Set parFace2toCenter, 0, parOffset2
        
        If flangeThick = 0 Then
            enPoint.Set parFace2toCenter - NEGLIGIBLE_VALUE, 0, parOffset2
        Else
            enPoint.Set parFace2toCenter - flangeThick - parInsulationThickness, 0, parOffset2
        End If
        
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
        Else
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        End If
        
        Set ObjInsPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
    '   Set the output
        m_OutputColl.AddOutput "InsPort2", ObjInsPort2
        Set ObjInsPort2 = Nothing
    
    ' Insert your code for output 5(Insulation for body Port2 Side)
        stPoint.Set parFace2toCenter - flangeThick, 0, parOffset2
        enPoint.Set parVentSilencerBodyLength2, 0, parOffset2
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        Set ObjInsBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
    
    '   Set the output
        m_OutputColl.AddOutput "InsBody2", ObjInsBody2
        Set ObjInsBody2 = Nothing
        
    Case 958
        'Setting the points for the line string
        Dim dPoints() As Double    'representing points in the X-Y plane
        ReDim dPoints(0 To 14)

        Dim dCyl2Length As Double
        Dim dEllipsoidLength As Double
        Dim dCyl2Diam As Double
        
        'Setting the diameter of the cylinder 2
        dCyl2Diam = pipeDiam
        'Check to see that cylinder 2 diameter is greater than silencer diameter
        If CmpDblGreaterthan(dCyl2Diam, parDiameter) Then dCyl2Diam = 0.5 * pipeDiam
        
        'Assumptions:
        '1)The length of the cylinder is assumed to be 10% of Height excluding flange thickness
        '2)The length of the ellipsoid portion is assumed to be 8% of height excluding flange thickness
        dCyl2Length = 0.1 * (parHeight - flangeThick)
        dEllipsoidLength = 0.08 * (parHeight - flangeThick)
        
        'Creating points for the complex string
        'Point 1
        dPoints(0) = flangeThick
        dPoints(1) = dCyl2Diam / 2 + parInsulationThickness
        dPoints(2) = 0

        'Point 2
        dPoints(3) = dPoints(0) + dCyl2Length
        dPoints(4) = dPoints(1)
        dPoints(5) = 0

        'Point 3
        dPoints(6) = dPoints(3)
        dPoints(7) = parDiameter / 2 + parInsulationThickness
        dPoints(8) = 0

        'Point 4
        dPoints(9) = parHeight
        dPoints(10) = dPoints(7)
        dPoints(11) = 0
        
        'Point 5
        dPoints(12) = parHeight
        dPoints(13) = 0
        dPoints(14) = 0

        Dim objCollection As Collection
        Set objCollection = New Collection

        Dim oLine As IngrGeom3D.Line3d
        
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        'Creating line 1
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(0), dPoints(1), dPoints(2), _
                                                         dPoints(3), dPoints(4), dPoints(5))
        objCollection.Add oLine

        'Creating line 2
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(3), dPoints(4), dPoints(5), _
                                                         dPoints(6), dPoints(7), dPoints(8))
        objCollection.Add oLine
        
        'Creating line 3
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(6), dPoints(7), dPoints(8), _
                                                         dPoints(9), dPoints(10), dPoints(11))

        objCollection.Add oLine

        'Creating line 4
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(9), dPoints(10), dPoints(11), _
                                                         dPoints(12), dPoints(13), dPoints(14))

        objCollection.Add oLine
        
        stPoint.Set dPoints(0), dPoints(1), dPoints(2)

        Dim objBodyOutline As IngrGeom3D.ComplexString3d
        Set objBodyOutline = PlaceTrCString(stPoint, objCollection)

        Set ObjInsVentSilencerBody = oGeomFactory.Revolutions3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                                           objBodyOutline, 1, 0, 0, 0, 0, 0, 2 * PI, True)
        Set objCollection = Nothing
        Set oLine = Nothing
        Set objBodyOutline = Nothing
        Set oGeomFactory = Nothing
    
    Case 959
        'Setting the points for the line string
        ReDim dPoints(0 To 8)
        
        'Assumption:The length of the cylinder is assumed to be 85% of length of the silencer excluding flange thickness
        Dim dCylLength As Double
        dCylLength = 0.85 * (parLength - flangeThick)
        
        'Creating points for the complex string
        'Point 1
        dPoints(0) = flangeThick
        dPoints(1) = parDiameter / 2 + parInsulationThickness
        dPoints(2) = 0

        'Point 2
        dPoints(3) = dPoints(0) + dCylLength
        dPoints(4) = dPoints(1)
        dPoints(5) = 0

        'Point 3
        dPoints(6) = parLength
        dPoints(7) = parInsulationThickness
        dPoints(8) = 0

        Set objCollection = New Collection

        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        
        'Creating line 1
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                         dPoints(0), dPoints(1), dPoints(2), _
                                                         dPoints(3), dPoints(4), dPoints(5))
        objCollection.Add oLine
        
        'Creating circular arc
        Dim oArc As IngrGeom3D.Arc3d
        Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                               dPoints(3), 0, 0, _
                                                               0, 0, -1, _
                                                               dPoints(3), dPoints(4), dPoints(5), _
                                                               dPoints(6), dPoints(7), dPoints(8))
        objCollection.Add oArc

        stPoint.Set dPoints(0), dPoints(1), dPoints(2)

        Set objBodyOutline = PlaceTrCString(stPoint, objCollection)

        Set ObjInsVentSilencerBody = oGeomFactory.Revolutions3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                                           objBodyOutline, 1, 0, 0, 0, 0, 0, 2 * PI, True)
        Set objCollection = Nothing

        Set oLine = Nothing
        Set oArc = Nothing
        Set objBodyOutline = Nothing
        Set oGeomFactory = Nothing
    End Select

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsVentSilencerBody
    Set ObjInsVentSilencerBody = Nothing
    
' Insert your code for output 2(Insulation Port1 Side)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, oPort1CylStPoint, oPort1CylEnPoint, _
                                        parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPort1CylStPoint = Nothing
    Set oPort1CylEnPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
