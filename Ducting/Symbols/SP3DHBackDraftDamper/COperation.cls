VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   COperation.cls
'   Author:         RRK
'   Creation Date:  Friday, Aug 17 2007
'   Description:
'   This symbol is prepared for Back draft damper that caters to Round and Rectangular shapes as per CR-124909

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   02.SEP.2008     PK      CR-149330 Updated the symbol for the PDB case 145
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Operation:" 'Used for error messages
Private geomFactory As IngrGeom3D.GeometryFactory
Private Sub Class_Initialize()
    Set geomFactory = New IngrGeom3D.GeometryFactory
End Sub
Private Sub Class_Terminate()
    Set geomFactory = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

   Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parDamperLength As Double
    Dim parLeverLength As Double
    Dim parLeverPositionX As Double
    Dim parLeverPositionY As Double
    Dim parLeverPositionZ As Double
    Dim parCounterWeightAngle As Double
    Dim parHVACShape As Double
    Dim parInsulationThickness As Double
    
    Dim CornerRadius As Double
    CornerRadius = 0

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parWidth = arrayOfInputs(2)
    'parDepth = arrayOfInputs(3)
    'parDamperLength = arrayOfInputs(4)
    'parLeverLength = arrayOfInputs(5)
    'parLeverPositionX = arrayOfInputs(6)
    'parLeverPositionY = arrayOfInputs(7)
    'parLeverPositionZ = arrayOfInputs(8)
    'parCounterWeightAngle = arrayOfInputs(9)
    'parHVACShape = arrayOfInputs(10)
    'parInsulationThickness = arrayOfInputs(11)
    
    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
    
    If lPartDataBasis <= 1 Or lPartDataBasis = 145 Then
        parWidth = arrayOfInputs(2)
        parDepth = arrayOfInputs(3)
        parDamperLength = arrayOfInputs(4)
        parLeverLength = arrayOfInputs(5)
        parLeverPositionX = arrayOfInputs(6)
        parLeverPositionY = arrayOfInputs(7)
        parLeverPositionZ = arrayOfInputs(8)
        parCounterWeightAngle = arrayOfInputs(9)
        parHVACShape = arrayOfInputs(10)
        parInsulationThickness = arrayOfInputs(11)
    
        
        If CmpDblEqual(parDepth, 0) Then
            parDepth = parWidth
        End If
        
        Dim dLeverPosfrmCentX As Double
        Dim dLeverPosfrmCentY As Double
        Dim dLeverPosfrmCentZ As Double
    
        'Length and height of the frame are assumed to be 32mm
        Dim dFrameLength As Double
        dFrameLength = 0.064
        
        dLeverPosfrmCentX = (parDamperLength / 2) - parLeverPositionX
        dLeverPosfrmCentY = (parDepth / 2) - parLeverPositionY
        If parHVACShape = 4 Then
            dLeverPosfrmCentZ = Sqr((parWidth / 2) ^ 2 - (dLeverPosfrmCentY + (dFrameLength / 2)) ^ 2) + parLeverPositionZ
        Else
            dLeverPosfrmCentZ = (parWidth / 2) + parLeverPositionZ
        End If
        
        'Creating Operation Envelope (Output 1)
        
        'Assuming Counter Weight dia to be 40mm and counter weight length to 25% of Lever Length
        Dim LineStrPoints(0 To 5) As Double
    
        Dim dLeverDia As Double
        Dim dCounterWtDia As Double
        Dim dCounterWtLength As Double
            
        dLeverDia = 0.013
        dCounterWtDia = 0.04
        dCounterWtLength = parLeverLength / 4
    
        'Creating points for the outline of operation envelope
        LineStrPoints(0) = -dLeverPosfrmCentX - (dLeverDia / 2)
        LineStrPoints(1) = dLeverPosfrmCentY - parLeverLength
        LineStrPoints(2) = dLeverPosfrmCentZ + (dCounterWtDia / 2)
        
        LineStrPoints(3) = LineStrPoints(0)
        LineStrPoints(4) = dLeverPosfrmCentY + parLeverLength
        LineStrPoints(5) = LineStrPoints(2)
    
        Dim objCurvesCollection  As Collection
        Set objCurvesCollection = New Collection
        
        'Defining lines (or curves) for the outline of operation envelope
        Dim oLine As IngrGeom3D.Line3d
        Set oLine = New IngrGeom3D.Line3d
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                LineStrPoints(0), LineStrPoints(1), LineStrPoints(2), _
                                LineStrPoints(3), LineStrPoints(4), LineStrPoints(5))
        objCurvesCollection.Add oLine
        Set oLine = Nothing
        
        Dim oArc As IngrGeom3D.Arc3d
        Set oArc = New IngrGeom3D.Arc3d
        Set oArc = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, -dLeverPosfrmCentX, dLeverPosfrmCentY, LineStrPoints(2), _
                                    0, 0, -1, LineStrPoints(3), LineStrPoints(4), LineStrPoints(5), _
                                    LineStrPoints(0), LineStrPoints(1), LineStrPoints(2))
        objCurvesCollection.Add oArc
        Set oArc = Nothing
            
        Dim oStPoint As AutoMath.DPosition
        Set oStPoint = New AutoMath.DPosition
        oStPoint.Set LineStrPoints(0), LineStrPoints(1), LineStrPoints(2)
        
        Dim objOpEnvOutline  As IngrGeom3D.ComplexString3d
        Set objOpEnvOutline = PlaceTrCString(oStPoint, objCurvesCollection)
        Set objCurvesCollection = Nothing
        
        Dim objOpEnvelope As Object
        Dim oAxisVec As AutoMath.DVector
        Set oAxisVec = New AutoMath.DVector
        
        oAxisVec.Set 0, 0, -1
        Set objOpEnvelope = PlaceProjection(m_OutputColl, objOpEnvOutline, oAxisVec, dCounterWtDia, True)
    
        'Set the output
        'iOutput = iOutput + 1
        m_OutputColl.AddOutput "OperEnvelope", objOpEnvelope
        Set objOpEnvelope = Nothing
        Set objOpEnvOutline = Nothing
        Set oStPoint = Nothing
        Set oAxisVec = Nothing
    End If
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub


