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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: Veena
'   Creation Date:  Wednesday, Dec 20 2006
'
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'******************************************************************************

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI       As Double

Private Const CCW = 1  'CounterClockWise
Private Const CW = 2   'ClockWise

Private Const SPIRAL_SUPPORTS = 1       'Display Points - Yes
Private Const NO_SPIRAL_SUPPORTS = 0    'Display Points - No

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     PI = 4 * Atn(1)
     Set m_oSymGeomHelper = New SymbolServices
     Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub
Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition
    Dim ObjStorageTank As Object
 
    Dim parVesselDiameter As Double
    Dim parVesselTantoTan As Double
    Dim parStartElevation As Double
    Dim parEndElevation As Double
    Dim parStartAngle As Double
    Dim parAnglebetSupports As Double
    Dim parRampAngle As Double
    Dim parSupportVerticalDistance As Double
    Dim parSupportVerticalSpacing As Double
    Dim parSpiralDirection As Double
    Dim parDisplaySpiralPoints As Long
    Dim parXboltHole As Double
    Dim parYboltHole As Double
    
    Dim iOutput As Double
    iOutput = 0
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselTantoTan = arrayOfInputs(3)
    parStartElevation = arrayOfInputs(4)
    parEndElevation = arrayOfInputs(5)
    parStartAngle = arrayOfInputs(6)
    parAnglebetSupports = arrayOfInputs(7)
    parRampAngle = arrayOfInputs(8)
    parSupportVerticalDistance = arrayOfInputs(9)
    parSupportVerticalSpacing = arrayOfInputs(10)
    parSpiralDirection = arrayOfInputs(11)
    parDisplaySpiralPoints = arrayOfInputs(12)
    parXboltHole = arrayOfInputs(13)
    parYboltHole = arrayOfInputs(14)
    
    'Changing the start angle based on the spiral direction
    If parSpiralDirection = CCW Then
        parStartAngle = 2 * PI - parStartAngle
        parAnglebetSupports = 2 * PI - parAnglebetSupports
    End If
  
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
     
    CenterPos.Set 0, 0, 0
    
    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, parVesselTantoTan
    Set ObjStorageTank = PlaceCylinder(m_outputColl, stPoint, enPoint, parVesselDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjStorageTank
    Set ObjStorageTank = Nothing
                
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set CenterPos = Nothing
        
    Dim dblNumberofSupports  As Double
    Dim intNumberofSupports As Integer
    Dim objPoint As Object
    Dim TopPoints(0 To 2) As Double
    Dim BottomPoints(0 To 2) As Double
    Dim VesselRadius As Double
    Dim i  As Integer
    VesselRadius = 0.5 * parVesselDiameter

    'Check whether RampAngle or SupportVerticalDistance is to be taken
    'If RampAngle is given, then one need to set the SupportVerticalDistance to Zero
    'and vice versa. If both Ramp angle and Support vertical distance is provided, use
    'SupportVerticalDistance for calculation of number of supports
  
    If parRampAngle <> 0 And parSupportVerticalDistance = 0 Then
        parSupportVerticalDistance = 2 * VesselRadius * Sin(parAnglebetSupports / 2) * Tan(parRampAngle)
    ElseIf parSupportVerticalDistance <> 0 And parRampAngle = 0 Then
        parSupportVerticalDistance = parSupportVerticalDistance
    ElseIf parRampAngle = 0 And parSupportVerticalDistance = 0 Then
        MsgBox "Invalid inputs: Either Ramp Angle or Support Vertical Distance should be provided"
        Exit Sub
    ElseIf parSupportVerticalDistance <> 0 And parRampAngle <> 0 Then
        parSupportVerticalDistance = parSupportVerticalDistance
    End If
    
    dblNumberofSupports = (parEndElevation - parStartElevation - parSupportVerticalSpacing) / parSupportVerticalDistance
    intNumberofSupports = Int(dblNumberofSupports)
    
    'Rounding of the number of supports
    If (dblNumberofSupports - intNumberofSupports) > 0.0001 Then
        intNumberofSupports = dblNumberofSupports + 1
    Else
        intNumberofSupports = dblNumberofSupports
    End If
       
    Dim InitialVertDist As Double
    InitialVertDist = parStartElevation
    'Set Top points
    If parDisplaySpiralPoints = SPIRAL_SUPPORTS Then
        For i = 1 To intNumberofSupports
           If i = 1 Then
               TopPoints(0) = VesselRadius * Sin(parStartAngle + (i - 1) * parAnglebetSupports)
               TopPoints(1) = VesselRadius * Cos(parStartAngle + (i - 1) * parAnglebetSupports)
               TopPoints(2) = InitialVertDist
           Else
               TopPoints(0) = VesselRadius * Sin(parStartAngle + (i - 1) * parAnglebetSupports)
               TopPoints(1) = VesselRadius * Cos(parStartAngle + (i - 1) * parAnglebetSupports)
               TopPoints(2) = InitialVertDist + (i - 1) * (parSupportVerticalDistance)
           End If
           Set objPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, TopPoints(0), TopPoints(1), TopPoints(2))
           m_outputColl.AddOutput "Point1_", objPoint
           Set objPoint = Nothing
       Next
    End If
    'Set Bottom points
    If parDisplaySpiralPoints = SPIRAL_SUPPORTS Then
             For i = 1 To intNumberofSupports
                If i = 1 Then
                    BottomPoints(0) = VesselRadius * Sin(parStartAngle + (i - 1) * parAnglebetSupports)
                    BottomPoints(1) = VesselRadius * Cos(parStartAngle + (i - 1) * parAnglebetSupports)
                    BottomPoints(2) = InitialVertDist - parSupportVerticalSpacing
                Else
                    BottomPoints(0) = VesselRadius * Sin(parStartAngle + (i - 1) * parAnglebetSupports)
                    BottomPoints(1) = VesselRadius * Cos(parStartAngle + (i - 1) * parAnglebetSupports)
                    BottomPoints(2) = InitialVertDist + (i - 1) * (parSupportVerticalDistance) - parSupportVerticalSpacing
                End If
                If BottomPoints(2) >= 0 Then
                    Set objPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, BottomPoints(0), BottomPoints(1), BottomPoints(2))
                    m_outputColl.AddOutput "Point2_", objPoint
                    Set objPoint = Nothing
               End If
            Next
      End If
      
    'Insert your code for output 4
    '==========================================
    'Construction of Equipment Foundation Port
    '==========================================
    'Coordinate sysytem is at the bottom of support.
    '
    '                      |-------|
    '   Y                  |       |
    '   ^                  |       |
    '   |                  |       |
    '   |                  |       |
    '   |                  |       |
    '   -----> X           |       |
    '   Symbol CS          |       |
    '                      |       |
    '                      |   X   |
    '                      |   ^   |
    '                      |   |   |
    '                      |   |   |
    '                      |---|---|----> Y
    '                   Port CS
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
   'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = 0

    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#
    
    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_outputColl.ResourceManager)
    Dim holes() As Variant
    Call ObjFoundationPort.GetHoles(holes())
    
    parXboltHole = VesselRadius / 2
    parYboltHole = VesselRadius

    holes(0, 1) = -(VesselRadius - parXboltHole)
    holes(0, 2) = VesselRadius - parYboltHole / 2
    holes(1, 1) = VesselRadius - parXboltHole
    holes(1, 2) = VesselRadius - parYboltHole / 2
    holes(2, 1) = VesselRadius - parXboltHole
    holes(2, 2) = -(VesselRadius - parYboltHole / 2)
    holes(3, 1) = -(VesselRadius - parXboltHole)
    holes(3, 2) = -(VesselRadius - parYboltHole / 2)
                                                   
    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
        
    Call ObjFoundationPort.SetHoles(holes)
    
    'Set the output
    m_outputColl.AddOutput "EqpFoundationPort", ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub




