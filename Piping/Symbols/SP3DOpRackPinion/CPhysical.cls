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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         dkl
'   Creation Date:  Wednesday, Jul 12 2006
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      Valve Operator Rack And Pinion (Source: VAT, Gate valve control system-Series 64,  at www.vatvalve.com)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI       As Double
Private Const MODULE = "Physical:" 'Used for error messages
Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize"
    On Error GoTo Errx

    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt         As PartFacelets.IJDPart
    Dim flangeThick       As Double
    Dim sptOffset         As Double
    Dim flangeDiam        As Double
    Dim pipeDiam          As Double
    Dim Depth             As Double
   
    Dim iOutput    As Double
            
    Dim parInstrumentHeight As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentHeight3 As Double
    Dim parWithdrawalDistance As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight = arrayOfInputs(2)  'N
    parInstrumentWidth = arrayOfInputs(3)   'O
    parInstrumentHeight2 = arrayOfInputs(4) 'U
    parInstrumentLength = arrayOfInputs(5)  'P
    parInstrumentLength1 = arrayOfInputs(6) 'V
    parInstrumentHeight3 = arrayOfInputs(7) 'R
    parWithdrawalDistance = arrayOfInputs(8)

    iOutput = 0
    
' output 1 (Rectangular Cap at the Top)
    Dim ObjRectCapAtTop  As Object
    Const CapThickness = 0.00635  'Assumption : CapThickness = 0.25"
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    stPoint.Set -parInstrumentLength / 2, parInstrumentHeight - CapThickness, parInstrumentWidth / 2
    enPoint.Set parInstrumentLength / 2, parInstrumentHeight, -parInstrumentWidth / 2
    
    Set ObjRectCapAtTop = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output 1
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRectCapAtTop
    Set ObjRectCapAtTop = Nothing

' Output 2 (The Rectangular base plate)
'     RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, Depth
    Dim ObjRectPlate   As Object
    Dim dLength   As Double
    Dim dWidth   As Double
       
    dLength = parInstrumentLength / 3 ' Assumption : Rectangular Plate dLength = parInstrumentLength / 3
    dWidth = parInstrumentWidth / 1.2 ' Assumption : Rectangular Plate dWidth = parInstrumentWidth / 1.2
    stPoint.Set -dLength / 2, 0, dWidth / 2
    enPoint.Set dLength / 2, parInstrumentHeight - CapThickness, -dWidth / 2
        
    Set ObjRectPlate = PlaceBox(m_OutputColl, stPoint, enPoint)
        
' Set the output 2
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRectPlate
    Set ObjRectPlate = Nothing
 
' output 3 (cylinder)
    Dim ObjCyl As Object
' Assumption: Instrument Box is at a distance dSpacingLength + Factor(=0.3" assumed) from the Rectangular plate
    Dim dSpacingLength As Double
    Const Factor = 0.00762
    dSpacingLength = parInstrumentLength / 2 + Factor
    stPoint.Set -dLength / 2, parInstrumentHeight3, 0
    enPoint.Set -dSpacingLength - dLength / 2, parInstrumentHeight3, 0
' Assumption: The length of this cylinder is parInstrumentWidth / 2
    Set ObjCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentHeight2 / 3, True)

' Set the output 3
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl
    Set ObjCyl = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

' output 4 (The Instrument Box)
    Dim ObjInstBox As Object
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim axisVect   As AutoMath.DVector
    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    
    Dim dDistUptoBottom As Double
    Dim dBaseLength As Double
    
'   Assumptions:
'   1. Instrument Box height, below the parInstrumentHeight3, is taken to be 1/3 of parInstrumentHeight2
'   2. Instrument Box length at the bottom is chamfered to half of the Instrument Box length at the top
    
    dDistUptoBottom = parInstrumentHeight2 / 3
    dBaseLength = (parInstrumentLength1 - dLength / 2) / 2

'1
    LineStrPoints(0) = -dLength / 2 - dSpacingLength
    LineStrPoints(1) = parInstrumentHeight3 - dDistUptoBottom
    LineStrPoints(2) = dWidth / 2
'2
    LineStrPoints(3) = -dLength / 2 - dBaseLength - dSpacingLength
    LineStrPoints(4) = LineStrPoints(1)
    LineStrPoints(5) = LineStrPoints(2)
'3
    LineStrPoints(6) = -parInstrumentLength1 - dSpacingLength - dLength / 2
    LineStrPoints(7) = parInstrumentHeight3
    LineStrPoints(8) = LineStrPoints(2)
'4
    LineStrPoints(9) = LineStrPoints(6)
    LineStrPoints(10) = parInstrumentHeight3 + parInstrumentHeight2 - dDistUptoBottom
    LineStrPoints(11) = LineStrPoints(2)
'5
    LineStrPoints(12) = LineStrPoints(0)
    LineStrPoints(13) = LineStrPoints(10)
    LineStrPoints(14) = LineStrPoints(2)
'6
    LineStrPoints(15) = LineStrPoints(0)
    LineStrPoints(16) = LineStrPoints(1)
    LineStrPoints(17) = LineStrPoints(2)

    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 6, LineStrPoints)
    
    Set axisVect = New AutoMath.DVector
    axisVect.Set 0, 0, -1
    
    Set ObjInstBox = PlaceProjection(m_OutputColl, oLineString, axisVect, dWidth, True)
        
' Set the output 4
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInstBox
    Set ObjInstBox = Nothing
    Set axisVect = Nothing
    Set geomFactory = Nothing
    Set oLineString = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
