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
'   Creation Date:  Thursday, Jul 6 2006
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
''     Pressure Transmitter (Source: EMCO, Model TEM at www.emcoflow.com)
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

        PI = 4 * Atn(1)
    
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
            
    Dim parInstrumentDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentNutWidth As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentHeight3 As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentDiameter = arrayOfInputs(2)
    parInstrumentHeight = arrayOfInputs(3)
    parInstrumentNutWidth = arrayOfInputs(4)
    parInstrumentHeight1 = arrayOfInputs(5)
    parInstrumentHeight2 = arrayOfInputs(6)
    parInstrumentHeight3 = arrayOfInputs(7)
       
    iOutput = 0
    


' Output 2 (The heaxagonal nut at the port i.e 'PortNut')
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, Depth
    
    Dim ObjPortNut   As Object
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim iCount            As Integer
    Dim dAngle             As Double

    Dim axisVect   As AutoMath.DVector
    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    dAngle = PI / 3
    stPoint.Set 0, 0, 0
'   First point is considered on positive Z-axis at an angle (PI / 6 + dAngle) from positive Y-axis.
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = stPoint.x
        LineStrPoints(3 * iCount - 2) = stPoint.y + (parInstrumentNutWidth / 2) * Cos(PI / 6 + iCount * dAngle)
        LineStrPoints(3 * iCount - 1) = stPoint.z + (parInstrumentNutWidth / 2) * Sin(PI / 6 + iCount * dAngle)
    Next iCount
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Set axisVect = New AutoMath.DVector
    axisVect.Set 1, 0, 0
    Set ObjPortNut = PlaceProjection(m_OutputColl, oLineString, axisVect, parInstrumentHeight2 - parInstrumentHeight1, True)

' Set the output 2
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPortNut
    Set ObjPortNut = Nothing
 
' output 3(The cylinder immediately above the Port Nut i.e Cylinder1 )
    Dim ObjCylinder1 As Object
        
    Const F = 0.0178 ' Assumption dcylinder1Dia = parInstrumentDiameter - F
    stPoint.Set parInstrumentHeight2 - parInstrumentHeight1, 0, 0
    enPoint.Set parInstrumentHeight2, 0, 0
    Set ObjCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter - F, True)

' Set the output 3
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1
    Set ObjCylinder1 = Nothing
        
' output 4 (The cylinder above Cylinder1 i.e Cylinder2 )
    Dim ObjCylinder2 As Object

    stPoint.Set parInstrumentHeight2, 0, 0
    enPoint.Set parInstrumentHeight2 + parInstrumentHeight3, 0, 0
    Set ObjCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)

' Set the output 4
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder2
    Set ObjCylinder2 = Nothing
    
'  output 5 (The Nut above Cylinder2 i.e Nut1)
    Dim ObjNut1 As Object
    Dim dHexDia As Double
    Dim dNutHeight As Double
    Const k = 0.85            'Assumption : Nut Height = K * Pipe Dia
    Const dPipeDia = 0.000635 'Assumption: Cylinder3 Dia is 0.25"
    
    dHexDia = 1.5 * pipeDiam  'Assumption:Distance between two faces of hexagon(dHexDia) = 1.5 * pipeDiam
    dNutHeight = k * pipeDiam
    
    stPoint.Set parInstrumentHeight2 + parInstrumentHeight3, 0, 0
'   First point is considered on positive Z-axis at an angle (PI / 6 + dAngle) from positive Y-axis.
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = stPoint.x
        LineStrPoints(3 * iCount - 2) = stPoint.y + (dHexDia / 2) * Cos(PI / 6 + iCount * dAngle)
        LineStrPoints(3 * iCount - 1) = stPoint.z + (dHexDia / 2) * Sin(PI / 6 + iCount * dAngle)
    Next iCount
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    axisVect.Set 1, 0, 0
    Set ObjNut1 = PlaceProjection(m_OutputColl, oLineString, axisVect, dNutHeight, True)
    
 ' Set the output 5
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNut1
    Set ObjNut1 = Nothing
    Set axisVect = Nothing
    Set geomFactory = Nothing
    Set oLineString = Nothing
     
' output 6 (The cylinder above Nut1 i.e cylinder3)
    Dim ObjCylinder3 As Object
    Const heightCylinder3 = 0.002032     'Assumption: Cylinder3 height is 0.80"
    stPoint.Set parInstrumentHeight2 + parInstrumentHeight3 + dNutHeight, 0, 0
    enPoint.Set parInstrumentHeight2 + parInstrumentHeight3 + dNutHeight + heightCylinder3, 0, 0
    Set ObjCylinder3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dPipeDia, True)

' Set the output 6
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder3
    Set ObjCylinder3 = Nothing

' output 7 (The cylinder above cylinder3 i.e cylinder4)
    Dim ObjCylinder4 As Object
    Dim heightCylinder4 As Double
    heightCylinder4 = parInstrumentHeight - parInstrumentHeight2 - parInstrumentHeight3 _
                       - dNutHeight - heightCylinder3
    stPoint.Set parInstrumentHeight - heightCylinder4, 0, 0
    enPoint.Set parInstrumentHeight, 0, 0
    Set ObjCylinder4 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dPipeDia, True)

' Set the output 7
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder4
    Set ObjCylinder4 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' output 1 (Place Nozzle 1)
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set 0, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
' Set the output 1
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
