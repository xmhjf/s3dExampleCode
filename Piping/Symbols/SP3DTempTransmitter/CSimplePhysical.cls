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
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         RRK
'   Creation Date:  Thursday, Jul 17 2006
'   Description:
'    Physical Aspect of Temperature Transmitter, Model TEM (at http://www.emcoflow.com/documents/990103%20Rev.G.pdf)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private PI As Double
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
  
    PI = Atn(1) * 4
    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
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
    
'   Declaration of temporary variable dX to store distances along
'   X direction
    Dim dX As Double
       
'Inputs
    Dim parInstrumentHeight As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    
    Dim iOutput     As Double
    iOutput = 0
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight = arrayOfInputs(2)
    parInstrumentLength = arrayOfInputs(3)
    parInstrumentLength1 = arrayOfInputs(4)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
' Place Output 1(Junction Box)
' -------------------------------------------------------------------------------------------------------------------
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    Dim dJnBoxWidth As Double
    Dim dJnBoxCenter As Double
    Dim dJnBoxDia As Double
    
    Dim dHexFacetoFace As Double
    
RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
'Assumption

    dHexFacetoFace = 1.5 * pipeDiam

'Assumption:
    dJnBoxDia = 8 * parInstrumentHeight / 9
    dJnBoxWidth = Sqr(3) * dHexFacetoFace / 2
    dJnBoxCenter = parInstrumentLength - parInstrumentLength1 - (5 * parInstrumentHeight / 9)
    
    stPoint.Set dJnBoxCenter, 0, -dJnBoxWidth / 2
    enPoint.Set dJnBoxCenter, 0, dJnBoxWidth / 2

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dJnBoxDia

   

' Place Output 2(Cylinder 1)
' -------------------------------------------------------------------------------------------------------------------
    Dim dCyl1Dia As Double

'Assumption:
    dCyl1Dia = 0.5 * pipeDiam
    
    dX = parInstrumentLength - parInstrumentLength1

    stPoint.Set dJnBoxCenter, 0, 0
    enPoint.Set dX, 0, 0


' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCyl1Dia


    
' Place Output 3(Cylinder 2)
' -------------------------------------------------------------------------------------------------------------------
    Dim dCyl2Dia As Double
    Dim dCyl2Lnth As Double
' Assumption:
    dCyl2Lnth = 5 * parInstrumentHeight / 9
    dCyl2Dia = 26.67
    
    If CmpDblLessThan(dJnBoxWidth, dCyl2Dia) Then dCyl2Dia = 0.9 * dJnBoxWidth
    stPoint.Set dJnBoxCenter, 0, 0
    enPoint.Set dJnBoxCenter, -dCyl2Lnth, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCyl2Dia
    
   
' Place Output 4(Cylinder 3)
' -------------------------------------------------------------------------------------------------------------------
    Dim dCyl3Dia As Double
    Dim dCyl3Ht As Double

' Assumption: The Diameter of the Cylinder 3 is assumed as 1.1 times of dDiam
    dCyl3Dia = 1.1 * pipeDiam
    dCyl3Ht = 5 * parInstrumentHeight / 9
    
    stPoint.Set dJnBoxCenter, 0, 0
    enPoint.Set dJnBoxCenter - dCyl3Ht, 0, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCyl3Dia
    
   
    
' Place Output 5 (Union)
' -------------------------------------------------------------------------------------------------------------------

    Dim dUnionHeight As Double
    
'Assumption
    dUnionHeight = 1.6 * pipeDiam
       
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    dX = dJnBoxCenter - dCyl3Ht
    
    stPoint.Set dX, 0, 0
    enPoint.Set dX - dUnionHeight, 0, 0
    
    'Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, pipeDiam
    
  
    
   
' Place Output 6 ( Heaxagonal nut)
'---------------------------------------------------------------------------------------------------------

     Dim ObjNut   As Object
     Dim geomFactory As IngrGeom3D.GeometryFactory
     Dim iCount            As Integer
     Dim dAngle             As Double

     Dim axisVect   As AutoMath.DVector
     Dim LineStrPoints(0 To 20)  As Double
     Dim oLineString As IngrGeom3D.LineString3d
     
'Assumption
    dAngle = PI / 3
    
    dX = dX - dUnionHeight
    
    stPoint.Set dX, 0, 0
    
'   First point is considered on positive Z-axis at an angle (PI / 6 + dAngle) from positive Y-axis.
     For iCount = 1 To 7
         LineStrPoints(3 * iCount - 3) = stPoint.x
         LineStrPoints(3 * iCount - 2) = stPoint.y + (dHexFacetoFace / 2) * Cos(PI / 6 + iCount * dAngle)
         LineStrPoints(3 * iCount - 1) = stPoint.z + (dHexFacetoFace / 2) * Sin(PI / 6 + iCount * dAngle)
     Next iCount
     
     Set ObjNut = New IngrGeom3D.Projection3d
     Set geomFactory = New IngrGeom3D.GeometryFactory
     Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
     Set axisVect = New AutoMath.DVector
     axisVect.Set -1, 0, 0
     
'    Assumption: Nut Height is Equal to Face to Face of Hexagonal Nut
     Set ObjNut = PlaceProjection(m_OutputColl, oLineString, axisVect, dHexFacetoFace, True)

' Set the output 1
     iOutput = iOutput + 1
     m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNut
     Set ObjNut = Nothing
     Set axisVect = Nothing
     Set geomFactory = Nothing
     Set oLineString = Nothing
     
   
' Place Output 7(Conical Part)
' -------------------------------------------------------------------------------------------------------------------
    Dim dConeLength As Double
'Assumption:
    dConeLength = parInstrumentLength / 100
  
    stPoint.Set 0, 0, 0
    enPoint.Set -dConeLength, 0, 0
  
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCone arrayOfOutputs(iOutput), stPoint, enPoint, pipeDiam, 0.9 * pipeDiam
    

    
'Place Output 8(Thermo Well End)
' -------------------------------------------------------------------------------------------------------------------

    stPoint.Set -dConeLength, 0, 0
    enPoint.Set -parInstrumentLength1, 0, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 0.9 * pipeDiam
    Set stPoint = Nothing
    Set enPoint = Nothing
    
  
'Place Output 9(Nozzle)
' -------------------------------------------------------------------------------------------------------------------

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    Dim dNozzleLength As Double
    dNozzleLength = dX - dHexFacetoFace
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    dX = -sptOffset + depth
    oPlacePoint.Set dX, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
 
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub


