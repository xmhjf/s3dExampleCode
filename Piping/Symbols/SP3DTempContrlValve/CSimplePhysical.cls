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
'   Creation Date:  Thursday, Jul 24 2006
'   Description:
'    Physical Aspect of Temperature Control Valve, Type 56-T (at http://www.sterlco.com/images_products/files/56-T%20Series1.pdf)
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
       
'Inputs
    Dim parFacetoFace As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    
    Dim iOutput     As Double
    iOutput = 1
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInstrumentHeight = arrayOfInputs(3)
    parInstrumentHeight1 = arrayOfInputs(4)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    
    
' Place Output 1 and Output 2 ( Hexagonal Nut 1 and Hexagonal Nut 2)
'---------------------------------------------------------------------------------------------------------

     Dim ObjNut   As Object
     Dim geomFactory As IngrGeom3D.GeometryFactory
     Dim iCount As Integer
     Dim dAngle As Double
     Dim dNutHeight As Double
     Dim dNutWidth As Double

     Dim axisVect   As AutoMath.DVector
     Dim LineStrPoints(0 To 20)  As Double
     Dim oLineString As IngrGeom3D.LineString3d
     
     Set ObjNut = New IngrGeom3D.Projection3d
     Set geomFactory = New IngrGeom3D.GeometryFactory
     Set axisVect = New AutoMath.DVector
     
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

'Assumptions
    dAngle = PI / 3
    dNutHeight = parFacetoFace / 4
    dNutWidth = 1.5 * pipeDiam
    
   For iOutput = 1 To 2
       Select Case iOutput
        Case 1:
        stPoint.Set -parFacetoFace / 2, 0, 0
        axisVect.Set 1, 0, 0
        
        Case 2:
        stPoint.Set parFacetoFace / 2, 0, 0
        axisVect.Set -1, 0, 0
       End Select
        
    '   First point is considered on positive Z-axis at an angle (PI / 6 + dAngle) from positive Y-axis.
         
        For iCount = 1 To 7
            LineStrPoints(3 * iCount - 3) = stPoint.x
            LineStrPoints(3 * iCount - 2) = stPoint.y + (dNutWidth / 2) * Cos(PI / 6 + iCount * dAngle)
            LineStrPoints(3 * iCount - 1) = stPoint.z + (dNutWidth / 2) * Sin(PI / 6 + iCount * dAngle)
        Next iCount
         
         Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
         Set ObjNut = PlaceProjection(m_OutputColl, oLineString, axisVect, dNutHeight, True)
    
    ' Set the output
         m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNut
         
    Next iOutput
    
     iOutput = 2
     
     Set ObjNut = Nothing
     Set axisVect = Nothing
     Set geomFactory = Nothing
     Set oLineString = Nothing
     

    
' Place Output 3 ( Cylinder 1)
'---------------------------------------------------------------------------------------------------------
    Dim dCyl1PosX As Double
    dCyl1PosX = parFacetoFace / 2 - dNutHeight
    
    stPoint.Set dCyl1PosX, 0, 0
    enPoint.Set -dCyl1PosX, 0, 0

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, pipeDiam

    
' Place Output 4(Cylinder 2)
' -------------------------------------------------------------------------------------------------------------------
    Dim dCyl2PosY As Double
    Dim dCyl2Dia As Double
    
    dCyl2PosY = pipeDiam / 2
    dCyl2Dia = 1.1 * pipeDiam
    
    stPoint.Set 0, dCyl2PosY, 0
    enPoint.Set 0, -dCyl2PosY, 0

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCyl2Dia
    
' Place Output 5(Box 1)
' -------------------------------------------------------------------------------------------------------------------
    Dim dBox1Ht As Double
    Dim dBox1Width As Double
    
'Assumption:
    dBox1Ht = parInstrumentHeight / 6
    dBox1Width = 2 * dCyl1PosX

    stPoint.Set -dCyl1PosX, dCyl2PosY, dBox1Width / 2
    enPoint.Set dCyl1PosX, dCyl2PosY + dBox1Ht, -dBox1Width / 2
    
    Dim ObjBox1 As Object
    Set ObjBox1 = New IngrGeom3D.Projection3d
    Set ObjBox1 = PlaceBox(m_OutputColl, stPoint, enPoint)
    
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox1
    Set ObjBox1 = Nothing
    
    
' Place Output 6(Box 2)
' -------------------------------------------------------------------------------------------------------------------
    Dim dBox2Ht As Double
    Dim dBox2Length As Double
    Dim dBox2Width As Double
    
'Assumption:
    dBox2Ht = parInstrumentHeight1 / 5
    dBox2Length = 0.8 * parFacetoFace
    dBox2Width = dBox2Length

    stPoint.Set -dBox2Length / 2, -dCyl2PosY, dBox2Width / 2
    enPoint.Set dBox2Length / 2, -dCyl2PosY - dBox2Ht, -dBox2Width / 2
    
    Dim ObjBox2 As Object
    Set ObjBox2 = New IngrGeom3D.Projection3d
    Set ObjBox2 = PlaceBox(m_OutputColl, stPoint, enPoint)
    
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox2
    Set ObjBox2 = Nothing

    
' Place Output 7(Cylinder 3)
' -------------------------------------------------------------------------------------------------------------------
    Dim dCyl3Dia As Double
    Dim dCyl3Pos1Y As Double
    Dim dCyl3Pos2Y As Double

' Assumption:
    dCyl3Dia = 0.75 * dBox1Width
    dCyl3Pos1Y = dCyl2PosY + dBox1Ht
    dCyl3Pos2Y = parInstrumentHeight - (parInstrumentHeight / 5)
    
    stPoint.Set 0, dCyl3Pos1Y, 0
    enPoint.Set 0, dCyl3Pos2Y, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCyl3Dia
    
    
' Place Output 8(Cylinder 4)
' -------------------------------------------------------------------------------------------------------------------
    Dim dCyl4Dia As Double
    
    dCyl4Dia = dCyl3Dia / 5
    
    stPoint.Set 0, dCyl3Pos2Y, 0
    enPoint.Set 0, parInstrumentHeight, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCyl4Dia
    
    
' Place Output 9 (Cylinder 5)
' -------------------------------------------------------------------------------------------------------------------

    Dim dCyl5Dia As Double
    Dim dCyl5PosY As Double
    
' Assumption:
    dCyl5Dia = 0.75 * dBox2Width
    dCyl5PosY = -dCyl2PosY - dBox2Ht
    
    stPoint.Set 0, dCyl5PosY, 0
    enPoint.Set 0, -parInstrumentHeight1, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCyl5Dia
    
    
'Place Output 10(Nozzle 1)
' -------------------------------------------------------------------------------------------------------------------

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    Dim dNozzleLength As Double
    dNozzleLength = dNutHeight
    

    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    
    
'Place Output 11(Nozzle 2)
' -------------------------------------------------------------------------------------------------------------------
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
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


