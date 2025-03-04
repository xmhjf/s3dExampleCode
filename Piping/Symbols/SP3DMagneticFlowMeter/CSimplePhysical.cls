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
'   Creation Date:  Thursday, Jul 19 2006
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      Source: Rosemount 8700 Series, http://www.emersonprocess.com/rosemount/document/pds/8700.pdf
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
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
    Dim parInstrumentLength As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentHeight1 As Double
    
    Dim iOutput     As Double
    iOutput = 0
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInstrumentHeight = arrayOfInputs(3)
    parInstrumentLength = arrayOfInputs(4)
    parInstrumentWidth = arrayOfInputs(5)
    parInstrumentHeight1 = arrayOfInputs(6)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
     
    
' Place Body of the Valve (Output 1)
' -------------------------------------------------------------------------------------------------------------------
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
    Dim dBodyLength As Double
    Dim dBodyDia As Double
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
' Assumption:Length of the body is assumed to be 50% of the Instrument Length
    dBodyLength = parFacetoFace / 2
    dBodyDia = 1.8 * (parInstrumentHeight - parInstrumentHeight1 - (10 * parInstrumentHeight1 / 13))

    stPoint.Set -dBodyLength / 2, 0, 0
    enPoint.Set dBodyLength / 2, 0, 0
    
'Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dBodyDia
    
 
    
' Place Cylinder 1 (Output 2)
' -------------------------------------------------------------------------------------------------------------------
    Dim dCyl1Ht As Double
    Dim dCyl1Dia As Double

' Assumption: The Height of the Cylinder 1 is assumed to be 60% of Body Diameter
'             The Diameter of the cylinder is assumed to be 50% of Instrument Width

    dCyl1Ht = parInstrumentHeight - parInstrumentHeight1 - (10 * parInstrumentHeight1 / 13)
    dCyl1Dia = parInstrumentWidth / 2
    
    stPoint.Set 0, 0, 0
    enPoint.Set 0, dCyl1Ht, 0
    
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCyl1Dia
    
    
' Place Box 1 (Output 3)
' -------------------------------------------------------------------------------------------------------------------
    Dim dBox1Ht As Double
    Dim dBox1Length As Double
    Dim dBox1Width As Double
    
'Assumption: The Length of Box1 is assumed to be 90% of the Cylinder1 Diameter
'            The Width of Box1 is assumed to be 50% of the Cylinder1 Diameter

    dBox1Ht = 9 * parInstrumentHeight1 / 13
                                                
                                                
    dBox1Length = 0.9 * dCyl1Dia
    dBox1Width = 0.5 * dCyl1Dia
        
    stPoint.Set -dBox1Length / 2, dCyl1Ht, dBox1Width / 2
    enPoint.Set dBox1Length / 2, dCyl1Ht + dBox1Ht, -dBox1Width / 2
    
    Dim ObjBox1 As Object
    Set ObjBox1 = New IngrGeom3D.Projection3d
    Set ObjBox1 = PlaceBox(m_OutputColl, stPoint, enPoint)
    
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox1
    Set ObjBox1 = Nothing
    
 

' Place Box 2 (Output 4)
' -------------------------------------------------------------------------------------------------------------------
    Dim dBox2Ht As Double
    Dim dBox2Length As Double
    Dim dBox2Width As Double
    
'Assumption:
    dBox2Ht = 5 * parInstrumentHeight1 / 13
    dBox2Width = 0.9 * parInstrumentWidth
    dBox2Length = dBox2Width
    
    stPoint.Set -dBox2Length / 2, dCyl1Ht + dBox1Ht, dBox2Width / 2
    enPoint.Set dBox2Length / 2, dCyl1Ht + dBox1Ht + dBox2Ht, -dBox2Width / 2
    
    Dim ObjBox2 As Object
    Set ObjBox2 = New IngrGeom3D.Projection3d
    Set ObjBox2 = PlaceBox(m_OutputColl, stPoint, enPoint)
    
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox2
    Set ObjBox2 = Nothing

' Place Cylinder 2 (Output 5)
' -------------------------------------------------------------------------------------------------------------------
    Dim dCyl2Ht As Double
    
'Assumption
    dCyl2Ht = 9 * parInstrumentHeight1 / 13
    
    stPoint.Set 0, dCyl1Ht + dBox1Ht + dBox2Ht, 0
    enPoint.Set 0, parInstrumentHeight, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parInstrumentWidth

   
' Place Conduit Cylinder 1 (Output 6)
' -------------------------------------------------------------------------------------------------------------------
    Dim dCondCylLnth As Double
    Dim dCondCylPosY As Double
    Dim dCondCylPosZ As Double
    Dim dCondCylDia As Double
    
'Assumption
    dCondCylPosY = dCyl1Ht + dBox1Ht + dBox2Ht
    dCondCylPosZ = parInstrumentWidth / 4.02
    dCondCylLnth = parInstrumentLength - (parInstrumentWidth / 2)
    dCondCylDia = 0.02667
    
    stPoint.Set 0, dCondCylPosY, dCondCylPosZ
    enPoint.Set -dCondCylLnth, dCondCylPosY, dCondCylPosZ
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCondCylDia
    
  

' Place Conduit Cylinder 2 (Output 7)
' -------------------------------------------------------------------------------------------------------------------
    
    stPoint.Set 0, dCondCylPosY, -dCondCylPosZ
    enPoint.Set -dCondCylLnth, dCondCylPosY, -dCondCylPosZ
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dCondCylDia
    
    Set stPoint = Nothing
    Set enPoint = Nothing

    

' Place Nozzle 1 (Output 8)
' -------------------------------------------------------------------------------------------------------------------
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim dNozzleLength As Double
        
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
' Assumption
    dNozzleLength = parFacetoFace / 4
    
    oPlacePoint.Set -parFacetoFace / 2, 0, 0
    oDir.Set -1, 0, 0
    
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    

' Place Nozzle 2 (Output 9)
' -------------------------------------------------------------------------------------------------------------------
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set parFacetoFace / 2, 0, 0
    oDir.Set 1, 0, 0

    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
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


