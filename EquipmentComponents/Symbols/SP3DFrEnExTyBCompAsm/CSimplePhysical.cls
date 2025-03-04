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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:05:45 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:        MS
'   Creation Date: Wednesday, Apr 7 2004
'   Description:
'    This is Exchanger Front End type B. Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E319 Symbol in Page no 310.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

    Private m_oSymGeomHelper As IJSymbolGeometryHelper
    Private PI As Double
    Private Const MODULE = "SimplePhysical:" 'Used for error messages
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
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double

    Dim ObjExchEndDome As Object

    Dim parFrontEndFlangeDia As Double
    Dim parFrontEndFlangeTk1 As Double
    Dim parFrontEndLength1 As Double
    Dim parFrontEndLength2 As Double
    Dim parChannelDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFrontEndFlangeDia = arrayOfInputs(2) 'P30
    parFrontEndFlangeTk1 = arrayOfInputs(3) 'P31
    parFrontEndLength1 = arrayOfInputs(4) 'P32
    parFrontEndLength2 = arrayOfInputs(5) 'P33
    parChannelDiameter = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
     
    
    iOutput = 0
'   Origin is taken on axis of Exchanger End, at the junction of P32 and P33 dimensions.

    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    Dim dDomeHeight As Double

'Dome height is taken as 1/4th of Exchanger Neck Diameter
    dDomeHeight = parChannelDiameter / 4

'   Variables needed for the Dome Ellipse revolution
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oExchEndDomeArc As IngrGeom3D.EllipticalArc3d
    Dim cenX As Double
    Dim cenY As Double
    Dim cenZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim MMRatio As Double
    Dim StartAngle As Double
    Dim sweepangle As Double
    Dim norX As Double
    Dim norY As Double
    Dim norZ As Double
    Dim axisVect As AutoMath.DVector
    Dim ExchEndRevCenPt As AutoMath.DPosition

    Set axisVect = New AutoMath.DVector
    Set ExchEndRevCenPt = New AutoMath.DPosition

 ' Insert your code for output 1(ExchFrontEndRHSFlange)
   'X-coordinate: P32
    stPoint.Set parFrontEndLength1, 0, 0

    'X-coordinate: P32 - P31
    enPoint.Set parFrontEndLength1 - parFrontEndFlangeTk1, 0, 0
                     
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parFrontEndFlangeDia
                         
' Insert your code for output 2(ExFrontEndBody)
   'X-coordinate: P32 - P31
    stPoint.Set parFrontEndLength1 - parFrontEndFlangeTk1, 0, 0

   'X-coordinate: -P33 + Dome height
    enPoint.Set -parFrontEndLength2 + dDomeHeight, 0, 0
                             
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parChannelDiameter
   
' Insert your code for output 3(ExFrontEndDome)
    'Revolve Left Elliptical Dome Curve about X-Axis
    'X-coordinate: -P33 + Dome height
    cenX = -parFrontEndLength2 + dDomeHeight
    cenY = 0
    cenZ = 0
    MajorX = 0
    MajorY = 0
    MajorZ = parChannelDiameter / 2
    MMRatio = 0.5
    StartAngle = PI / 2
    sweepangle = PI / 2
    norX = 0
    norY = -1
    norZ = 0

    Set oExchEndDomeArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                           cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, MMRatio, _
                           StartAngle, sweepangle)

    axisVect.Set 1, 0, 0
    ExchEndRevCenPt.Set -parFrontEndLength2 + dDomeHeight, 0, 0

    Set ObjExchEndDome = PlaceRevolution(m_OutputColl, oExchEndDomeArc, axisVect, ExchEndRevCenPt, 2 * PI, True)

   'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchEndDome
    Set ObjExchEndDome = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oGeomFactory = Nothing
    Set axisVect = Nothing
    Set ExchEndRevCenPt = Nothing
    Set oExchEndDomeArc = Nothing
      
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

