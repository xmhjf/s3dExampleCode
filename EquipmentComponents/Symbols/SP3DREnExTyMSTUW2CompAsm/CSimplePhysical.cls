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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:25 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:        MS
'   Creation Date: Wednesday, Apr 7 2004
'   Description:
'    This is Exchanger Rear End type  M/S/T/U/W2 . Symbol details are taken from PDS Equipment Modeling User's Guide,
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
    
    Dim parFrontEndLength1 As Double
    Dim parExchangerLength As Double
    Dim parRearEndFlangeDia As Double
    Dim parRearEndFlangeTk1 As Double
    Dim parRearEndLength As Double
    Dim parChannelDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFrontEndLength1 = arrayOfInputs(2)
    parExchangerLength = arrayOfInputs(3)
    parRearEndFlangeDia = arrayOfInputs(4) 'P40
    parRearEndFlangeTk1 = arrayOfInputs(5) 'P41
    parRearEndLength = arrayOfInputs(6) 'P42
    parChannelDiameter = arrayOfInputs(7) 'P43
    parInsulationThickness = arrayOfInputs(8)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
'   Origin is taken on axis of Exchanger at (parFrontEndLength1 + parExchangerLength) from the left face of the Rear end

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

 ' Insert your code for output 1(ExchRearEndRHSFlange)
    stPoint.Set (parFrontEndLength1 + parExchangerLength), 0, 0
    enPoint.Set (parFrontEndLength1 + parExchangerLength) + parRearEndFlangeTk1, 0, 0
                     
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parRearEndFlangeDia
                         
' Insert your code for output 2(ExRearEndBody)
    stPoint.Set (parFrontEndLength1 + parExchangerLength) + parRearEndFlangeTk1, 0, 0
    enPoint.Set (parFrontEndLength1 + parExchangerLength) + parRearEndLength - dDomeHeight, 0, 0
                             
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parChannelDiameter
   
' Insert your code for output 3(ExRearEndDome)
    'Revolve Left Elliptical Dome Curve about X-Axis
    cenX = (parFrontEndLength1 + parExchangerLength) + parRearEndLength - dDomeHeight
    cenY = 0
    cenZ = 0
    MajorX = 0
    MajorY = 0
    MajorZ = parChannelDiameter / 2
    MMRatio = 0.5
    StartAngle = 1.5 * PI
    sweepangle = PI / 2
    norX = 0
    norY = -1
    norZ = 0

    Set oExchEndDomeArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                           cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, MMRatio, _
                           StartAngle, sweepangle)

    axisVect.Set 1, 0, 0
    ExchEndRevCenPt.Set (parFrontEndLength1 + parExchangerLength) + parRearEndLength - dDomeHeight, 0, 0

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

