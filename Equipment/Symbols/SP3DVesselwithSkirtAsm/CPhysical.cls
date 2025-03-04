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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:32 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          JF
'   Creation Date:  Thursday, Mar 21 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO         Made compatible with Smart Occurrence based Equipments
'   20.May.2005     MS      CR-76070:   Modified the nozzle creation code(as members),added datum points.
'   08.SEP.2006     KKC     DI-95670    Replace names with initials in all revision history sheets and symbols
'   05.MAR.2008     PK      TR-136972   Code for creation of foundation port from place holder has been removed
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

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
    Dim CenterPos       As New AutoMath.DPosition
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjSupportSkirt As Object
    Dim ObjTank As Object
    Dim ObjTopDome As Object
    Dim ObjBottomDome As Object
    Dim ObjOperationEnv As Object
    Dim ObjMaintenanceEnv As Object
    Dim ObjInsulationEnv As Object
    
    Dim parSkirtHeight As Double
    Dim parSkirtDiameter As Double
    Dim parVesselDiameter As Double
    Dim parVesselHeight As Double
    Dim parDomeHeight As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSkirtHeight = arrayOfInputs(2)
    parSkirtDiameter = arrayOfInputs(3)
    parVesselDiameter = arrayOfInputs(4)
    parVesselHeight = arrayOfInputs(5)
    parDomeHeight = arrayOfInputs(6)

    
    iOutput = 0

 ' Insert your code for output 1(Cone)
 
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    stPoint.Set 0, 0, 0 'X, Y, Z
    enPoint.Set 0, 0, parSkirtHeight
    Set ObjSupportSkirt = PlaceCone(m_OutputColl, stPoint, enPoint, (parSkirtDiameter / 2), (parVesselDiameter / 2))

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupportSkirt
    Set ObjSupportSkirt = Nothing
    
 ' Insert your code for output 2(Cylinder)

    stPoint.Set 0, 0, parSkirtHeight
    enPoint.Set 0, 0, (parSkirtHeight + parVesselHeight)
    Set ObjTank = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTank
    Set ObjTank = Nothing
    
 ' Insert your code for output 3(Top Dome: Revolved quarter Ellipse)
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim EllipNormX As Double, EllipNormY As Double, EllipNormZ As Double
    Dim MajorX As Double, MajorY As Double, MajorZ As Double
    Dim axis    As New AutoMath.DVector
    Dim PI As Double
    Dim parRatio As Double
    Dim oDomeEllip As IngrGeom3D.EllipticalArc3d
    Dim pDomeCenter As New AutoMath.DPosition
    Dim SP3DElem As IJDObject
    Set geomFactory = New IngrGeom3D.GeometryFactory
       
    PI = 4 * Atn(1)
    
    axis.Set 0, 0, 1
    MajorX = parVesselDiameter / 2
    MajorY = 0
    MajorZ = 0
'The Ratio of minor axis over major axis
    parRatio = parDomeHeight / (parVesselDiameter / 2)
'Normal vector to plane of ellipce)
    EllipNormX = 0
    EllipNormY = -1
    EllipNormZ = 0

    
    
    pDomeCenter.Set 0, 0, parSkirtHeight + parVesselHeight
     
    Set oDomeEllip = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            m_OutputColl.ResourceManager, pDomeCenter.x, _
                                            pDomeCenter.y, pDomeCenter.z, EllipNormX, _
                                            EllipNormY, EllipNormZ, MajorX, MajorY, MajorZ, _
                                            parRatio, 0, PI / 2)


    Set ObjTopDome = PlaceRevolution(m_OutputColl, oDomeEllip, axis, pDomeCenter, PI * 2, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopDome
    Set ObjTopDome = Nothing
'Delete the elliptical arc that was created for the revolution
    Set SP3DElem = oDomeEllip
    Set oDomeEllip = Nothing
    SP3DElem.Remove
' Insert your code for output 4(Bottom Dome: Revolved quarter Ellipse)
    
    'Normal vector to plane of ellipce
    EllipNormX = 0
    EllipNormY = 1
    EllipNormZ = 0

    pDomeCenter.Set 0, 0, parSkirtHeight
    Set oDomeEllip = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                            m_OutputColl.ResourceManager, pDomeCenter.x, _
                                            pDomeCenter.y, pDomeCenter.z, EllipNormX, _
                                            EllipNormY, EllipNormZ, MajorX, MajorY, MajorZ, _
                                            parRatio, 0, PI / 2)

    
    Set ObjBottomDome = PlaceRevolution(m_OutputColl, oDomeEllip, axis, pDomeCenter, PI * 2, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomDome
    Set ObjBottomDome = Nothing
    Set geomFactory = Nothing
'Delete the elliptical arc that was created for the revolution
    Set SP3DElem = oDomeEllip
    Set oDomeEllip = Nothing
    SP3DElem.Remove
    

'=============================================
'Construction of Equipment Foundation Port 8
'=============================================
'    Dim NozzlePHFactory As New NozzlePHFactory
'    Dim oEqpFoundationPort As IJEqpFoundationPort
'
'    Set oEqpFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "VesselFoundationPort", _
'                                                False, m_OutputColl.ResourceManager)
'
'    Call oEqpFoundationPort.PutCS(0#, 0#, 0#, 1#, 0#, 0#, 0#, 0#, -1#)
'
'    ' Set the output
'    iOutput = iOutput + 1
'    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oEqpFoundationPort
'    Set oEqpFoundationPort = Nothing
    
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0, , , True)
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "VesselwithSkirtControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing


    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


