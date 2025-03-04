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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         Mallepally Srinivas
'   Creation Date:  Thursday, Sep 9 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: FS11b Page No. D-89 PDS Piping Component Data Reference Guide.
'    Symbol is created with Seven Outputs
'    The five physical aspect outputs are created as follows:
'    ObjHydrantBody , ObjWaterOutlet1, ObjWaterOutlet2, ObjMonitorHorzBody, ObjMonitorFlange - using 'PlaceCylinder' function and
'    ObjMonitorElbow using 'PlaceRevolution' function.
'    One ObjNozzle object by using 'CreateNozzle' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   01.Aug.2006     Kishore(kkkottur)       TR: 80254 Replaced 'm_OutputColl.ResourceManager' with "Nothing'.

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



Option Explicit
Private PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
PI = 4 * Atn(1)
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
    
    Dim iOutput     As Double
    Dim ObjHydrantBody As Object
    Dim ObjWaterOutlet1 As Object
    Dim ObjWaterOutlet2 As Object
    Dim ObjMonitorHorzBody As Object
    Dim ObjMonitorElbow As Object
    Dim ObjMonitorFlange As Object
    
    Dim parFace1toCenter As Double
    Dim parWaterOutlettoTop As Double
    Dim parHydrantDiameter As Double
    Dim parFlangeOutlettoCenter As Double
    Dim parFace2toCenter As Double
    Dim parWaterOutlettoCenter As Double
    Dim parWaterOutletDiameter As Double
    Dim parMonitorOutletDiameter As Double
    Dim parFlangeWidth As Double
    Dim parFlangeDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D1
    parWaterOutlettoTop = arrayOfInputs(3)  'D2
    parHydrantDiameter = arrayOfInputs(4)   'D3
    parFlangeOutlettoCenter = arrayOfInputs(5)  'D4
    parFace2toCenter = arrayOfInputs(6) 'D5
    parWaterOutlettoCenter = arrayOfInputs(7)   'D6
    parWaterOutletDiameter = arrayOfInputs(8)   'D7
    parMonitorOutletDiameter = arrayOfInputs(9) 'D8
    parFlangeWidth = arrayOfInputs(10)  'D9
    parFlangeDiameter = arrayOfInputs(11)   'D10
    

    iOutput = 0
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

 ' Insert your code for output 1(Hydrant Body)
    stPoint.Set 0, 0, -parFace1toCenter
    enPoint.Set 0, 0, parWaterOutlettoTop
    
    Set ObjHydrantBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parHydrantDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHydrantBody
    Set ObjHydrantBody = Nothing
    
 ' Insert your code for output 2(Water Outlet1)
    Dim RefAngle As Double
    Dim CylinderMeetPoint As Double
    
    RefAngle = 3 * (PI / 4)
    CylinderMeetPoint = Sqr((parHydrantDiameter / 2) * (parHydrantDiameter / 2) - (parWaterOutletDiameter / 2) * (parWaterOutletDiameter / 2))
    
    stPoint.x = parWaterOutlettoCenter * Cos(RefAngle)
    stPoint.y = parWaterOutlettoCenter * Sin(RefAngle)
    stPoint.z = 0
    
    enPoint.x = CylinderMeetPoint * Cos(RefAngle)
    enPoint.y = CylinderMeetPoint * Sin(RefAngle)
    enPoint.z = 0
    
    Set ObjWaterOutlet1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWaterOutletDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWaterOutlet1
    Set ObjWaterOutlet1 = Nothing
    
 ' Insert your code for output 3(Water Outlet2)
    RefAngle = 5 * (PI / 4)
    
    stPoint.x = parWaterOutlettoCenter * Cos(RefAngle)
    stPoint.y = parWaterOutlettoCenter * Sin(RefAngle)
    stPoint.z = 0
    
    enPoint.x = CylinderMeetPoint * Cos(RefAngle)
    enPoint.y = CylinderMeetPoint * Sin(RefAngle)
    enPoint.z = 0
    
    Set ObjWaterOutlet2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWaterOutletDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWaterOutlet2
    Set ObjWaterOutlet2 = Nothing
    
 ' Insert your code for output 4(Monitor Horzontal Body)
    stPoint.Set 0, 0, 0
    enPoint.Set parFlangeOutlettoCenter - parFace2toCenter + parFlangeWidth, 0, 0
    
    Set ObjMonitorHorzBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parMonitorOutletDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMonitorHorzBody
    Set ObjMonitorHorzBody = Nothing
    
 ' Insert your code for output 5(Monitor Elbow)
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim objCirc1 As IngrGeom3D.Circle3d
    Dim cirNormalX As Double
    Dim cirNormalY As Double
    Dim cirNormalZ As Double
    Dim cir1CenterX As Double
    Dim cir1CenterY As Double
    Dim cir1CenterZ As Double
    Dim centPoint   As New AutoMath.DPosition
    Dim axis        As New AutoMath.DVector
    
    cirNormalX = 0
    cirNormalY = 0
    cirNormalZ = 1
        
    cir1CenterX = parFlangeOutlettoCenter
    cir1CenterY = 0
    cir1CenterZ = parFace2toCenter - parFlangeWidth
    
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                        Nothing, cir1CenterX, _
                                        cir1CenterY, cir1CenterZ, cirNormalX, _
                                        cirNormalY, cirNormalZ, parMonitorOutletDiameter / 2)
    
    centPoint.Set parFlangeOutlettoCenter - parFace2toCenter + parFlangeWidth, 0, parFace2toCenter - parFlangeWidth
    axis.Set 0, 1, 0
    
    Set ObjMonitorElbow = PlaceRevolution(m_OutputColl, objCirc1, axis, centPoint, PI / 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMonitorElbow
    Set ObjMonitorElbow = Nothing
    Set geomFactory = Nothing
    Set centPoint = Nothing
    Set axis = Nothing
' remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCirc1
    Set objCirc1 = Nothing
    sp3dElem1.Remove
    
 ' Insert your code for output 6(Monitor Flange)
    stPoint.Set parFlangeOutlettoCenter, 0, parFace2toCenter
    enPoint.Set parFlangeOutlettoCenter, 0, parFace2toCenter - parFlangeWidth
    
    Set ObjMonitorFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFlangeDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMonitorFlange
    Set ObjMonitorFlange = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set 0, 0, -parFace1toCenter - sptOffset + depth
    oDir.Set 0, 0, -1
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

