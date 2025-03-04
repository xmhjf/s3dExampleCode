VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Wednesday, March 27 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'This Symbol is Constructed using 11 number of Outputs (Physical and Insulation)
'Physical Aspect: 1.port-1 with length, 2.port-2 with length,3.port-3 with length and 4.Return body
'Insulation Aspect: Closed Insulation is provided.
'1.Port -1 Insulation, 2.Port-1 Tangent Insulation,3.Port -2 Insulation, 4.Port-2 Tangent Insulation,
'5.Port -3 Insulation, 6.Port-3 Tangent Insulation, and 7.Insulation for return
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI           As Double


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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjSP3DReturn As Object
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsDiameter      As Double
    Dim parInsulationRadius     As Double
    Dim parFacetoFace As Double
    Dim parFace1toFace3 As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parFacetoCenter = arrayOfInputs(3)
    parFace1toFace3 = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0

' Insert your code for output 5(Insulated Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim objCirc1 As IngrGeom3D.Circle3d
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    Dim centPoint   As New AutoMath.DPosition
    Dim axis        As New AutoMath.DVector
    
    cirNormalX = -1
    cirNormalY = 0
    cirNormalZ = 0
    cir1CenterX = CenterPos.x - parFacetoFace / 2
    cir1CenterY = CenterPos.y
    cir1CenterZ = CenterPos.z
    parInsulationRadius = pipeDiam / 2 + parInsulationThickness
    If parInsulationRadius > (parFacetoFace / 2) Then parInsulationRadius = (parFacetoFace / 2)
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                m_OutputColl.ResourceManager, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, parInsulationRadius)
    
    centPoint.Set CenterPos.x - parFacetoFace / 2, _
                  CenterPos.y + parFacetoFace / 2, _
                  CenterPos.z
    axis.Set 0, 0, 1
    Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCirc1, _
                                    axis, centPoint, PI, _
                                    True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set geomFactory = Nothing
    Set axis = Nothing
    Set centPoint = Nothing
    
'''''' Insert your code for output 3 (Insulated Port1)
'''''    stPoint.Set CenterPos.x - parFacetoCenter, CenterPos.y, CenterPos.z
'''''    enPoint.Set CenterPos.x - (parFacetoCenter - flangeThick - parInsulationThickness), _
'''''                CenterPos.y, CenterPos.z
'''''    parInsDiameter = flangeDiam + parInsulationThickness * 2
'''''    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameter, True)
''''''   Set the output
'''''    iOutput = iOutput + 1
'''''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
'''''    Set ObjInsulatedPort1 = Nothing

' Insert your code for output 6 (Insulated Port1 Tangent)
    Dim ObjInsTangent1  As Object
    stPoint.Set CenterPos.x - parFacetoCenter + flangeThick, CenterPos.y, CenterPos.z
    enPoint.Set CenterPos.x - (parFacetoCenter - parFacetoFace / 2), _
                CenterPos.y, CenterPos.z
    parInsDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsTangent1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTangent1
    Set ObjInsTangent1 = Nothing
        
''''''' Insert your code for output 7(Insulated Port2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
'''''''    stPoint.Set CenterPos.x - parFacetoCenter, _
'''''''                CenterPos.y + parFacetoFace, CenterPos.z
'''''''    enPoint.Set CenterPos.x - (parFacetoCenter - flangeThick - parInsulationThickness), _
'''''''                CenterPos.y + parFacetoFace, CenterPos.z
'''''''    parInsDiameter = flangeDiam + parInsulationThickness * 2
'''''''    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameter, True)
''''''''   Set the output
'''''''    iOutput = iOutput + 1
'''''''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
'''''''    Set ObjInsulatedPort2 = Nothing

' Insert your code for output 7(Insulated Port2 Tangent)
    Dim ObjInsTangent2 As Object
    stPoint.Set CenterPos.x - parFacetoCenter + flangeThick, _
                CenterPos.y + parFacetoFace, CenterPos.z
    enPoint.Set CenterPos.x - (parFacetoCenter - parFacetoFace / 2), _
                CenterPos.y + parFacetoFace, CenterPos.z
    parInsDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsTangent2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTangent2
    Set ObjInsTangent2 = Nothing

''''''''' Insert your code for output 4(Insulated Port3)
''''''''    Dim ObjInsulatedPort3 As Object
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
''''''''    stPoint.Set CenterPos.x + (parFace1toFace3 - parFacetoCenter), _
''''''''                CenterPos.y + parFacetoFace / 2, CenterPos.z
''''''''    enPoint.Set stPoint.x - flangeThick - parInsulationThickness, _
''''''''                stPoint.y, CenterPos.z
''''''''    parInsDiameter = flangeDiam + parInsulationThickness * 2
''''''''    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameter, True)
'''''''''   Set the output
''''''''    iOutput = iOutput + 1
''''''''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
''''''''    Set ObjInsulatedPort3 = Nothing

' Insert your code for output 8(Insulated Port3 Tangent)
    Dim ObjInsTangent3 As Object
    stPoint.Set CenterPos.x, _
                CenterPos.y + parFacetoFace / 2, CenterPos.z
    enPoint.Set CenterPos.x + (parFace1toFace3 - parFacetoCenter - flangeThick), _
                CenterPos.y + parFacetoFace / 2, CenterPos.z
    parInsDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsTangent3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTangent3
    Set ObjInsTangent3 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
' remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCirc1
    Set objCirc1 = Nothing
    sp3dElem1.Remove
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
