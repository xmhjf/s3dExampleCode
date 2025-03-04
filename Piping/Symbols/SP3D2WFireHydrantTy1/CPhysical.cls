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
'   Author:         MS
'   Creation Date:  Thursday, Sep 9 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: FS10a Page No. D-88 PDS Piping Component Data Reference Guide.
'    SP3D2WFireHydrantTy1  Symbol is created with Four Outputs
'    The five physical aspect outputs are created as follows:
'    ObjHydrantBody, ObjWaterOutlet1, ObjWaterOutlet2 - Using 'PlaceCylinder' function and
'    One ObjNozzle object by using 'CreateNozzleWithLength' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
        
    Dim parFacetoCenter As Double
    Dim parBottomtoCenter As Double
    Dim parGradetoCenter As Double
    Dim parWaterOutlettoGrade As Double
    Dim parWaterOutlettoTop As Double
    Dim parHydrantDiameter As Double
    Dim parWaterOutlettoCenter As Double
    Dim parWaterOutletDiameter As Double
    Dim parRotation As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)              'D1
    parBottomtoCenter = arrayOfInputs(3)            'D2
    parGradetoCenter = arrayOfInputs(4)             'D3
    parWaterOutlettoGrade = arrayOfInputs(5)    'D4
    parWaterOutlettoTop = arrayOfInputs(6)      'D5
    parHydrantDiameter = arrayOfInputs(7)       'D6
    parWaterOutlettoCenter = arrayOfInputs(8)   'D7
    parWaterOutletDiameter = arrayOfInputs(9)   'D8
    parRotation = arrayOfInputs(10)   'D9

    iOutput = 0

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
' Insert your code for output 1(HydrantBody)
    stPoint.Set 0, 0, parGradetoCenter + parWaterOutlettoGrade + parWaterOutlettoTop
    enPoint.Set 0, 0, -parBottomtoCenter
    
    Set ObjHydrantBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parHydrantDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHydrantBody
    Set ObjHydrantBody = Nothing
    
 ' Insert your code for output 2(WaterOutlet1)
    Dim RefAngle As Double
    Dim HydrantHeight As Double
    Dim CylinderMeetPoint As Double

    HydrantHeight = parGradetoCenter + parWaterOutlettoGrade 'D3+D4
    RefAngle = parRotation - (PI / 4)
    CylinderMeetPoint = Sqr((parHydrantDiameter / 2) * (parHydrantDiameter / 2) - (parWaterOutletDiameter / 2) * (parWaterOutletDiameter / 2))
 
     stPoint.x = parWaterOutlettoCenter * Cos(RefAngle)
     stPoint.y = -parWaterOutlettoCenter * Sin(RefAngle)
    stPoint.z = HydrantHeight
    
    enPoint.x = CylinderMeetPoint * Cos(RefAngle)
    enPoint.y = -CylinderMeetPoint * Sin(RefAngle)
    enPoint.z = HydrantHeight

    Set ObjWaterOutlet1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWaterOutletDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWaterOutlet1
    Set ObjWaterOutlet1 = Nothing
     
' Insert your code for output 3(WaterOutlet2)
    RefAngle = parRotation + (PI / 4)
  
     stPoint.x = parWaterOutlettoCenter * Cos(RefAngle)
     stPoint.y = -parWaterOutlettoCenter * Sin(RefAngle)
    stPoint.z = HydrantHeight
    
    enPoint.x = CylinderMeetPoint * Cos(RefAngle)
    enPoint.y = -CylinderMeetPoint * Sin(RefAngle)
    enPoint.z = HydrantHeight

    Set ObjWaterOutlet2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWaterOutletDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjWaterOutlet2
    Set ObjWaterOutlet2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Place Nozzle 1
     RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
 
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFacetoCenter)

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


