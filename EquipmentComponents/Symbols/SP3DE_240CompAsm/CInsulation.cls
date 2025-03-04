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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:52 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CInsulation.cls
'   Author:         ACM
'   Creation Date:  Wednesday, Feb 25, 2004
'   Description:
'   This Symbol detail is taken from PDS Equipment Modeling User's Guide,
'   E240 Symbol in Page no 295.  Physical and Insulation aspect outputs are variable outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006     kkc       DI 95670-Replaced names with initials in the revision history.
'   26.Sep.2006     dkl       TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double

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
    Dim ObjInsulatedVessel As Object
    
    Dim parVesselLength As Double
    Dim parVesselDiameter As Double
    Dim parBootHorLocation As Double
    Dim parBootVertLocation As Double
    Dim parBootDiameter As Double
    Dim parBootAngularLocation As Double
    Dim parVesselCenterHeight As Double
    Dim parFirstSupportLocation As Double
    Dim parSecondSupportLocation As Double
    Dim parThirdSupportLocation As Double
    Dim parSupportThickness As Double
    Dim parSupportLength As Double
    Dim parStiffenerRadius As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim InsulationDia As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselLength = arrayOfInputs(2)  'P2
    parVesselDiameter = arrayOfInputs(3)        'P3
    parBootHorLocation = arrayOfInputs(4)       'P4
    parBootVertLocation = arrayOfInputs(5)      'P5
    parBootDiameter = arrayOfInputs(6)          'P6
    parBootAngularLocation = arrayOfInputs(7)   'P7
    parVesselCenterHeight = arrayOfInputs(8)    'P8
    parFirstSupportLocation = arrayOfInputs(9)  'P9
    parSecondSupportLocation = arrayOfInputs(10) 'P10
    parThirdSupportLocation = arrayOfInputs(11) 'P11
    parSupportThickness = arrayOfInputs(12)     'P12
    parSupportLength = arrayOfInputs(13)        'P13
    parStiffenerRadius = arrayOfInputs(14)      'P14
    parInsulationThickness = arrayOfInputs(15)
    
    iOutput = 0

' Insert your code for output 4(InsulatedVessel)
    stPoint.Set 0, 0, 0
    enPoint.Set parVesselLength, 0, 0
    InsulationDia = parVesselDiameter + 2 * parInsulationThickness
    Set ObjInsulatedVessel = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedVessel
    Set ObjInsulatedVessel = Nothing

' Insert your code for output 5(Insulation for Left Hand Side Dome)
    Dim axisVect As New AutoMath.DVector
    Dim RevCenPt As New AutoMath.DPosition
    
    Dim cenX As Double
    Dim cenY As Double
    Dim cenZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim mMRatio As Double
    Dim StartAngle As Double
    Dim SweepAngle As Double
    Dim normalX As Double
    Dim normalY As Double
    Dim normalZ As Double
    
    Dim oEqpLeftArcI As IngrGeom3D.EllipticalArc3d
    Dim oEqpRightArcI As IngrGeom3D.EllipticalArc3d
    
    cenX = 0
    cenY = 0
    cenZ = 0
    MajorX = 0
    MajorY = 0
    MajorZ = InsulationDia / 2
    mMRatio = (parVesselDiameter / 4 + parInsulationThickness) / _
                    (parVesselDiameter / 2 + parInsulationThickness)
    StartAngle = PI / 2
    SweepAngle = PI / 2
    normalX = 0
    normalY = -1
    normalZ = 0
    
    Dim geomFactory     As New IngrGeom3D.GeometryFactory
    Set oEqpLeftArcI = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, normalX, normalY, normalZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    axisVect.Set 1, 0, 0
    RevCenPt.Set 0, 0, 0
    Dim ObjInsLHSDome As Object
    Set ObjInsLHSDome = PlaceRevolution(m_OutputColl, oEqpLeftArcI, axisVect, RevCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsLHSDome
    Set ObjInsLHSDome = Nothing

' Insert your code for output 6(Insulation for Right Hand Side Dome)
    cenX = parVesselLength
    cenY = 0
    cenZ = 0
    MajorX = 0
    MajorY = 0
    MajorZ = InsulationDia / 2
    mMRatio = (parVesselDiameter / 4 + parInsulationThickness) / _
                    (parVesselDiameter / 2 + parInsulationThickness)
    StartAngle = PI
    SweepAngle = PI / 2
    normalX = 0
    normalY = -1
    normalZ = 0
    axisVect.Set 1, 0, 0
    RevCenPt.Set parVesselLength, 0, 0
  
    Set oEqpRightArcI = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, normalX, normalY, normalZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    
    Dim ObjInsRHSDome As Object
    Set ObjInsRHSDome = PlaceRevolution(m_OutputColl, oEqpRightArcI, axisVect, RevCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsRHSDome
    Set ObjInsRHSDome = Nothing
    Set RevCenPt = Nothing

' Insert your code for output (InsulatedBoot)
    If CmpDblGreaterthan(parBootHorLocation, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBootVertLocation, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBootDiameter, LINEAR_TOLERANCE) Then
    '       Boot cylinder
            Dim meetDistance As Double
            meetDistance = Sqr((parVesselDiameter / 2) ^ 2 - (parBootDiameter / 2) ^ 2)
            stPoint.Set parBootHorLocation, meetDistance * Sin(parBootAngularLocation), _
                                        meetDistance * Cos(parBootAngularLocation)
            enPoint.Set parBootHorLocation, parBootVertLocation * Sin(parBootAngularLocation), _
                                        parBootVertLocation * Cos(parBootAngularLocation)
            InsulationDia = parBootDiameter + 2 * parInsulationThickness
           
            Dim ObjInsBootVessel As Object
            Set ObjInsBootVessel = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)
             
             'Set the output
            m_OutputColl.AddOutput "InsBootVessel", ObjInsBootVessel
            Set ObjInsBootVessel = Nothing
        
        '       Boot Dome
            Dim ObjBootDome As IngrGeom3D.Revolution3d
            
            cenX = parBootHorLocation
            cenY = 0 'parBootVertLocation * Sin(parBootAngularLocation)
            cenZ = parBootVertLocation '* Cos(parBootAngularLocation)
            
            InsulationDia = parBootDiameter + 2 * parInsulationThickness

             MajorX = 0
             MajorY = InsulationDia / 2 ''(parBootDiameter / 2) '* Cos(parBootAngularLocation)
             MajorZ = 0 '(parBootDiameter / 2) * Sin(parBootAngularLocation)
             
             mMRatio = (parBootDiameter / 4 + parInsulationThickness) / _
                             (parBootDiameter / 2 + parInsulationThickness)
             StartAngle = PI / 2   '-0.5 * PI
             SweepAngle = PI / 2
             
             normalX = 1
             normalY = 0
             normalZ = 0
         
             Dim oBootBottomArc As IngrGeom3D.EllipticalArc3d
             Set oBootBottomArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                 cenX, cenY, cenZ, normalX, normalY, normalZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                                 StartAngle, SweepAngle)
        
             axisVect.Set 0, 0, 1
             
             Dim BootRevCenPt As New AutoMath.DPosition
                           
             BootRevCenPt.Set parBootHorLocation, 0, parBootVertLocation
                           
             Set ObjBootDome = PlaceRevolution(m_OutputColl, oBootBottomArc, axisVect, BootRevCenPt, 2 * PI, True)

          '   Apply rotation transformations to the object created.
            Dim oTransMat     As New AutoMath.DT4x4
            Dim oRotVec As New AutoMath.DVector
            oTransMat.LoadIdentity
            oRotVec.Set 1, 0, 0
                 
            oTransMat.Rotate -parBootAngularLocation, oRotVec
            ObjBootDome.Transform oTransMat

        '   Set the output
            m_OutputColl.AddOutput "BootDome", ObjBootDome
            Set ObjBootDome = Nothing
            
            Set BootRevCenPt = Nothing
            Set oBootBottomArc = Nothing
            Set oTransMat = Nothing
            Set oRotVec = Nothing

    End If
    
    Set axisVect = Nothing
    Set geomFactory = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
 
End Sub
