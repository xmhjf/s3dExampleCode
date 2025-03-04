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
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         VRK
'   Creation Date:  Friday,March 9 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Three Outputs:
'    Insulated Vessel, Insulated BootVessel,and Insulated BootDome
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    Dim ObjInsulatedVessel As Object
    
    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parVesselTantoTan As Double
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
    Dim parEndHeadConeTopDiameter As Double
    Dim parEndHeadConeHeight As Double
    Dim parEndHeadKnuckleRadius As Double
    Dim parEndHeadDomeradius As Double
    Dim parEndHeadFlangedThick1 As Double
    Dim parEndHeadFlangedThick2 As Double
    Dim parEndHeadSphericalRadius As Double
    Dim parInsulationThickness As Double
    Dim parXBoltHole As Double
    Dim parYBoltHole As Double
         
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim InsulationDia As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parEnd1Type = arrayOfInputs(2)
    parEnd2Type = arrayOfInputs(3)
    parVesselTantoTan = arrayOfInputs(4)            'P2
    parVesselDiameter = arrayOfInputs(5)            'P3
    parBootHorLocation = arrayOfInputs(6)           'P4
    parBootVertLocation = arrayOfInputs(7)          'P5
    parBootDiameter = arrayOfInputs(8)              'P6
    parBootAngularLocation = arrayOfInputs(9)       'P7
    parVesselCenterHeight = arrayOfInputs(10)       'P8
    parFirstSupportLocation = arrayOfInputs(11)     'P9
    parSecondSupportLocation = arrayOfInputs(12)    'P10
    parThirdSupportLocation = arrayOfInputs(13)     'P11
    parSupportThickness = arrayOfInputs(14)         'P12
    parSupportLength = arrayOfInputs(15)            'P13
    parStiffenerRadius = arrayOfInputs(16)          'P14
    parEndHeadConeTopDiameter = arrayOfInputs(17)   'P51
    parEndHeadConeHeight = arrayOfInputs(18)        'P50
    parEndHeadKnuckleRadius = arrayOfInputs(19)     'P51 for TORC n TORS
    parEndHeadDomeradius = arrayOfInputs(20)        'P50
    parEndHeadFlangedThick1 = arrayOfInputs(21)     'P50 for FLGD
    parEndHeadFlangedThick2 = arrayOfInputs(22)     'P51
    parEndHeadSphericalRadius = arrayOfInputs(23)   'P50
    parInsulationThickness = arrayOfInputs(24)
    parXBoltHole = arrayOfInputs(25)
    parYBoltHole = arrayOfInputs(26)
  
    iOutput = 0

    'Insert your code for output InsulatedVessel
    stPoint.Set 0, 0, 0
    enPoint.Set parVesselTantoTan, 0, 0
    InsulationDia = parVesselDiameter + 2 * parInsulationThickness
    Set ObjInsulatedVessel = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDia, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedVessel
    Set ObjInsulatedVessel = Nothing

    'Insert your code for output 5(Insulation for Left Hand Side Dome)
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

    'Insert your code for output (InsulatedBoot)
    If CmpDblGreaterthan(parBootHorLocation, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBootVertLocation, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBootDiameter, LINEAR_TOLERANCE) Then
            'Boot cylinder
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
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "InsBootVessel", ObjInsBootVessel
            Set ObjInsBootVessel = Nothing

            'Boot Dome
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
             Set oBootBottomArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                 cenX, cenY, cenZ, normalX, normalY, normalZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                                 StartAngle, SweepAngle)

             axisVect.Set 0, 0, 1
             Dim BootRevCenPt As New AutoMath.DPosition
             BootRevCenPt.Set parBootHorLocation, 0, parBootVertLocation
             Set ObjBootDome = PlaceRevolution(m_OutputColl, oBootBottomArc, axisVect, BootRevCenPt, 2 * PI, True)
            'Apply rotation transformations to the object created.
            Dim oTransMat     As New AutoMath.DT4x4
            Dim oRotVec As New AutoMath.DVector
            oTransMat.LoadIdentity
            oRotVec.Set 1, 0, 0
            oTransMat.Rotate -parBootAngularLocation, oRotVec
            ObjBootDome.Transform oTransMat
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "InsBootDome", ObjBootDome
            Set ObjBootDome = Nothing
            Set BootRevCenPt = Nothing
            Set oBootBottomArc = Nothing
            Set oTransMat = Nothing
            Set oRotVec = Nothing
    End If
    
    Set axisVect = Nothing
    Set oGeomFactory = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub


