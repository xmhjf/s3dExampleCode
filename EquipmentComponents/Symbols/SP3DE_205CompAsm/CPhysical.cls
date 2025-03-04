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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:00 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'    This is Complex Vertical Cylindrical Equipment Skirt Component symbol.
'    Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E205 Symbol in Page no 286.
'   Symbol is created using the following Outputs:
'   i)  4 standard outputs Consisting of the following:
'       a) One Insulation aspect output,
'       b) One Physical aspect output: Vessel uses 'PlaceRevolution'
'       c) Two ReferenceGeometry aspect outputs: a Default Surface and a Control point
'   ii) Variable Outputs:
'        a) Support
'        b) Surface for the support and
'        c) Intermediate dome for shell section 3
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006     kkc       DI 95670-Replaced names with initials in the revision history.
'   25.Sep.2006     dkl       TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const MAX_SECTIONS = 4
Private PI       As Double
Private Sub Class_Initialize()
        PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    Dim ObjVessel As Object

    Dim parVesselHeight1 As Double
    Dim parVesselDiameter1 As Double
    Dim parSupportHeight1 As Double
    Dim parVesselHeight2 As Double
    Dim parVesselDiameter2 As Double
    Dim parSupportHeight2 As Double
    Dim parVesselHeight3 As Double
    Dim parVesselDiameter3 As Double
    Dim parSupportHeight3 As Double
    Dim parVesselHeight4 As Double
    Dim parVesselDiameter4 As Double
    Dim parSkirtHeight As Double
    Dim parSkirtTopDiameter As Double
    Dim parSkirtBottomDiameter As Double
    Dim parVesselReferencePoint As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselHeight1 = arrayOfInputs(2)   'P2
    parVesselDiameter1 = arrayOfInputs(3) 'P3
    parSupportHeight1 = arrayOfInputs(4)  'P4
    parVesselHeight2 = arrayOfInputs(5)   'P5
    parVesselDiameter2 = arrayOfInputs(6) 'P6
    parSupportHeight2 = arrayOfInputs(7)  'P7
    parVesselHeight3 = arrayOfInputs(8)   'P8
    parVesselDiameter3 = arrayOfInputs(9) 'P9
    parSupportHeight3 = arrayOfInputs(10) 'P10
    parVesselHeight4 = arrayOfInputs(11)  'P11
    parVesselDiameter4 = arrayOfInputs(12)      'P12
    parSkirtHeight = arrayOfInputs(13)          'P13
    parSkirtTopDiameter = arrayOfInputs(14)     'P14
    parSkirtBottomDiameter = arrayOfInputs(15)  'P15
    parVesselReferencePoint = arrayOfInputs(16) 'P16
    parInsulationThickness = arrayOfInputs(17)

    iOutput = 0

    Dim Sign As Integer
    Dim NextDiameter    As Double
    Dim TotalHeight     As Double
    Dim TopDomeMajor    As Double
    Dim BottomDomeFlag  As Integer
    Dim V3BotDomeFlag   As Integer
    Dim iCount               As Integer
    Dim jCount               As Integer

    BottomDomeFlag = 0
    V3BotDomeFlag = 0

' Insert your code for output 2(Vessel)
'   Notes:
'   - arrayOfInputs(2) to arrayOfInputs(16) refers to PDS Equipment E205 dimensions P2 to P16
'   - Array variables are used for Height,Diameter and Support Height to provide:
'       i) Compact code, ii) Number of shell sections can vary from One to Four,
'       checking for geometry conditions is easier with array index.
'   - If dimensions from P9 to P12 are not zero then three semi-elliptical domes are placed
'   - If VesselDiameter or VesselHeight is negative or zero both of them are to be taken as zero and
'       the corresponding Support height is taken as zero
'   - For some unknown reasons, Inclined lines are needing negative offset, so 'Sign' is used to
'       provide this.
    
    '' Origin is at DP/PP 1 at the base of Skirt.
    
    Sign = -1
    Dim dVesselHt(1 To MAX_SECTIONS) As Double
    Dim dVesselDia(1 To MAX_SECTIONS) As Double
    Dim dJunctionHt(1 To MAX_SECTIONS) As Double 'i.e. Support heights

    Dim NumShellSections    As Integer
    NumShellSections = 0
     For iCount = 1 To MAX_SECTIONS
        dVesselHt(iCount) = 0
        dVesselDia(iCount) = 0
        dJunctionHt(iCount) = 0
        If CmpDblGreaterthan(arrayOfInputs(3 * iCount - 1), 0) And CmpDblGreaterthan(arrayOfInputs(3 * iCount), 0) Then
            dVesselHt(iCount) = arrayOfInputs(3 * iCount - 1)
            dVesselDia(iCount) = arrayOfInputs(3 * iCount)
            If iCount < 4 Then
'               Store distance from the current shell section to the next shell section (Taper)
                dJunctionHt(iCount) = arrayOfInputs(3 * iCount + 1)
            End If
            NumShellSections = NumShellSections + 1
        End If
     Next iCount
     If NumShellSections = 0 Then GoTo ErrorLabel

'    Store distance from the Ground (Place point 1) to the shell section 4
     dJunctionHt(4) = parVesselReferencePoint - dJunctionHt(3) - dVesselHt(4)

' Vessel is formed by creating a Complex string and rotating about Z-Axis
    Dim cenX As Double
    Dim cenY As Double
    Dim cenZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim mMRatio As Double
    Dim StartAngle As Double
    Dim SweepAngle As Double
    Dim norX As Double
    Dim norY As Double
    Dim norZ As Double

    Dim oVertLine(1 To MAX_SECTIONS) As IngrGeom3D.Line3d
    Dim oInclLine(1 To MAX_SECTIONS) As IngrGeom3D.Line3d
    Dim oEqpArcTop As IngrGeom3D.EllipticalArc3d
    Dim oEqpArcMiddle As IngrGeom3D.EllipticalArc3d
    Dim oEqpArcBottom As IngrGeom3D.EllipticalArc3d
    Dim oEqpComplexStr As IngrGeom3D.ComplexString3d

'   Top Convex Dome Curve
    cenX = 0
    cenY = 0
    Dim TopDomeCenterHeight As Double
    TopDomeCenterHeight = 0
    For iCount = MAX_SECTIONS To 1 Step -1
        TopDomeCenterHeight = TopDomeCenterHeight + dJunctionHt(iCount) + dVesselHt(iCount)
    Next iCount
    cenZ = TopDomeCenterHeight
    For iCount = 1 To MAX_SECTIONS
        If CmpDblGreaterthan(dVesselDia(iCount), 0) Then
'           Use the first non-zero vessel diameter in the array
            MajorX = dVesselDia(iCount) / 2
            Exit For
        End If
    Next iCount
    TopDomeMajor = MajorX
    MajorY = 0
    MajorZ = 0
    mMRatio = 0.5
    StartAngle = PI / 2
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set oEqpArcTop = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    Dim Linepts(0 To 5) As Double
    For iCount = 1 To MAX_SECTIONS
        If CmpDblGreaterthan(dVesselHt(iCount), 0) Then         'Place shell sections which have positive height values
'           Create vertical Line along dVesselHt
'           Current shell section's Top point
            Linepts(0) = dVesselDia(iCount) / 2
            Linepts(1) = 0
            Linepts(2) = 0
            For jCount = MAX_SECTIONS To iCount Step -1
                Linepts(2) = Linepts(2) + dJunctionHt(jCount) + dVesselHt(jCount)
            Next jCount

'           Current shell section's bottom point
            Linepts(3) = Linepts(0)
            Linepts(4) = Linepts(1)
            Linepts(5) = Linepts(2) - dVesselHt(iCount)
            Set oVertLine(iCount) = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                                Linepts(3), Linepts(4), Linepts(5))
'           Inclined Line at dJunctionHt dimension
            If iCount < 4 Then       'Only three inclined lines are required for placing three Tapered sections
'               Top point of the Inclined line
                Linepts(0) = dVesselDia(iCount) / 2
                Linepts(1) = 0
                Linepts(2) = Linepts(5)

'               Bottom point of the Inclined line
                For jCount = iCount + 1 To MAX_SECTIONS
                    If CmpDblGreaterthan(dVesselDia(jCount), 0) Then    'Consider the next non-zero diameter
                        NextDiameter = dVesselDia(jCount)
                        Exit For
                    End If
                Next jCount
'               Use the next non-zero diameter for preparing the inclined line
                Linepts(3) = Linepts(0) + Sign * ((NextDiameter - dVesselDia(iCount)) / 2)
                Linepts(4) = Linepts(1)
                Linepts(5) = 0
                For jCount = MAX_SECTIONS To iCount + 1 Step -1
                    Linepts(5) = Linepts(5) + dJunctionHt(jCount) + dVesselHt(jCount)
                Next jCount

                Set oInclLine(iCount) = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                                    Linepts(3), Linepts(4), Linepts(5))
            End If
        End If
    Next iCount

'   Bottom Dome Curve
    If CmpDblGreaterthan(dVesselDia(4), 0) Then
        cenX = 0
        cenY = 0
        cenZ = parVesselReferencePoint - dJunctionHt(3) - dVesselHt(4)
        MajorX = dVesselDia(4) / 2
        MajorY = 0
        MajorZ = 0
        mMRatio = 0.5
        StartAngle = PI
        SweepAngle = PI / 2
        norX = 0
        norY = -1
        norZ = 0
        Set oEqpArcBottom = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                            StartAngle, SweepAngle)
        BottomDomeFlag = 1
    End If

'   Prepare complex string
    Dim startPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition

    Dim EleCollection           As Collection
    Set EleCollection = New Collection
    EleCollection.Add oEqpArcTop
    For iCount = 1 To MAX_SECTIONS
        If CmpDblGreaterthan(dVesselHt(iCount), 0) Then
'           Use vertical and Inclined lines which have positive shell section height values
            EleCollection.Add oVertLine(iCount)
            For jCount = MAX_SECTIONS To iCount + 1 Step -1
                If CmpDblGreaterthan(dVesselHt(jCount), 0) And iCount < MAX_SECTIONS Then  'Check if next shell section exists
'                   Place current Inclined line
                    EleCollection.Add oInclLine(iCount)
                    Exit For
                End If
            Next jCount
        End If
    Next iCount
    If CmpDblGreaterthan(dVesselDia(4), 0) Then EleCollection.Add oEqpArcBottom

'   Compute TotalHeight
    TotalHeight = 0
    For iCount = MAX_SECTIONS To 1 Step -1
        TotalHeight = TotalHeight + dJunctionHt(iCount) + dVesselHt(iCount)
    Next iCount
    TotalHeight = TotalHeight + TopDomeMajor / 2

'   Prepare complex string and revolve it about Z-Axis
    startPoint.Set 0, 0, TotalHeight
    Set oEqpComplexStr = PlaceTrCString(startPoint, EleCollection)
    Set startPoint = Nothing
    axisVect.Set 0, 0, 1
    revCenPt.Set 0, 0, parVesselReferencePoint
    Set ObjVessel = PlaceRevolution(m_OutputColl, oEqpComplexStr, axisVect, revCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVessel
    Set ObjVessel = Nothing

'   Place revolution at shell section3
    If CmpDblGreaterthan(dVesselDia(3), 0) Then
        cenX = 0
        cenY = 0
        cenZ = parVesselReferencePoint
        MajorX = dVesselDia(3) / 2
        MajorY = 0
        MajorZ = 0
        mMRatio = 0.5
        StartAngle = PI
        SweepAngle = PI / 2
        norX = 0
        norY = -1
        norZ = 0
        Dim ObjVes3Bottom As Object
        Set oEqpArcMiddle = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                            StartAngle, SweepAngle)
        Set ObjVes3Bottom = PlaceRevolution(m_OutputColl, oEqpArcMiddle, axisVect, revCenPt, 2 * PI, True)
        V3BotDomeFlag = 1

'       Set the output
        m_OutputColl.AddOutput "Ves3Bottom", ObjVes3Bottom
        Set ObjVes3Bottom = Nothing
    End If
    Set axisVect = Nothing
    Set revCenPt = Nothing

    Dim Objcurves As IJDObject
    Set Objcurves = oEqpComplexStr
    Objcurves.Remove
    Set Objcurves = Nothing

    For iCount = 1 To MAX_SECTIONS
        Set oVertLine(iCount) = Nothing
        Set oInclLine(iCount) = Nothing
    Next iCount
    Set oEqpArcTop = Nothing
    Set oEqpArcMiddle = Nothing
    Set oEqpArcBottom = Nothing

    Dim count As Integer
    For count = 1 To EleCollection.count
        EleCollection.Remove 1
    Next count
    Set EleCollection = Nothing
    Set oEqpComplexStr = Nothing

' Insert your code for output Skirt (Truncated Cone) Support
    If CmpDblGreaterthan(parSkirtHeight, 0) And CmpDblGreaterthan(parSkirtTopDiameter, 0) And CmpDblGreaterthan(parSkirtBottomDiameter, 0) Then
        Dim baseCenter As New AutoMath.DPosition
        Dim topCenter As New AutoMath.DPosition
        Dim topRadius As Double
        Dim baseRadius As Double

        topCenter.Set 0, 0, parSkirtHeight
        baseCenter.Set 0, 0, 0
        baseRadius = parSkirtBottomDiameter / 2
        topRadius = parSkirtTopDiameter / 2
        Dim ObjSupport As Object
        Set ObjSupport = PlaceCone(m_OutputColl, baseCenter, topCenter, baseRadius, topRadius, False)

'       Set the output
        m_OutputColl.AddOutput "SkirtSupport", ObjSupport
        Set ObjSupport = Nothing
        Set topCenter = Nothing
    End If

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub

