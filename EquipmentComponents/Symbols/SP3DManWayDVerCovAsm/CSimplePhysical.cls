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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:10:52 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:   Tuesday, April 222003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This symbol is based on "Davit for Vertical Cover" symbol
' of The DOW Chemical company.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const POS_RIGHT = 3
Private PI       As Double

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oGeomHelper = New SymbolServices
    PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
    
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    Dim ObjHexNut As Object
    Dim ObjDavitArmCurved As Object
    Dim ObjHingePlate As Object

    Dim parCentertoDavitCover As Double
    Dim parVesselDiameter As Double
    Dim parDavitHinge    As Long
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCentertoDavitCover = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parDavitHinge = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    m_oGeomHelper.OutputCollection = m_OutputColl

    Dim NozzleLength As Double
    NozzleLength = parCentertoDavitCover - parVesselDiameter / 2

' Place Nozzle 1
    Dim oNozzle             As IJDNozzle
    Dim oDir As New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition
'   Compute nozzle length and set the nozzle length property
    Dim IntrusionDist As Double 'Intrusion Distance of Manhole Body into tower

    IntrusionDist = parVesselDiameter / 2 - Sqr(Abs((parVesselDiameter / 2) ^ 2 - (pipeDiam / 2) ^ 2)) _
                    + 0.01   '10mm intrusion in addition to the minimum needed
    'If the nozzle length is than 0.1, set the nozzle length equal to flange thickness
    If (NozzleLength + IntrusionDist) < 0.1 Then
        NozzleLength = flangeThick
    Else
        NozzleLength = (NozzleLength + IntrusionDist)
    End If
    

    oDir.Set 0, 1, 0
    oPlacePoint.Set 0, parCentertoDavitCover, 0
    
    
    Set oNozzle = CreateNozzlePHWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

' Insert your code for output 2(Manhole Cover)
    Dim Gap As Double  'Gap between Flange and Manhole Cover
    Gap = 0.004

    stPoint.Set 0, parCentertoDavitCover + Gap, 0
    enPoint.Set stPoint.x, stPoint.y + flangeThick, stPoint.z

    'Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, _
                                    flangeDiam

' Insert your code for output 3(EyeBolt)
    Dim EyeBoltDia As Double
    Dim EyeBoltLength As Double
    Dim HexNutHeight As Double
    Dim HexNutDia   As Double
    Dim NutTopPoint As New AutoMath.DPosition

    EyeBoltDia = 0.02
    HexNutDia = Sqr(3) * EyeBoltDia     '1.732 times the bolt diameter
    HexNutHeight = 0.86 * EyeBoltDia
    EyeBoltLength = 0.15 + HexNutHeight + 0.005 'Considered 5mm projection above the HexNut.

    stPoint.Set 0, parCentertoDavitCover + Gap + flangeThick / 2, flangeDiam / 2
    enPoint.Set stPoint.x, stPoint.y, stPoint.z + EyeBoltLength
    NutTopPoint.Set stPoint.x, stPoint.y, stPoint.z + EyeBoltLength - 0.005 'Considered 5mm projection above the HexNut.

    'Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, EyeBoltDia

' Insert your code for output 4(Hexagonal Nut for EyeBolt)
    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim angle As Double

    angle = PI / 3

'   First point is considered on positive Z-axis at an angle (PI / 6 + parAngle) from positive Y-axis.
    Dim iCount As Integer
    For iCount = 1 To 7
        LineStrPoints(3 * iCount - 3) = NutTopPoint.x + (HexNutDia / 2) * Cos(PI / 6 + iCount * angle)
        LineStrPoints(3 * iCount - 2) = NutTopPoint.y + (HexNutDia / 2) * Sin(PI / 6 + iCount * angle)
        LineStrPoints(3 * iCount - 1) = NutTopPoint.z
    Next iCount

    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, LineStrPoints)
    Dim axisvect As New AutoMath.DVector

    axisvect.Set 0, 0, -1
    Set ObjHexNut = PlaceProjection(m_OutputColl, oLineString, axisvect, HexNutHeight, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHexNut
    Set ObjHexNut = Nothing
    Set oLineString = Nothing
    Set NutTopPoint = Nothing

' Insert your code for output 5(Davit Arm Top)
'   Decide DavitRodCSDia based on the Manhole pipeOD and pressure rating
    Dim DavitRodCSDia As Double
    DavitRodCSDia = 0.04


'   Assumption: Extension of 'Davit Arm Top' from eyebolt axis is taken to be equal to HexNutDia
    Dim ArmTopExtn As Double
    Dim ArmTopLength As Double

    ArmTopExtn = HexNutDia
    ArmTopLength = Abs((Sqr((flangeDiam / 2 + 0.05) ^ 2 + _
                    (flangeThick + Gap) ^ 2) + _
                    ArmTopExtn) - 0.25)
'   Compute angle of Arm Top with the x-axis.
    Dim ArmTopAng As Double
    Dim centerPoint As New AutoMath.DPosition

    ArmTopAng = Atn((flangeThick + Gap) / (flangeDiam / 2 + 0.05))

    stPoint.Set -ArmTopExtn * Cos(ArmTopAng), ArmTopExtn * Sin(ArmTopAng) + _
            parCentertoDavitCover + Gap + flangeThick / 2, _
            flangeDiam / 2 + 0.15

    enPoint.Set (ArmTopLength - ArmTopExtn) * Cos(ArmTopAng), _
            -(ArmTopLength - ArmTopExtn) * Sin(ArmTopAng) + _
            parCentertoDavitCover + Gap + flangeThick / 2, _
            flangeDiam / 2 + 0.15

    If parDavitHinge = POS_RIGHT Then
        stPoint.x = -stPoint.x
        enPoint.x = -enPoint.x
    End If
    centerPoint.Set enPoint.x, enPoint.y, enPoint.z

    'Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, DavitRodCSDia

' Insert your code for output 6(Davit Arm Curved)
    Dim objCircle       As IngrGeom3D.Circle3d
    Dim Norm As New AutoMath.DVector

    Norm.Set Cos(ArmTopAng), -Sin(ArmTopAng), 0
    axisvect.Set Sin(ArmTopAng), Cos(ArmTopAng), 0

    If parDavitHinge = POS_RIGHT Then
        Norm.Set -Cos(ArmTopAng), Sin(ArmTopAng), 0
        axisvect.Set Sin(ArmTopAng), -Cos(ArmTopAng), 0
    End If

    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            centerPoint.x, centerPoint.y, centerPoint.z, _
                            Norm.x, Norm.y, Norm.z, _
                            DavitRodCSDia / 2)

'   Set revolution center point
    Dim RevCenPoint As New AutoMath.DPosition
    RevCenPoint.Set centerPoint.x, centerPoint.y, _
            centerPoint.z - 0.25

    Set ObjDavitArmCurved = PlaceRevolution(m_OutputColl, objCircle, _
                axisvect, RevCenPoint, PI / 2, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDavitArmCurved
    Set ObjDavitArmCurved = Nothing
    Set centerPoint = Nothing
    Set RevCenPoint = Nothing

' Insert your code for output 7(Davit Arm Collar)
    stPoint.Set flangeDiam / 2 + 0.05, _
                parCentertoDavitCover - flangeThick / 2, _
                0.15
    enPoint.Set stPoint.x, stPoint.y, stPoint.z + 0.025
    If parDavitHinge = POS_RIGHT Then
        stPoint.x = -stPoint.x
        enPoint.x = -enPoint.x
    End If

    'Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, _
                                    (DavitRodCSDia + 0.025)

' Insert your code for output 8(Davit Arm Bottom)
    stPoint.Set flangeDiam / 2 + 0.05, _
                parCentertoDavitCover - flangeThick / 2, _
                flangeDiam / 2 - 0.1

'   Assumption: Davit Arm Bottom extends below the hinge plate by DavitRodCSDia
    enPoint.Set stPoint.x, stPoint.y, -DavitRodCSDia

    If parDavitHinge = POS_RIGHT Then
        stPoint.x = -stPoint.x
        enPoint.x = -enPoint.x
    End If

    'Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, DavitRodCSDia
    Set stPoint = Nothing
    Set enPoint = Nothing

' Insert your code for output 9(Hinge Plate)
    Dim BktTopMeetDist As Double

    BktTopMeetDist = Sqr(Abs((flangeDiam / 2) ^ 2 - 0.15 ^ 2))
    LineStrPoints(0) = BktTopMeetDist
    LineStrPoints(1) = parCentertoDavitCover - flangeThick / 2 _
                        - (DavitRodCSDia + 0.05) / 2
    LineStrPoints(2) = 0

    LineStrPoints(3) = LineStrPoints(0)
    LineStrPoints(4) = LineStrPoints(1)
    LineStrPoints(5) = 0.15

'   As per the skecth Hinge plate thickness is 0.015, if we consider its bend radius
'   is 0.015, then LineStrPoints(6) is as follows:
    Dim bendRadius As Double
    bendRadius = 0.015
    LineStrPoints(6) = flangeDiam / 2 + 0.05 + _
                        (DavitRodCSDia + 0.025) / 2 + bendRadius + 0.015 / 2
    LineStrPoints(7) = LineStrPoints(1)
    LineStrPoints(8) = LineStrPoints(5)

    LineStrPoints(9) = LineStrPoints(6)
    LineStrPoints(10) = LineStrPoints(1)
    LineStrPoints(11) = 0

    LineStrPoints(12) = LineStrPoints(0)
    LineStrPoints(13) = LineStrPoints(1)
    LineStrPoints(14) = LineStrPoints(2)

    If parDavitHinge = POS_RIGHT Then
        For iCount = 1 To 5
            LineStrPoints(3 * iCount - 3) = -LineStrPoints(3 * iCount - 3)
        Next iCount
    End If

    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPoints)
    axisvect.Set 0, 1, 0
    Set ObjHingePlate = PlaceProjection(m_OutputColl, oLineString, axisvect, DavitRodCSDia + 0.05, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHingePlate
    Set ObjHingePlate = Nothing
    Set geomFactory = Nothing
    Set axisvect = Nothing

    Exit Sub

ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
