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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:02 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RUK
'   Creation Date:  Monday, March 26 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Seven Outputs:
'       Four Cylinders and Three Cones
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private Const MAX_SECTIONS = 4


Private Sub Class_Initialize()
    ''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt           As PartFacelets.IJDPart
    Dim iOutput             As Double
    Dim ObjInsulatedVessel  As Object

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
    Dim parEnd1Type As Long
    Dim parEnd2Type As Long
    Dim parEnd3Type As Long
    Dim parEndHead1ConeTopDiameter As Double
    Dim parEndHead1ConeHeight As Double
    Dim parEndHead1KnuckleRadius As Double
    Dim parEndHead1Domeradius As Double
    Dim parEndHead1FlangedThick1 As Double
    Dim parEndHead1FlangedThick2 As Double
    Dim parEndHead1SphericalRadius As Double
    Dim parEndHead2ConeTopDiameter As Double
    Dim parEndHead2ConeHeight As Double
    Dim parEndHead2KnuckleRadius As Double
    Dim parEndHead2Domeradius As Double
    Dim parEndHead2FlangedThick1 As Double
    Dim parEndHead2FlangedThick2 As Double
    Dim parEndHead2SphericalRadius As Double
    Dim parEndHead3ConeTopDiameter As Double
    Dim parEndHead3ConeHeight As Double
    Dim parEndHead3KnuckleRadius As Double
    Dim parEndHead3Domeradius As Double
    Dim parEndHead3FlangedThick1 As Double
    Dim parEndHead3FlangedThick2 As Double
    Dim parEndHead3SphericalRadius As Double
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselHeight1 = arrayOfInputs(2)             'P2
    parVesselDiameter1 = arrayOfInputs(3)           'P3
    parSupportHeight1 = arrayOfInputs(4)            'P4
    parVesselHeight2 = arrayOfInputs(5)             'P5
    parVesselDiameter2 = arrayOfInputs(6)           'P6
    parSupportHeight2 = arrayOfInputs(7)            'P7
    parVesselHeight3 = arrayOfInputs(8)             'P8
    parVesselDiameter3 = arrayOfInputs(9)           'P9
    parSupportHeight3 = arrayOfInputs(10)           'P10
    parVesselHeight4 = arrayOfInputs(11)            'P11
    parVesselDiameter4 = arrayOfInputs(12)          'P12
    parSkirtHeight = arrayOfInputs(13)              'P13
    parSkirtTopDiameter = arrayOfInputs(14)         'P14
    parSkirtBottomDiameter = arrayOfInputs(15)      'P15
    parVesselReferencePoint = arrayOfInputs(16)     'P16
    parEnd1Type = arrayOfInputs(17)
    parEnd2Type = arrayOfInputs(18)
    parEnd3Type = arrayOfInputs(19)
    parEndHead1ConeTopDiameter = arrayOfInputs(20)  'P51
    parEndHead1ConeHeight = arrayOfInputs(21)       'P50
    parEndHead1KnuckleRadius = arrayOfInputs(22)    'P51 for TORC & TORS
    parEndHead1Domeradius = arrayOfInputs(23)       'P50 for Dome
    parEndHead1FlangedThick1 = arrayOfInputs(24)    'P50 for FLGD
    parEndHead1FlangedThick2 = arrayOfInputs(25)    'P51 for FLGD
    parEndHead1SphericalRadius = arrayOfInputs(26)  'P50 for TORS
    parEndHead2ConeTopDiameter = arrayOfInputs(27)  'P51
    parEndHead2ConeHeight = arrayOfInputs(28)       'P50
    parEndHead2KnuckleRadius = arrayOfInputs(29)    'P51 for TORC & TORS
    parEndHead2Domeradius = arrayOfInputs(30)       'P50 for Dome
    parEndHead2FlangedThick1 = arrayOfInputs(31)    'P50 for FLGD
    parEndHead2FlangedThick2 = arrayOfInputs(32)    'P51 for FLGD
    parEndHead2SphericalRadius = arrayOfInputs(33)  'P50 for TORS
    parEndHead3ConeTopDiameter = arrayOfInputs(34)  'P51
    parEndHead3ConeHeight = arrayOfInputs(35)       'P50
    parEndHead3KnuckleRadius = arrayOfInputs(36)    'P51 for TORC & TORS
    parEndHead3Domeradius = arrayOfInputs(37)       'P50 for Dome
    parEndHead3FlangedThick1 = arrayOfInputs(38)    'P50 for FLGD
    parEndHead3FlangedThick2 = arrayOfInputs(39)    'P51 for FLGD
    parEndHead3SphericalRadius = arrayOfInputs(40)  'P50 for TORS
    parInsulationThickness = arrayOfInputs(41)

    'Origin is at DP/PP 2 at the top of Skirt.
    iOutput = 0
    'Inset your code for Insulated Vessel
        'Start from DP3
            '--Cylider, Cone, Cylinder, Cone, Cylinder are Upper parts
            '--Cone, Cylinder are lower parts
    
    Dim NumShellSections    As Integer
    Dim dVesselHt(1 To MAX_SECTIONS) As Double
    Dim dVesseldia(1 To MAX_SECTIONS) As Double
    Dim dJunctionHt(1 To MAX_SECTIONS) As Double
    Dim iCount As Integer
    NumShellSections = 0
    For iCount = 1 To MAX_SECTIONS
        dVesselHt(iCount) = 0
        dVesseldia(iCount) = 0
        dJunctionHt(iCount) = 0
        If CmpDblGreaterthan(arrayOfInputs(3 * iCount - 1), LINEAR_TOLERANCE) And CmpDblGreaterthan(arrayOfInputs(3 * iCount), LINEAR_TOLERANCE) Then
            dVesselHt(iCount) = arrayOfInputs(3 * iCount - 1)
            dVesseldia(iCount) = arrayOfInputs(3 * iCount) + 2 * parInsulationThickness
            If iCount < 4 Then
                'Store distance from the current shell section to the next shell section (Taper)
                dJunctionHt(iCount) = arrayOfInputs(3 * iCount + 1)
            End If
            NumShellSections = NumShellSections + 1
        End If
     Next iCount
     If NumShellSections = 0 Then GoTo ErrorLabel

    'Store distance from the Place point to the shell section 4
    dJunctionHt(4) = parVesselReferencePoint - dJunctionHt(3) - dVesselHt(4)
    
    Dim ObjInsCylinder As Object
    Dim ObjInsCone As Object
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim SupportBottomDia As Double
    
    'First Cylinder
    If CmpDblGreaterthan(parVesselDiameter1, LINEAR_TOLERANCE) And CmpDblGreaterthan(parVesselHeight1, LINEAR_TOLERANCE) Then
        stPoint.Set 0, 0, parVesselReferencePoint + dVesselHt(3) + _
                            dJunctionHt(2) + dVesselHt(2) + dJunctionHt(1)
        enPoint.Set stPoint.x, stPoint.y, stPoint.z + dVesselHt(1)
        Set ObjInsCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, dVesseldia(1), True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsCylinder
        
    'First Support
        If CmpDblGreaterthan(parSupportHeight1, LINEAR_TOLERANCE) Then
            stPoint.Set 0, 0, parVesselReferencePoint + dVesselHt(3) + _
                                dJunctionHt(2) + dVesselHt(2)
            enPoint.Set stPoint.x, stPoint.y, stPoint.z + parSupportHeight1
            For iCount = 2 To MAX_SECTIONS
                If CmpDblGreaterthan(dVesseldia(iCount), LINEAR_TOLERANCE) Then
                    'Use the first non-zero vessel diameter in the array
                    SupportBottomDia = dVesseldia(iCount)
                Exit For
                End If
            Next iCount
            If CmpDblGreaterthan(SupportBottomDia, LINEAR_TOLERANCE) Then
                Set ObjInsCone = PlaceCone(m_OutputColl, stPoint, enPoint, SupportBottomDia / 2, dVesseldia(1) / 2, True)
                'Set the Output
                iOutput = iOutput + 1
                m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsCone
            End If
        End If
    End If
    'Second Cylinder
    If CmpDblGreaterthan(parVesselDiameter2, LINEAR_TOLERANCE) And CmpDblGreaterthan(parVesselHeight2, LINEAR_TOLERANCE) Then
        stPoint.Set 0, 0, parVesselReferencePoint + dVesselHt(3) + dJunctionHt(2)
        enPoint.Set stPoint.x, stPoint.y, stPoint.z + dVesselHt(2)
        Set ObjInsCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, dVesseldia(2), True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsCylinder
    
    'Second Support
        If CmpDblGreaterthan(parSupportHeight2, LINEAR_TOLERANCE) Then
            stPoint.Set 0, 0, parVesselReferencePoint + dVesselHt(3)
            enPoint.Set stPoint.x, stPoint.y, stPoint.z + parSupportHeight2
            SupportBottomDia = 0
            For iCount = 3 To MAX_SECTIONS
                If CmpDblGreaterthan(dVesseldia(iCount), LINEAR_TOLERANCE) Then
                    'Use the first non-zero vessel diameter in the array
                    SupportBottomDia = dVesseldia(iCount)
                Exit For
                End If
            Next iCount
            If CmpDblGreaterthan(SupportBottomDia, LINEAR_TOLERANCE) Then
                Set ObjInsCone = PlaceCone(m_OutputColl, stPoint, enPoint, SupportBottomDia / 2, dVesseldia(2) / 2, True)
                'Set the output
                iOutput = iOutput + 1
                m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsCone
            End If
        End If
    End If
    
    'Third Cylinder
    If CmpDblGreaterthan(parVesselDiameter3, LINEAR_TOLERANCE) And CmpDblGreaterthan(parVesselHeight3, LINEAR_TOLERANCE) Then
        stPoint.Set 0, 0, parVesselReferencePoint
        enPoint.Set 0, 0, parVesselReferencePoint + dVesselHt(3)
        Set ObjInsCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, dVesseldia(3), True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsCylinder
    
    'Third Support
        If CmpDblGreaterthan(parSupportHeight3, LINEAR_TOLERANCE) Then
            stPoint.Set 0, 0, parVesselReferencePoint
            enPoint.Set 0, 0, parVesselReferencePoint - parSupportHeight3
            SupportBottomDia = 0
            For iCount = 4 To MAX_SECTIONS
                If CmpDblGreaterthan(dVesseldia(iCount), LINEAR_TOLERANCE) Then
                    'Use the first non-zero vessel diameter in the array
                    SupportBottomDia = dVesseldia(iCount)
                Exit For
                End If
            Next iCount
            If CmpDblGreaterthan(SupportBottomDia, LINEAR_TOLERANCE) Then
                Set ObjInsCone = PlaceCone(m_OutputColl, stPoint, enPoint, dVesseldia(3) / 2, SupportBottomDia / 2, True)
                'Set the Output
                iOutput = iOutput + 1
                m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsCone
            End If
        End If
    End If
    
    'Fourth Cylinder
    If CmpDblGreaterthan(parVesselDiameter4, LINEAR_TOLERANCE) And CmpDblGreaterthan(parVesselHeight4, LINEAR_TOLERANCE) Then
        stPoint.Set 0, 0, parVesselReferencePoint - dJunctionHt(3)
        enPoint.Set 0, 0, stPoint.z - dVesselHt(4)
        Set ObjInsCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, dVesseldia(4), True)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsCylinder
    End If
    Set ObjInsCylinder = Nothing
    Set ObjInsCone = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
