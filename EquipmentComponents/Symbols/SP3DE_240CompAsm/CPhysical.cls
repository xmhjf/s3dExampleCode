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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:51 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CPhysical.cls
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

Private Const MODULE = "Physical:" 'Used for error messages
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

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
' Insert your code for output 1(Vessel)
    stPoint.Set 0, 0, 0
    enPoint.Set parVesselLength, 0, 0
    Set ObjVessel = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter, False)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVessel
    Set ObjVessel = Nothing

' Insert your code for output 2(Left Hand Side Dome)
    Dim oAxisVect As New AutoMath.DVector
    Dim oRevCenPt As New AutoMath.DPosition
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
    
    Dim oEqpLeftArc As IngrGeom3D.EllipticalArc3d
    Dim oEqpRightArc As IngrGeom3D.EllipticalArc3d
    
    cenX = 0
    cenY = 0
    cenZ = 0
    MajorX = 0
    MajorY = 0
    MajorZ = parVesselDiameter / 2
    mMRatio = 0.5
    StartAngle = PI / 2
    SweepAngle = PI / 2
    normalX = 0
    normalY = -1
    normalZ = 0
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Set oEqpLeftArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, normalX, normalY, normalZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
                                                        
    oAxisVect.Set 1, 0, 0
    oRevCenPt.Set 0, 0, 0
    
    Dim ObjLHSDome As Object
    Set ObjLHSDome = PlaceRevolution(m_OutputColl, oEqpLeftArc, oAxisVect, oRevCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLHSDome
    Set ObjLHSDome = Nothing
    Set oEqpLeftArc = Nothing

' Insert your code for output 3(Right Hand Side Dome)
    cenX = parVesselLength
    cenY = 0
    cenZ = 0
    MajorX = 0
    MajorY = 0
    MajorZ = parVesselDiameter / 2
    mMRatio = 0.5
    StartAngle = PI
    SweepAngle = PI / 2
    normalX = 0
    normalY = -1
    normalZ = 0

    Set oEqpRightArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, normalX, normalY, normalZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    oRevCenPt.Set parVesselLength, 0, 0
    
    Dim ObjRHSDome As Object
    Set ObjRHSDome = PlaceRevolution(m_OutputColl, oEqpRightArc, oAxisVect, oRevCenPt, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSDome
    Set ObjRHSDome = Nothing
    Set oEqpRightArc = Nothing
                                                            
' Insert your code for output (Boot)
    If CmpDblGreaterthan(parBootHorLocation, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBootVertLocation, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBootDiameter, LINEAR_TOLERANCE) Then
        '       Boot cylinder
                Dim ObjBootVessel As Object
                Dim meetDistance As Double
                
                meetDistance = Sqr((parVesselDiameter / 2) ^ 2 - (parBootDiameter / 2) ^ 2)
                stPoint.Set parBootHorLocation, meetDistance * Sin(parBootAngularLocation), _
                                            meetDistance * Cos(parBootAngularLocation)
                enPoint.Set parBootHorLocation, parBootVertLocation * Sin(parBootAngularLocation), _
                                            parBootVertLocation * Cos(parBootAngularLocation)
                Set ObjBootVessel = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBootDiameter, True)
                
                'Set the output
                m_OutputColl.AddOutput "BootVessel", ObjBootVessel
                Set ObjBootVessel = Nothing
        
        '       Boot Dome
                Dim ObjBootDome As IngrGeom3D.Revolution3d
                
                cenX = parBootHorLocation
                cenY = 0
                cenZ = parBootVertLocation
                
                MajorX = 0
                MajorY = (parBootDiameter / 2)
                MajorZ = 0
                
                mMRatio = 0.5
                StartAngle = PI / 2   '-0.5 * PI
                SweepAngle = PI / 2

                normalX = 1
                normalY = 0
                normalZ = 0
            
                Dim oBootBottomArc As IngrGeom3D.EllipticalArc3d
                Set oBootBottomArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                                    cenX, cenY, cenZ, normalX, normalY, normalZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                                    StartAngle, SweepAngle)
                
                oAxisVect.Set 0, 0, 1
                
                Dim BootRevCenPt As New AutoMath.DPosition
                              
                BootRevCenPt.Set parBootHorLocation, 0, parBootVertLocation
                              
                Set ObjBootDome = PlaceRevolution(m_OutputColl, oBootBottomArc, oAxisVect, BootRevCenPt, 2 * PI, True)

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

'   Placement of supports
    If CmpDblGreaterthan(parVesselCenterHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) Then
                If CmpDblGreaterthan(parFirstSupportLocation, LINEAR_TOLERANCE) Then
                    ' Insert your code for output 3(Support1)
                    Dim ObjSupport1 As Object
                    stPoint.Set parFirstSupportLocation - parSupportThickness / 2, _
                                 -parSupportLength / 2, _
                                 -parVesselCenterHeight
                    enPoint.Set stPoint.x + parSupportThickness, _
                                stPoint.y + parSupportLength, _
                                stPoint.z + parVesselCenterHeight
                    Set ObjSupport1 = PlaceBox(m_OutputColl, stPoint, enPoint)
                
                '   Set the output
                    m_OutputColl.AddOutput "Support_", ObjSupport1
                    Set ObjSupport1 = Nothing
                End If
                
                If CmpDblGreaterthan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
                    ' Insert your code for output 4(Support2)
                      Dim ObjSupport2 As Object
                    stPoint.Set parFirstSupportLocation + parSecondSupportLocation - parSupportThickness / 2, _
                                 -parSupportLength / 2, _
                                 -parVesselCenterHeight
                    enPoint.Set stPoint.x + parSupportThickness, _
                                stPoint.y + parSupportLength, _
                                stPoint.z + parVesselCenterHeight
                    Set ObjSupport2 = PlaceBox(m_OutputColl, stPoint, enPoint)
                
                '   Set the output
                    m_OutputColl.AddOutput "Support_", ObjSupport2
                    Set ObjSupport2 = Nothing
                End If
        
                If CmpDblGreaterthan(parThirdSupportLocation, LINEAR_TOLERANCE) Then
                    ' Insert your code for output 5(Support3)
                    Dim ObjSupport3 As Object
                    stPoint.Set parFirstSupportLocation + parSecondSupportLocation + parThirdSupportLocation - parSupportThickness / 2, _
                                 -parSupportLength / 2, _
                                 -parVesselCenterHeight
                    enPoint.Set stPoint.x + parSupportThickness, _
                                stPoint.y + parSupportLength, _
                                stPoint.z + parVesselCenterHeight
                    Set ObjSupport3 = PlaceBox(m_OutputColl, stPoint, enPoint)
                    
                '   Set the output
                    m_OutputColl.AddOutput "Support_", ObjSupport3
                    Set ObjSupport3 = Nothing
                End If
    End If
    
'   Placement of stiffeners
    If CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) And CmpDblGreaterthan(parStiffenerRadius, LINEAR_TOLERANCE) Then
        If CmpDblGreaterthan(parFirstSupportLocation, LINEAR_TOLERANCE) Then
'           Insert your code for output 6(Stiffner1)
            Dim ObjStiffner1 As Object
            Dim StiffnerPoints(0 To 14)  As Double
            
'           Consider a section of Stiffener in XZ plane
'           Left hand side Bottom point
            StiffnerPoints(0) = parFirstSupportLocation - parSupportThickness / 2
            StiffnerPoints(1) = 0
            StiffnerPoints(2) = parVesselDiameter / 2
        
'           Right hand side Bottom point
            StiffnerPoints(3) = parFirstSupportLocation + parSupportThickness / 2
            StiffnerPoints(4) = 0
            StiffnerPoints(5) = parVesselDiameter / 2
            
'           Right hand side Top point
            StiffnerPoints(6) = parFirstSupportLocation + parSupportThickness / 2
            StiffnerPoints(7) = 0
            StiffnerPoints(8) = parStiffenerRadius
                
'           Left hand side Top point
            StiffnerPoints(9) = parFirstSupportLocation - parSupportThickness / 2
            StiffnerPoints(10) = 0
            StiffnerPoints(11) = parStiffenerRadius
            
'           Left hand side Bottom point
            StiffnerPoints(12) = StiffnerPoints(0)
            StiffnerPoints(13) = StiffnerPoints(1)
            StiffnerPoints(14) = StiffnerPoints(2)
                
        '   Prepare profile
            oAxisVect.Set 1, 0, 0
            oRevCenPt.Set parFirstSupportLocation, 0, 0
            Dim oLineString As IngrGeom3D.LineString3d
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, StiffnerPoints)
            Set ObjStiffner1 = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oRevCenPt, 2 * PI, True)
        
        '   Set the output
            m_OutputColl.AddOutput "Stiffner_", ObjStiffner1
            Set ObjStiffner1 = Nothing
            Set oLineString = Nothing
        End If
        
        If CmpDblGreaterthan(parSecondSupportLocation, LINEAR_TOLERANCE) Then
'           Insert your code for output 7(Stiffner2)
            Dim ObjStiffner2 As Object
'           Left hand side Bottom point
            StiffnerPoints(0) = parFirstSupportLocation + parSecondSupportLocation - _
                                    parSupportThickness / 2
            StiffnerPoints(1) = 0
            StiffnerPoints(2) = parVesselDiameter / 2
        
        '   Right hand side Bottom point
            StiffnerPoints(3) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parSupportThickness / 2
            StiffnerPoints(4) = 0
            StiffnerPoints(5) = parVesselDiameter / 2
            
        '   Right hand side Top point
            StiffnerPoints(6) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parSupportThickness / 2
            StiffnerPoints(7) = 0
            StiffnerPoints(8) = parStiffenerRadius
                
        '   Left hand side Top point
            StiffnerPoints(9) = parFirstSupportLocation + parSecondSupportLocation - _
                                    parSupportThickness / 2
            StiffnerPoints(10) = 0
            StiffnerPoints(11) = parStiffenerRadius
            
        '   Left hand side Bottom point
            StiffnerPoints(12) = StiffnerPoints(0)
            StiffnerPoints(13) = StiffnerPoints(1)
            StiffnerPoints(14) = StiffnerPoints(2)
        
            Set oLineString = New IngrGeom3D.LineString3d
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, StiffnerPoints)
            oAxisVect.Set 1, 0, 0
            oRevCenPt.Set parFirstSupportLocation + parSecondSupportLocation, 0, 0
            Set ObjStiffner2 = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oRevCenPt, 2 * PI, True)
            
        '   Set the output
            m_OutputColl.AddOutput "Stiffner_", ObjStiffner2
            Set ObjStiffner2 = Nothing
        
        '   Remove  linestring
            Dim objLineString As IJDObject
            Set objLineString = oLineString
            Set oLineString = Nothing
            objLineString.Remove
    End If
    
    If CmpDblGreaterthan(parThirdSupportLocation, LINEAR_TOLERANCE) Then
        
        ' Insert your code for output 8(Stiffner3)
            Dim ObjStiffner3 As Object
        '   Left hand side Bottom point
            StiffnerPoints(0) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parThirdSupportLocation - parSupportThickness / 2
            StiffnerPoints(1) = 0
            StiffnerPoints(2) = parVesselDiameter / 2
        
        '   Right hand side Bottom point
            StiffnerPoints(3) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parThirdSupportLocation + parSupportThickness / 2
            StiffnerPoints(4) = 0
            StiffnerPoints(5) = parVesselDiameter / 2
            
        '   Right hand side Top point
            StiffnerPoints(6) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parThirdSupportLocation + parSupportThickness / 2
            StiffnerPoints(7) = 0
            StiffnerPoints(8) = parStiffenerRadius
                
        '   Left hand side Top point
            StiffnerPoints(9) = parFirstSupportLocation + parSecondSupportLocation + _
                                    parThirdSupportLocation - parSupportThickness / 2
            StiffnerPoints(10) = 0
            StiffnerPoints(11) = parStiffenerRadius
            
        '   Left hand side Bottom point
            StiffnerPoints(12) = StiffnerPoints(0)
            StiffnerPoints(13) = StiffnerPoints(1)
            StiffnerPoints(14) = StiffnerPoints(2)
        
            Set oLineString = New IngrGeom3D.LineString3d
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, StiffnerPoints)
            oAxisVect.Set 1, 0, 0
            oRevCenPt.Set parFirstSupportLocation + parSecondSupportLocation + parThirdSupportLocation, 0, 0
            Set ObjStiffner3 = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oRevCenPt, 2 * PI, True)
            
        '   Set the output
            m_OutputColl.AddOutput "Stiffner_", ObjStiffner3
            Set ObjStiffner3 = Nothing
            
        '   Remove  linestring
            Set objLineString = oLineString
            Set oLineString = Nothing
            objLineString.Remove
            Set objLineString = Nothing
        End If
    End If
    Set oAxisVect = Nothing
    Set oRevCenPt = Nothing

    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
       
End Sub

