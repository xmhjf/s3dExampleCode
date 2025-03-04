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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:09:46 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Friday, May 9, 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   This is a Vessel Saddle Support for 1.2mts to 3 mts Dia Vessels.
'   This symbol consists of 13 inputs and 12 outputs  and It has two aspects one is Simple Physical
'   another is ReferenceGeometry. Among them 9 are physical and 3 are ReferenceGeometry.
'   This is a Dow Equipment symbol
'
'   Change History:
'   dd.mmm.yyyy      who          change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav       CR-89878    Removed reference to Dow Emetl Standards (replaced existing symbol).
'   17.Dec.2007        PK          TR-117519   Modified code so that no bad representation is shown on modification of vessel diameter
'   20.FEB.2008        PK          TR-117519   Code has been modified so that extra supports are created only when
'                                              stiffener spacing is less than or equal to half of the support length.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double

Dim m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
    PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim parVesselDiameter As Double
    Dim parSupportLength As Double
    Dim parSupportHeight As Double
    Dim parSupportWidth As Double
    Dim parCPtoSaddle As Double
    Dim parBasePlateThickness As Double
    Dim parBackingPlateThickness As Double
    Dim parBackingPlateWidth As Double
    Dim parStiffnerThickness As Double
    Dim parStiffnerThickness1 As Double
    Dim parStiffnerSpacing As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double
    
    Dim iOutput     As Double
    Dim ObjSaddlePlate As Object
    Dim ObjStiffnerPlate As Object
    Dim ObjBackingPlate1 As Object
    Dim ObjBackingPlate2  As Object
    Dim ObjBackingPlate3 As Object
    Dim ObjBackingPlate4 As Object
    Dim ObjBackingPlate5 As Object
    Dim ObjBackingPlate6 As Object
    Dim ObjFoundationPort As IJEqpFoundationPort
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parSupportLength = arrayOfInputs(3)
    parSupportHeight = arrayOfInputs(4)
    parSupportWidth = arrayOfInputs(5)
    parCPtoSaddle = arrayOfInputs(6)
    parBasePlateThickness = arrayOfInputs(7)
    parBackingPlateThickness = arrayOfInputs(8)
    parBackingPlateWidth = arrayOfInputs(9)
    parStiffnerThickness = arrayOfInputs(10)
    parStiffnerThickness1 = arrayOfInputs(11)
    parStiffnerSpacing = arrayOfInputs(12)
    parXboltHole = arrayOfInputs(13)
    parYboltHole = arrayOfInputs(14)
    
    
   
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
 ' Insert your code for output 1(SaddlePlate Body)
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oLine As IngrGeom3D.Line3d
    Dim oEleCollection           As Collection
    Set oEleCollection = New Collection
    
    Dim dBackingPlateAngle As Double
    Dim dSupportPlateAngle As Double
    Dim dBackingPlateArcLength As Double
    Dim dLength As Double
    Dim dBackingPlateProjFromSupportPlate As Double  ''Parameter S in the sketch
    
    dBackingPlateProjFromSupportPlate = 0.14  ''5 1/2 inch
    
    ''dBackingPlateProjFromSupportPlate is an extra Projection of saddle from the Support plate( S )
    
    dLength = Sqr(Abs((parVesselDiameter / 2 + parBackingPlateThickness) ^ 2 - (parSupportLength / 2) ^ 2))
    dSupportPlateAngle = 2 * Atn((parSupportLength / 2) / dLength)
    dBackingPlateArcLength = (parVesselDiameter / 2) * dSupportPlateAngle + 2 * dBackingPlateProjFromSupportPlate
    dBackingPlateAngle = dBackingPlateArcLength / (parVesselDiameter / 2)
    
    Dim dlinepts() As Double
    ReDim dlinepts(0 To 17) As Double
    
    ''SaddlePlate Points
    dlinepts(0) = parCPtoSaddle + parBackingPlateWidth / 2
    dlinepts(1) = (parVesselDiameter / 2) * Sin(dBackingPlateAngle / 2)
    dlinepts(2) = -(parVesselDiameter / 2) * Cos(dBackingPlateAngle / 2)

    dlinepts(3) = dlinepts(0)
    dlinepts(4) = 0
    dlinepts(5) = -(parVesselDiameter / 2)
    
    dlinepts(6) = dlinepts(0)
    dlinepts(7) = -dlinepts(1)
    dlinepts(8) = dlinepts(2)
    
    dlinepts(9) = dlinepts(0)
    dlinepts(10) = -(parVesselDiameter / 2 + parBackingPlateThickness) * Sin(dBackingPlateAngle / 2)
    dlinepts(11) = -(parVesselDiameter / 2 + parBackingPlateThickness) * Cos(dBackingPlateAngle / 2)
    
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = 0
    dlinepts(14) = -(parVesselDiameter / 2 + parBackingPlateThickness)
    
    dlinepts(15) = dlinepts(0)
    dlinepts(16) = (parVesselDiameter / 2 + parBackingPlateThickness) * Sin(dBackingPlateAngle / 2)
    dlinepts(17) = -(parVesselDiameter / 2 + parBackingPlateThickness) * Cos(dBackingPlateAngle / 2)
    
    'Use Alternate Approach, Create a surface of revolution instead of projection, it is becuase constraint
    'mechanisms does not work with projections well if you have a complex string. Constraint design is to
    'Aligh the centroid of the projection. Since projection vector does not give the real location it gives
    'only the direction. for Constraint we need both location and direction. hence the code is modified
    Dim oStPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
    Dim oAxisVec As New AutoMath.DVector
    Dim dhalfSweepAngle As Double
    'Construct the First Cross section line
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                     dlinepts(0) - parBackingPlateWidth, dlinepts(1), dlinepts(2))
    oEleCollection.Add oLine
    Set oLine = Nothing
    'Construct the Second Cross section line (Second point to 3rd point)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    dlinepts(0) - parBackingPlateWidth, dlinepts(1), dlinepts(2), _
                                    dlinepts(15) - parBackingPlateWidth, dlinepts(16), dlinepts(17))
                                    
    oEleCollection.Add oLine
    Set oLine = Nothing
    'Construct the Third Cross section line (3rd point to 4th point)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    dlinepts(15) - parBackingPlateWidth, dlinepts(16), dlinepts(17), _
                                    dlinepts(15), dlinepts(16), dlinepts(17))
                                    
    oEleCollection.Add oLine
    Set oLine = Nothing
    'Construct the 4th Cross section line (4th point to 1st point)
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    dlinepts(15), dlinepts(16), dlinepts(17), _
                                    dlinepts(0), dlinepts(1), dlinepts(2))
                                    
    oEleCollection.Add oLine
    Set oLine = Nothing
    'Now Create a Complex string out of the above lines
    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
    'Create a surface of revolution
    oAxisVec.Set -1#, 0#, 0#
    oStPoint.Set parCPtoSaddle - parSupportWidth / 2, 0#, 0#
    dhalfSweepAngle = Atn(Abs(dlinepts(1) / dlinepts(2)))
    Set ObjSaddlePlate = PlaceRevolution(m_OutputColl, oComplexStr, oAxisVec, oStPoint, 2 * dhalfSweepAngle, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSaddlePlate
    Set ObjSaddlePlate = Nothing
    
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
    
    Dim count As Integer
    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count
    
    
'Insert your code for output 2(Ends Support Plate)
    
    ReDim dlinepts(0 To 14) As Double
    
    dlinepts(0) = parCPtoSaddle + parStiffnerThickness / 2
    dlinepts(1) = (parVesselDiameter / 2 + parBackingPlateThickness) * Sin(dSupportPlateAngle / 2)
    dlinepts(2) = -(parVesselDiameter / 2 + parBackingPlateThickness) * Cos(dSupportPlateAngle / 2)

    dlinepts(3) = dlinepts(0)
    dlinepts(4) = parSupportLength / 2
    dlinepts(5) = -parSupportHeight + parBasePlateThickness

    dlinepts(6) = dlinepts(0)
    dlinepts(7) = -dlinepts(4)
    dlinepts(8) = dlinepts(5)
    
    dlinepts(9) = dlinepts(0)
    dlinepts(10) = -dlinepts(1)
    dlinepts(11) = dlinepts(2)
     
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = 0
    dlinepts(14) = -(parVesselDiameter / 2 + parBackingPlateThickness)
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                      dlinepts(3), dlinepts(4), dlinepts(5))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                      dlinepts(6), dlinepts(7), dlinepts(8))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oArc
    Set oArc = Nothing
    
    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    oAxisVec.Set -1, 0, 0
    
    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
        
    Set ObjStiffnerPlate = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, parStiffnerThickness, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStiffnerPlate
    Set ObjStiffnerPlate = Nothing
    
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
    
    For count = 1 To oEleCollection.count
        oEleCollection.Remove 1
    Next count
    
''Note:The same symbol is used for both 8ft vesseldia and 10ft vesseldia. Number of Support plates depends
      ''on Vessel Dia. So support spacing is calculte based on the Vessel dia.
    
    
    Dim dSpacingAngle As Double ''Variable for angle between Spacings
         
    ' Insert your code for output 3 ( Center Support Plate1 )
    
       dSpacingAngle = Atn((parStiffnerSpacing / 2) / (Sqr(Abs((parVesselDiameter / 2 + parBackingPlateThickness) ^ 2 - (parStiffnerSpacing / 2) ^ 2))))
       
       oStPoint.Set parCPtoSaddle + parSupportWidth / 2, -parStiffnerSpacing / 2 + parStiffnerThickness1, _
                    -parSupportHeight + parBasePlateThickness
                    
       oEndPoint.Set parCPtoSaddle - parSupportWidth / 2, -parStiffnerSpacing / 2, _
                     -(parVesselDiameter / 2 + parBackingPlateThickness) * Cos(dSpacingAngle)
       
       Set ObjBackingPlate1 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)
    
    ' Set the output
       iOutput = iOutput + 1
       m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate1
            
         
' Insert your code for output 4( Middle Support Plate )
            
        oStPoint.Set parCPtoSaddle + parSupportWidth / 2, -parStiffnerThickness1 / 2, _
                    -parSupportHeight + parBasePlateThickness
                    
        oEndPoint.Set parCPtoSaddle - parSupportWidth / 2, parStiffnerThickness1 / 2, _
                     -(parVesselDiameter / 2 + parBackingPlateThickness)
       
        Set ObjBackingPlate2 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)
    
    ' Set the output
       iOutput = iOutput + 1
       m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate2
       Set ObjBackingPlate2 = Nothing
        
    ' Insert your code for output 5 ( Center Support Plate2 )
       
       oStPoint.Set parCPtoSaddle + parSupportWidth / 2, parStiffnerSpacing / 2 - parStiffnerThickness1, _
                    -parSupportHeight + parBasePlateThickness
                    
       oEndPoint.Set parCPtoSaddle - parSupportWidth / 2, parStiffnerSpacing / 2, _
                     -(parVesselDiameter / 2 + parBackingPlateThickness) * Cos(dSpacingAngle)
       
       Set ObjBackingPlate3 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)
    
    ' Set the output
       iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate3
           Set ObjBackingPlate3 = Nothing

'The check has been modified such that only saddles of diameter greater than 1.8m will have the two extra supports
If CmpDblLessThanOrEqualTo(parVesselDiameter, 1.8) Then
        ''If vessel dia is less than or equal to  1.8 mts then skipping two outputs.
        iOutput = iOutput + 2
Else ''this is for vessel dia more than 1.8mts it has Two additional Stiffner plates(two more outputs)
    'A check is being inserted here so that the two extra supports are created only when the stiffener spacing
    'is less than or equal to half of support length even though the vessel diameter is greater than 1.8 meters
    If CmpDblLessThanOrEqualTo(parStiffnerSpacing, parSupportLength / 2) Then
 
         ' Insert your code for output 6( Center Support Plate 1 )

            dSpacingAngle = Atn((parStiffnerSpacing) / (Sqr(Abs((parVesselDiameter / 2 + parBackingPlateThickness) ^ 2 - (parStiffnerSpacing) ^ 2))))

            oStPoint.Set parCPtoSaddle + parSupportWidth / 2, -parStiffnerSpacing + parStiffnerThickness1, _
                         -parSupportHeight + parBasePlateThickness

            oEndPoint.Set parCPtoSaddle - parSupportWidth / 2, -parStiffnerSpacing, _
                          -(parVesselDiameter / 2 + parBackingPlateThickness) * Cos(dSpacingAngle)

            Set ObjBackingPlate4 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

         ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate4
            Set ObjBackingPlate4 = Nothing
                

         ' Insert your code for output 7( Center Support Plate4 )

            oStPoint.Set parCPtoSaddle + parSupportWidth / 2, parStiffnerSpacing - parStiffnerThickness1, _
                         -parSupportHeight + parBasePlateThickness

            oEndPoint.Set parCPtoSaddle - parSupportWidth / 2, parStiffnerSpacing, _
                          -(parVesselDiameter / 2 + parBackingPlateThickness) * Cos(dSpacingAngle)

            Set ObjBackingPlate5 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

         ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate5
            Set ObjBackingPlate5 = Nothing
    End If
  End If
  
        
' Insert your code for output 8( Bottom Support Plate )
   
   oStPoint.Set parCPtoSaddle + parSupportWidth / 2, parSupportLength / 2, _
                -parSupportHeight
                
   oEndPoint.Set parCPtoSaddle - parSupportWidth / 2, -parSupportLength / 2, _
                 -parSupportHeight + parBasePlateThickness
   
   Set ObjBackingPlate6 = PlaceBox(m_OutputColl, oStPoint, oEndPoint)

' Set the output
   iOutput = iOutput + 1
   m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBackingPlate6
   Set ObjBackingPlate6 = Nothing
   
' Insert your code for output 9( Fundation Port)

    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    
      'Coordinate sysytem is at the bottom of support.
    '
    '               |-------|
    '   Y           |       |
    '   ^           |       |
    '   |           |       |
    '   |           |       |
    '   |           |       |
    '   -----> X    |       |
    '   Symbol CS   |       |
    '               |       |
    '               |   X   |
    '               |   ^   |
    '               |   |   |
    '               |   |   |
    '               |---|---|----> Y
    '                   Port CS
    
    
    Dim objEqpFoundationPort As IJEqpFoundationPort
    Set objEqpFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "EqpFoundationPort", _
                                                False, m_OutputColl.ResourceManager)
    Dim holes() As Variant
    Call objEqpFoundationPort.GetHoles(holes())
    holes(0, 1) = (parSupportLength - parYboltHole) / 2
    holes(0, 2) = parXboltHole
    holes(1, 1) = (parSupportLength + parYboltHole) / 2
    holes(1, 2) = parXboltHole
    holes(2, 1) = (parSupportLength + parYboltHole) / 2
    holes(2, 2) = -parXboltHole
    holes(3, 1) = (parSupportLength - parYboltHole) / 2
    holes(3, 2) = -parXboltHole

    Call objEqpFoundationPort.PutCS(parCPtoSaddle, -parSupportLength / 2, -parSupportHeight, 0, 1, 0, 0, 0, -1)
    Call objEqpFoundationPort.SetHoles(holes)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEqpFoundationPort
    Set objEqpFoundationPort = Nothing
    
    Set NozzlePHFactory = Nothing
    
    Set oStPoint = Nothing
    Set oEndPoint = Nothing

    Set oEleCollection = Nothing
    Set oGeomFactory = Nothing

    Exit Sub

ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

