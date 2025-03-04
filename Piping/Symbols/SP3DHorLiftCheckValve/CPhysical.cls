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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          ACM
'   Creation Date:  Tuesday, Nov 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Horizontal Lift Check Valve, Taken from www.resistoflex.com web catalog.
'   This symbol has total 11 outputs among them 9 are Physical and 2 are Insulation.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   01.Aug.2006     kkk       TR: 80254 Replaced 'm_OutputColl.ResourceManager' with "Nothing'.
'   02.Jan.2006     kkk       CR-91106  Updated the symbol to be based on either face-to-face
'                                           or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)       Modified to accept default value of lPartDataBasis to be less than or equal to 1.
'   25.sep.06       KKC            TR-102869 Modified symbol code to use Face to Center or Face to Face to Dimension.
'                                            Modified symbol code to use Face to Face as a input parameter.
'   17.Jul.2007     MA             CR-123497 Implemented Part data basis to use Face 1 to Center and Face 2 to Center Dimensions.
'   07.Sept.2007    RRK            CR-126919  Enhanced horizontal lift check valve to support Y pattern by adding three part data basis values(385,386 and 387)
'                                             Added some of the existing out puts to variable collection
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim PI As Double
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''
PI = 4 * Atn(1)

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam1      As Double
    Dim depth1           As Double
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjValBodyCyl As Object
    Dim ObjValBodyRev As Object
    Dim ObjFillerBody As Object
    Dim ObjBotRectFlange As Object
    Dim ObjTopRectFlange  As Object
    Dim ObjFlangeFiller1 As Object
    Dim ObjFlangeFiller2 As Object
    
    Dim parFacetoFace As Double
    Dim parValveHeight As Double
    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parBranchLength As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parValveHeight = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'    parFacetoCenter=arrayOfInputs(5)
'    parFace1toCenter = arrayOfInputs(6)
'    parFace2toCenter = arrayOfInputs(7)
'    parBranchLength = arrayOfInputs(8)

    iOutput = 0
    
' Insert your code for output 1(Valve Cylindrical Body)
    
   RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
   RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset1, depth1
 
 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -5): Face-to-face dimension basis
'   (Part data Basis value -10): Face-to-center dimension basis
'   (Part data Basis value -13): Asymmetrical Face-to-Center dimension basis
'   (Part data Basis value -385): Horizontal lift check valve, Y pattern, face-to-face dimension basis
'   (Part data Basis value -386): Horizontal lift check valve, Y pattern, face-to-center dimension basis
'   (Part data Basis value -387): Horizontal lift check valve, Y pattern, asymmetrical face-to-center dimension basis
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 5
        parFacetoFace = arrayOfInputs(2)
        parValveHeight = arrayOfInputs(3)
        
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
    
    Case 10
        parValveHeight = arrayOfInputs(3)
        parFacetoCenter = arrayOfInputs(5)
        
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13
        parValveHeight = arrayOfInputs(3)
        parFace1toCenter = arrayOfInputs(6)
        parFace2toCenter = arrayOfInputs(7)
        
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        
    Case 385
        parFacetoFace = arrayOfInputs(2)
        parBranchLength = arrayOfInputs(8)
        
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2

    Case 386
        parFacetoCenter = arrayOfInputs(5)
        parBranchLength = arrayOfInputs(8)
        
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
        
    Case 387
        parFace1toCenter = arrayOfInputs(6)
        parFace2toCenter = arrayOfInputs(7)
        parBranchLength = arrayOfInputs(8)
        
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
 
    Case Else
        GoTo ErrorLabel:
    End Select
    
    If CmpDblLessThanOrEqualTo(flangeThick1, -LINEAR_TOLERANCE) Then
       flangeThick1 = NEGLIGIBLE_THICKNESS
    ElseIf CmpDblLessThanOrEqualTo(flangeThick2, -LINEAR_TOLERANCE) Then
       flangeThick2 = NEGLIGIBLE_THICKNESS
    End If
    
   Dim oStPoint As AutoMath.DPosition
   Dim oEnPoint As AutoMath.DPosition
   Set oStPoint = New AutoMath.DPosition
   Set oEnPoint = New AutoMath.DPosition
   
   oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
   oEnPoint.Set dFace2toCenter - flangeThick2, 0, 0
      
   Set ObjValBodyCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam1, True)
' Set the output
   iOutput = iOutput + 1
   m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValBodyCyl
   Set ObjValBodyCyl = Nothing

    Select Case lPartDataBasis
    
    Case Is <= 1, 5, 10, 13
'Default Horizontal Lift Check Valve Cases
    ' Insert your code for output 2(Filler of Valve Body)
       Dim oGeomFactory    As IngrGeom3D.GeometryFactory
       Dim oComplexStr As IngrGeom3D.ComplexString3d
       Dim dFillerPoints(0 To 8) As Double
       Dim oLine As IngrGeom3D.Line3d
       Dim oArc As IngrGeom3D.Arc3d
       Dim oObjColl As New Collection
       
       Set oGeomFactory = New IngrGeom3D.GeometryFactory
       
        Dim dMinFacetoCenter As Double
        
        If CmpDblLessThan(dFace1toCenter, dFace2toCenter) Then
            dMinFacetoCenter = dFace1toCenter
        Else
            dMinFacetoCenter = dFace2toCenter
        End If
        
       dFillerPoints(0) = (dMinFacetoCenter) - (2 * flangeThick1)
       dFillerPoints(1) = 0
       dFillerPoints(2) = pipeDiam1 / 2
       
       dFillerPoints(3) = (dMinFacetoCenter) - (2 * flangeThick1)
       dFillerPoints(4) = pipeDiam1 / 2
       dFillerPoints(5) = pipeDiam1 / 2
       
       dFillerPoints(6) = (dMinFacetoCenter) - (2 * flangeThick1)
       dFillerPoints(7) = pipeDiam1 / 2
       dFillerPoints(8) = 0
       
       Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                dFillerPoints(0), dFillerPoints(1), dFillerPoints(2), _
                                dFillerPoints(3), dFillerPoints(4), dFillerPoints(5))
       oObjColl.Add oLine
       
       Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dFillerPoints(3), dFillerPoints(4), dFillerPoints(5), _
                                        dFillerPoints(3), dFillerPoints(4), -dFillerPoints(5))
       oObjColl.Add oLine
       
       Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dFillerPoints(3), dFillerPoints(4), -dFillerPoints(5), _
                                        dFillerPoints(0), dFillerPoints(1), -dFillerPoints(2))
       oObjColl.Add oLine
          
       Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                        dFillerPoints(0), dFillerPoints(1), -dFillerPoints(2), _
                                        dFillerPoints(6), dFillerPoints(7), dFillerPoints(8), _
                                        dFillerPoints(0), dFillerPoints(1), dFillerPoints(2))
       oObjColl.Add oArc
       
       oStPoint.Set dFillerPoints(0), dFillerPoints(1), dFillerPoints(2)
       Set oComplexStr = PlaceTrCString(oStPoint, oObjColl)
       
       Dim oDir As New AutoMath.DVector
       oDir.Set -1, 0, 0
       
       Set ObjFillerBody = PlaceProjection(m_OutputColl, oComplexStr, oDir, _
                                            2 * ((dMinFacetoCenter) - (2 * flangeThick1)), True)
    
    
    ' Set the output
        m_OutputColl.AddOutput "FillerBody", ObjFillerBody
        Set ObjFillerBody = Nothing
        
        Dim iCount As Integer
        For iCount = 1 To oObjColl.Count
            oObjColl.Remove 1
        Next iCount
       
    ' Insert your code for output 3(Bottom Rectangular Flange)
        Dim dFlangePoints(0 To 35) As Double
    
        dFlangePoints(0) = (dMinFacetoCenter) - (2 * flangeThick1)
        dFlangePoints(1) = pipeDiam1 / 2
        dFlangePoints(2) = pipeDiam1 / 2
        ''centerpoint of arc1
        dFlangePoints(3) = pipeDiam1 / 2
        dFlangePoints(4) = pipeDiam1 / 2
        dFlangePoints(5) = pipeDiam1 / 2
        
        dFlangePoints(6) = pipeDiam1 / 2
        dFlangePoints(7) = pipeDiam1 / 2
        dFlangePoints(8) = ((dMinFacetoCenter) - (2 * flangeThick1))
        
        dFlangePoints(9) = -pipeDiam1 / 2
        dFlangePoints(10) = pipeDiam1 / 2
        dFlangePoints(11) = ((dMinFacetoCenter) - (2 * flangeThick1))
        ''centerpoint of arc2
        dFlangePoints(12) = -pipeDiam1 / 2
        dFlangePoints(13) = pipeDiam1 / 2
        dFlangePoints(14) = pipeDiam1 / 2
        
        dFlangePoints(15) = -((dMinFacetoCenter) - (2 * flangeThick1))
        dFlangePoints(16) = pipeDiam1 / 2
        dFlangePoints(17) = pipeDiam1 / 2
        
        dFlangePoints(18) = -((dMinFacetoCenter) - (2 * flangeThick1))
        dFlangePoints(19) = pipeDiam1 / 2
        dFlangePoints(20) = -pipeDiam1 / 2
        ''centerpoint of arc3
        dFlangePoints(21) = -pipeDiam1 / 2
        dFlangePoints(22) = pipeDiam1 / 2
        dFlangePoints(23) = -pipeDiam1 / 2
        
        dFlangePoints(24) = -pipeDiam1 / 2
        dFlangePoints(25) = pipeDiam1 / 2
        dFlangePoints(26) = -((dMinFacetoCenter) - (2 * flangeThick1))
        
        dFlangePoints(27) = pipeDiam1 / 2
        dFlangePoints(28) = pipeDiam1 / 2
        dFlangePoints(29) = -((dMinFacetoCenter) - (2 * flangeThick1))
        ''centerpoint of arc4
        dFlangePoints(30) = pipeDiam1 / 2
        dFlangePoints(31) = pipeDiam1 / 2
        dFlangePoints(32) = -pipeDiam1 / 2
        
        dFlangePoints(33) = (dMinFacetoCenter) - (2 * flangeThick1)
        dFlangePoints(34) = pipeDiam1 / 2
        dFlangePoints(35) = -pipeDiam1 / 2
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        dFlangePoints(3), dFlangePoints(4), dFlangePoints(5), _
                                        dFlangePoints(0), dFlangePoints(1), dFlangePoints(2), _
                                        dFlangePoints(6), dFlangePoints(7), dFlangePoints(8))
        oObjColl.Add oArc
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dFlangePoints(6), dFlangePoints(7), dFlangePoints(8), _
                                        dFlangePoints(9), dFlangePoints(10), dFlangePoints(11))
        oObjColl.Add oLine
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        dFlangePoints(12), dFlangePoints(13), dFlangePoints(14), _
                                        dFlangePoints(9), dFlangePoints(10), dFlangePoints(11), _
                                        dFlangePoints(15), dFlangePoints(16), dFlangePoints(17))
        oObjColl.Add oArc
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dFlangePoints(15), dFlangePoints(16), dFlangePoints(17), _
                                        dFlangePoints(18), dFlangePoints(19), dFlangePoints(20))
        oObjColl.Add oLine
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        dFlangePoints(21), dFlangePoints(22), dFlangePoints(23), _
                                        dFlangePoints(18), dFlangePoints(19), dFlangePoints(20), _
                                        dFlangePoints(24), dFlangePoints(25), dFlangePoints(26))
        oObjColl.Add oArc
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dFlangePoints(24), dFlangePoints(25), dFlangePoints(26), _
                                        dFlangePoints(27), dFlangePoints(28), dFlangePoints(29))
        oObjColl.Add oLine
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        dFlangePoints(30), dFlangePoints(31), dFlangePoints(32), _
                                        dFlangePoints(27), dFlangePoints(28), dFlangePoints(29), _
                                        dFlangePoints(33), dFlangePoints(34), dFlangePoints(35))
        oObjColl.Add oArc
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dFlangePoints(33), dFlangePoints(34), dFlangePoints(35), _
                                        dFlangePoints(0), dFlangePoints(1), dFlangePoints(2))
        oObjColl.Add oLine
        
        oStPoint.Set dFlangePoints(0), dFlangePoints(1), dFlangePoints(2)
        
        Set oComplexStr = PlaceTrCString(oStPoint, oObjColl)
        oDir.Set 0, 1, 0
        ''Bottom Flange thickness is taken as flangeThick1
        Set ObjBotRectFlange = PlaceProjection(m_OutputColl, oComplexStr, oDir, flangeThick1, True)
    ' Set the output
        m_OutputColl.AddOutput "BotRectFlange", ObjBotRectFlange
        Set ObjBotRectFlange = Nothing
        
        For iCount = 1 To oObjColl.Count
            oObjColl.Remove 1
        Next iCount
        
    ' Insert your code for output 4(Top Rectangular Flange)
      ''Gap between two flanges is taken as 2 mm
      ''Top Rectangular flange is at a distance of flangethick1+0.002 in the Y-dirrection.
      ''Top Flange thickenss is taken as flangethick1
        Dim FlangeToFlangeDist As Double
        FlangeToFlangeDist = flangeThick1 + 0.002
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                    dFlangePoints(3), dFlangePoints(4) + FlangeToFlangeDist, dFlangePoints(5), _
                    dFlangePoints(0), dFlangePoints(1) + FlangeToFlangeDist, dFlangePoints(2), _
                    dFlangePoints(6), dFlangePoints(7) + FlangeToFlangeDist, dFlangePoints(8))
        oObjColl.Add oArc
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                    dFlangePoints(6), dFlangePoints(7) + FlangeToFlangeDist, dFlangePoints(8), _
                    dFlangePoints(9), dFlangePoints(10) + FlangeToFlangeDist, dFlangePoints(11))
        oObjColl.Add oLine
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                    dFlangePoints(12), dFlangePoints(13) + FlangeToFlangeDist, dFlangePoints(14), _
                    dFlangePoints(9), dFlangePoints(10) + FlangeToFlangeDist, dFlangePoints(11), _
                    dFlangePoints(15), dFlangePoints(16) + FlangeToFlangeDist, dFlangePoints(17))
        oObjColl.Add oArc
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                    dFlangePoints(15), dFlangePoints(16) + FlangeToFlangeDist, dFlangePoints(17), _
                    dFlangePoints(18), dFlangePoints(19) + FlangeToFlangeDist, dFlangePoints(20))
        oObjColl.Add oLine
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                    dFlangePoints(21), dFlangePoints(22) + FlangeToFlangeDist, dFlangePoints(23), _
                    dFlangePoints(18), dFlangePoints(19) + FlangeToFlangeDist, dFlangePoints(20), _
                    dFlangePoints(24), dFlangePoints(25) + FlangeToFlangeDist, dFlangePoints(26))
        oObjColl.Add oArc
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                    dFlangePoints(24), dFlangePoints(25) + FlangeToFlangeDist, dFlangePoints(26), _
                    dFlangePoints(27), dFlangePoints(28) + FlangeToFlangeDist, dFlangePoints(29))
        oObjColl.Add oLine
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                    dFlangePoints(30), dFlangePoints(31) + FlangeToFlangeDist, dFlangePoints(32), _
                    dFlangePoints(27), dFlangePoints(28) + FlangeToFlangeDist, dFlangePoints(29), _
                    dFlangePoints(33), dFlangePoints(34) + FlangeToFlangeDist, dFlangePoints(35))
        oObjColl.Add oArc
        
        Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                    dFlangePoints(33), dFlangePoints(34) + FlangeToFlangeDist, dFlangePoints(35), _
                    dFlangePoints(0), dFlangePoints(1) + FlangeToFlangeDist, dFlangePoints(2))
        oObjColl.Add oLine
        
        oStPoint.Set dFlangePoints(0), dFlangePoints(1) + FlangeToFlangeDist, dFlangePoints(2)
        
        Set oComplexStr = PlaceTrCString(oStPoint, oObjColl)
        oDir.Set 0, 1, 0
        
        Set ObjTopRectFlange = PlaceProjection(m_OutputColl, oComplexStr, oDir, flangeThick1, True)
    ' Set the output
        m_OutputColl.AddOutput "TopRectFlange", ObjTopRectFlange
        Set ObjTopRectFlange = Nothing
       
        For iCount = 1 To oObjColl.Count
            oObjColl.Remove 1
        Next iCount
    
    ' Insert your code for output 5(Flange Filler Body1)
      Dim dFlangeFillerPts(0 To 11) As Double
      
      dFlangeFillerPts(0) = pipeDiam1 / 2
      dFlangeFillerPts(1) = pipeDiam1 / 2
      dFlangeFillerPts(2) = -((dMinFacetoCenter) - (2 * flangeThick1))
      
      dFlangeFillerPts(3) = pipeDiam1 / 2
      dFlangeFillerPts(4) = pipeDiam1 / 2
      dFlangeFillerPts(5) = -pipeDiam1 / 2
      
      dFlangeFillerPts(6) = pipeDiam1 / 2
      dFlangeFillerPts(7) = 0
      dFlangeFillerPts(8) = -pipeDiam1 / 2
      
      dFlangeFillerPts(9) = pipeDiam1 / 2
      dFlangeFillerPts(10) = pipeDiam1 / 4
      dFlangeFillerPts(11) = -0.6 * ((dMinFacetoCenter) - (2 * flangeThick1))
      
      Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                dFlangeFillerPts(0), dFlangeFillerPts(1), dFlangeFillerPts(2), _
                                dFlangeFillerPts(3), dFlangeFillerPts(4), dFlangeFillerPts(5))
                                                                  
      oObjColl.Add oLine
      
      Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                dFlangeFillerPts(3), dFlangeFillerPts(4), dFlangeFillerPts(5), _
                                dFlangeFillerPts(6), dFlangeFillerPts(7), dFlangeFillerPts(8))
                                                                
      
      oObjColl.Add oLine
      
    
        Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                dFlangeFillerPts(6), dFlangeFillerPts(7), dFlangeFillerPts(8), _
                                dFlangeFillerPts(9), dFlangeFillerPts(10), dFlangeFillerPts(11), _
                                dFlangeFillerPts(0), dFlangeFillerPts(1), dFlangeFillerPts(2))
                                                                
      oObjColl.Add oArc
      
      oStPoint.Set dFlangeFillerPts(0), dFlangeFillerPts(1), dFlangeFillerPts(2)
      Set oComplexStr = PlaceTrCString(oStPoint, oObjColl)
      
      oDir.Set -1, 0, 0
      Set ObjFlangeFiller1 = PlaceProjection(m_OutputColl, oComplexStr, oDir, pipeDiam1, True)
    ' Set the output
        m_OutputColl.AddOutput "FlangeFiller1", ObjFlangeFiller1
        Set ObjFlangeFiller1 = Nothing
        
        For iCount = 1 To oObjColl.Count
            oObjColl.Remove 1
        Next iCount
    ' Insert your code for output 6(Flange Filler Body2)
    
      Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                dFlangeFillerPts(0), dFlangeFillerPts(1), -dFlangeFillerPts(2), _
                                dFlangeFillerPts(3), dFlangeFillerPts(4), -dFlangeFillerPts(5))
                                                                  
      oObjColl.Add oLine
      
      Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                dFlangeFillerPts(3), dFlangeFillerPts(4), -dFlangeFillerPts(5), _
                                dFlangeFillerPts(6), dFlangeFillerPts(7), -dFlangeFillerPts(8))
                                                                  
      oObjColl.Add oLine
      
      
      Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, _
                                dFlangeFillerPts(6), dFlangeFillerPts(7), -dFlangeFillerPts(8), _
                                dFlangeFillerPts(9), dFlangeFillerPts(10), -dFlangeFillerPts(11), _
                                dFlangeFillerPts(0), dFlangeFillerPts(1), -dFlangeFillerPts(2))
                                                                
      oObjColl.Add oArc
      
      oStPoint.Set dFlangeFillerPts(0), dFlangeFillerPts(1), -dFlangeFillerPts(2)
      Set oComplexStr = PlaceTrCString(oStPoint, oObjColl)
    
      oDir.Set -1, 0, 0
      Set ObjFlangeFiller2 = PlaceProjection(m_OutputColl, oComplexStr, oDir, pipeDiam1, True)
    ' Set the output
        m_OutputColl.AddOutput "FlangeFiller2", ObjFlangeFiller2
        Set ObjFlangeFiller2 = Nothing
        
        For iCount = 1 To oObjColl.Count
            oObjColl.Remove 1
        Next iCount
    
    ' Insert your code for output 5(Valve Body Upper portion)
      Dim dBodyCstrPts(0 To 23) As Double
      
       dBodyCstrPts(0) = 0
       dBodyCstrPts(1) = (pipeDiam1 / 2) + 2 * flangeThick1 + 0.002 ''Gap between two flanges is 2mm
       dBodyCstrPts(2) = 0
       
       dBodyCstrPts(3) = (dMinFacetoCenter) - 3 * flangeThick1
       dBodyCstrPts(4) = (pipeDiam1 / 2) + 2 * flangeThick1 + 0.002 ''Gap between two flanges is 2mm
       dBodyCstrPts(5) = 0
       ''Center Point of elliptical section
       dBodyCstrPts(6) = 0.3 * ((dMinFacetoCenter) - 2 * flangeThick1)
       dBodyCstrPts(7) = (pipeDiam1 / 2) + 2 * flangeThick1 + 0.002 ''Gap between two flanges is 2mm
       dBodyCstrPts(8) = 0
       
       dBodyCstrPts(9) = 0.3 * ((dMinFacetoCenter) - 2 * flangeThick1)
       dBodyCstrPts(10) = 0.6 * parValveHeight
       dBodyCstrPts(11) = 0
       
       dBodyCstrPts(12) = 0.3 * ((dMinFacetoCenter) - 2 * flangeThick1)
       dBodyCstrPts(13) = 0.75 * parValveHeight
       dBodyCstrPts(14) = 0
       ''Centre point of arc
       dBodyCstrPts(15) = 0.2 * ((dMinFacetoCenter) - 2 * flangeThick1)
       dBodyCstrPts(16) = 0.75 * parValveHeight
       dBodyCstrPts(17) = 0
       
       dBodyCstrPts(18) = 0.2 * ((dMinFacetoCenter) - 2 * flangeThick1)
       dBodyCstrPts(19) = parValveHeight
       dBodyCstrPts(20) = 0
       
       dBodyCstrPts(21) = 0
       dBodyCstrPts(22) = parValveHeight
       dBodyCstrPts(23) = 0
      
      Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dBodyCstrPts(0), dBodyCstrPts(1), dBodyCstrPts(2), _
                                        dBodyCstrPts(3), dBodyCstrPts(4), dBodyCstrPts(5))
      oObjColl.Add oLine
       
        Dim ellipseMMRatio  As Double
        Dim ellipseMajorX As Double
        Dim ellipseMajorY As Double
        Dim ellipseMajorZ As Double
        Dim oEllipticalArc As IngrGeom3D.EllipticalArc3d
        
        ellipseMMRatio = (0.6 * parValveHeight - ((pipeDiam1 / 2) + 2 * flangeThick1)) / (0.7 * ((dMinFacetoCenter) - 2 * flangeThick1) - flangeThick1)
        ellipseMajorX = 0.7 * ((dMinFacetoCenter) - 2 * flangeThick1) - flangeThick1
        ellipseMajorY = 0
        ellipseMajorZ = 0
    
      Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                dBodyCstrPts(6), dBodyCstrPts(7), dBodyCstrPts(8), 0, 0, 1, _
                                ellipseMajorX, ellipseMajorY, ellipseMajorZ, ellipseMMRatio, _
                                0, PI / 2)
      oObjColl.Add oEllipticalArc
    
      
      Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dBodyCstrPts(9), dBodyCstrPts(10), dBodyCstrPts(11), _
                                        dBodyCstrPts(12), dBodyCstrPts(13), dBodyCstrPts(14))
      oObjColl.Add oLine
      
      Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                        dBodyCstrPts(15), dBodyCstrPts(16), dBodyCstrPts(17), _
                                        dBodyCstrPts(12), dBodyCstrPts(13), dBodyCstrPts(14), _
                                        dBodyCstrPts(18), dBodyCstrPts(19), dBodyCstrPts(20))
      oObjColl.Add oArc
      
      Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        dBodyCstrPts(18), dBodyCstrPts(19), dBodyCstrPts(20), _
                                        dBodyCstrPts(21), dBodyCstrPts(22), dBodyCstrPts(23))
      oObjColl.Add oLine
    
      oStPoint.Set dBodyCstrPts(0), dBodyCstrPts(1), dBodyCstrPts(2)
      Set oComplexStr = PlaceTrCString(oStPoint, oObjColl)
      
      oDir.Set 0, 1, 0
      Dim oCenterPoint As New AutoMath.DPosition
      oCenterPoint.Set 0, parValveHeight / 2, 0
      
      Set ObjValBodyRev = PlaceRevolution(m_OutputColl, oComplexStr, oDir, oCenterPoint, _
                                                                             2 * PI, True)
    ' Set the output
        m_OutputColl.AddOutput "ValBodyRev", ObjValBodyRev
        Set ObjValBodyRev = Nothing
    
    Case 385, 386, 387
'Non-Default Horizontal Lift Check Valve Cases
      Dim objYBranch As Object
      Dim oCenterPos As AutoMath.DPosition
      Set oCenterPos = New AutoMath.DPosition
      
      Dim dAngle As Double
      Dim dAngle2 As Double
      Dim dTotAngle As Double
      Dim dBranchDia As Double
      Dim dLimBranchDia1 As Double
      Dim dLimBranchDia2 As Double
      Dim dLimLength1 As Double
      
      'Assumptions: Branch Angle is assumed to be equal to 45
      dAngle = PI / 4
      
      'Parameters useful for calculating limiting branch diameter
      dLimLength1 = Sqr((pipeDiam1 / 2) ^ 2 + (dFace2toCenter ^ 2))
      dAngle2 = Atn(0.5 * pipeDiam1 / dFace2toCenter)
      dTotAngle = dAngle + dAngle2
      
      'Calculating limiting branch diameter above which the top face of the branch crosses port 2
      dLimBranchDia1 = 2 * dLimLength1 * Cos(dTotAngle)
      'Calculating limiting branch diameter above which the bottom face of the branch crosses port 1
      dLimBranchDia2 = 2 * dFace1toCenter / Sin(dAngle)
      
      'Assumption for branch diameter equal to length of the line between ends of cylinder inclined at dAngle
      dBranchDia = pipeDiam1 / Cos(dAngle)
      
      'Checking with limiting branch diameter values
      If CmpDblGreaterthan(dBranchDia, dLimBranchDia1) Or CmpDblGreaterthan(dBranchDia, dLimBranchDia2) Then
        dBranchDia = 0.95 * SmallerDim(dLimBranchDia1, dLimBranchDia2)
      End If
    
      If lPartDataBasis = 385 Then 'Face to Face basis
        oCenterPos.Set -0.25 * parFacetoFace, 0, 0
      ElseIf lPartDataBasis = 386 Or lPartDataBasis = 387 Then 'Face to Center or Assymmetrical Face to Center basis
        oCenterPos.Set 0, 0, 0
      End If
      
      oStPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
      oEnPoint.Set oStPoint.x + parBranchLength * Cos(dAngle), parBranchLength * Sin(dAngle), 0
          
      Set objYBranch = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBranchDia, True)
    
    ' Set the output
      m_OutputColl.AddOutput "Branch", objYBranch
      Set objYBranch = Nothing
    
      'Creating filler in case branch diameter exceeds pipe outside diameter
      If CmpDblGreaterthan(dBranchDia, pipeDiam1) Then
          Dim objFiller As Object
          oStPoint.Set oCenterPos.x, 0, 0
          oEnPoint.Set oCenterPos.x - (pipeDiam1 / (2 * Tan(dAngle))), -pipeDiam1 / 2, 0
          Set objFiller = PlaceCone(m_OutputColl, oStPoint, oEnPoint, dBranchDia / 2, 0, True)
        ' Set the output
          m_OutputColl.AddOutput "Filler", objFiller
          Set objFiller = Nothing
       End If
         Set oCenterPos = Nothing
'End Select
    End Select
        
' Place Nozzle 1
   
    Dim oPlacePoint As AutoMath.DPosition
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    
    oPlacePoint.Set -dFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set dFace2toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oComplexStr = Nothing
    Set oArc = Nothing
    Set oEllipticalArc = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oGeomFactory = Nothing
    Set oCenterPoint = Nothing
    Set oObjColl = Nothing
    Set oLine = Nothing
   
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub


