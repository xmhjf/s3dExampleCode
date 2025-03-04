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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Feb 15 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This symbol is based on the following Part data basis value that govern its geometry,
'   1) Inspection elbow, specified by face-to-center - 960
'   2) Conduit elbow, type LB, specified by length, height, width, cover width and cover length - 9000 (Source: Thomas and Betts Fittings Catalog (http://www.tnb.com/contractor/docs/tc_gm101_xx.pdf), Page 41).
'   3) Conduit Elbow, type LL, specified by length, height, width, cover width and cover length - 9001 (Source: Thomas and Betts Fittings Catalog (http://www.tnb.com/contractor/docs/tc_gm101_xx.pdf), Page 44).
'   4) Conduit Elbow, type LR, specified by length, height, width, cover width and cover length - 9002 (Source: Thomas and Betts Fittings Catalog (http://www.tnb.com/contractor/docs/tc_gm101_xx.pdf), Page 45).
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   01.MAY.2008     MA     CR-67916     Implemented part data basis for the value 960 to provide inspection conduit elbow.
'   20.May.2008     dkl    CR 89059     Enhanced the elbow implementing part data bases - 9000, 9001 and 9002.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private m_oErrors As IJEditErrors
Private Const MODULE = "Physical:" 'Used for error messages
Private Const E_FAIL = -2147467259
Private Const COVERHEIGHT = 0.002    'Assumed height of cover
Private PI As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim ConduitOD1        As Double
    Dim ConduitOD2        As Double
    Dim flangeThick     As Double
    Dim cptOffset1       As Double
    Dim cptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim depth2          As Double
    
    Dim iOutput     As Integer
 
    Dim parFacetoCenter As Double
    Dim parLength               As Double
    Dim parHeight               As Double
    Dim parWidth                As Double
    Dim parCoverLength    As Double
    Dim parCoverWidth     As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
'    parLength = arrayOfInputs(3)
'    parHeight = arrayOfInputs(4)
'    parWidth = arrayOfInputs(5)
'    parCoverLength = arrayOfInputs(6)
'    parCoverWidth = arrayOfInputs(7)
    
    iOutput = 0

    'Insert your code for output 1(ConduitElbow)
    RetrieveParameters 1, oPartFclt, m_outputColl, ConduitOD1, flangeThick, flangeDiam, _
                                                                        cptOffset1, depth1
    ConduitOD1 = BodyOD(flangeDiam, ConduitOD1)
    RetrieveParameters 2, oPartFclt, m_outputColl, ConduitOD2, flangeThick, flangeDiam, _
                                                                        cptOffset2, depth2
    ConduitOD2 = BodyOD(flangeDiam, ConduitOD2)
    Dim ConduitBodyOD As Double
    ConduitBodyOD = BodyOD(ConduitOD1, ConduitOD2)

    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oPlacePoint1 As AutoMath.DPosition
    Dim oPlacePoint2 As AutoMath.DPosition
    Set oPlacePoint1 = New AutoMath.DPosition
    Set oPlacePoint2 = New AutoMath.DPosition
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    Select Case lPartDataBasis

        Case Is <= 1, 960
            parFacetoCenter = arrayOfInputs(2)
            'Nozzle locations
            oPlacePoint1.Set -parFacetoCenter - cptOffset1 + depth1, 0, 0
            oPlacePoint2.Set 0, parFacetoCenter + cptOffset2 - depth2, 0
            
            Dim CP As New AutoMath.DPosition 'revolution center point
            Dim CV As New AutoMath.DVector 'rotation vector for rotation
            Dim objCirc1 As IngrGeom3D.Circle3d
            Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
            Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
            Dim CenterPos       As New AutoMath.DPosition
            CenterPos.Set 0, 0, 0
            'Place Circle
            cirNormalX = 1
            cirNormalY = 0
            cirNormalZ = 0
             
            cir1CenterX = CenterPos.x - parFacetoCenter
            cir1CenterY = CenterPos.y
            cir1CenterZ = CenterPos.z
             
            Set objCirc1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                m_outputColl.ResourceManager, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, ConduitBodyOD / 2)
                                                         
            'Revolve the circle for 90 degree about the Z axis passing by the origin
            CP.Set -parFacetoCenter, parFacetoCenter, 0
            CV.Set 0, 0, 1
            
            Dim ObjConduitElbow As Object
            Set ObjConduitElbow = PlaceRevolution(m_outputColl, objCirc1, _
                             CV, CP, PI / 2, False)
            
            'Remove the Conduit circle
            Dim ob1 As IJDObject
            Set ob1 = objCirc1
            ob1.Remove
            Set ob1 = Nothing
            
            'Set the output
            iOutput = iOutput + 1
            m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitElbow
            Set ObjConduitElbow = Nothing
            Set CenterPos = Nothing

            If lPartDataBasis = 960 Then
                'Create a complex string for box cross-section
                Dim oLine1 As IngrGeom3D.Line3d
                Dim oLine2 As IngrGeom3D.Line3d
                Dim oLine3 As IngrGeom3D.Line3d
                Dim oLine4 As IngrGeom3D.Line3d
                'Line1
                oStPoint.Set -parFacetoCenter + flangeThick, 0.4 * ConduitBodyOD, 0
                oEnPoint.Set -parFacetoCenter + flangeThick, -0.4 * ConduitBodyOD, 0
            
                Set oLine1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
                'Line2
                oStPoint.Set -parFacetoCenter + flangeThick, -0.4 * ConduitBodyOD, 0
                oEnPoint.Set -parFacetoCenter + flangeThick, -0.4 * ConduitBodyOD, 0.5 * ConduitBodyOD
            
                Set oLine2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
                'Line3
                oStPoint.Set -parFacetoCenter + flangeThick, -0.4 * ConduitBodyOD, 0.5 * ConduitBodyOD
                oEnPoint.Set -parFacetoCenter + flangeThick, 0.4 * ConduitBodyOD, 0.5 * ConduitBodyOD
            
                Set oLine3 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
                'Line4
                oStPoint.Set -parFacetoCenter + flangeThick, 0.4 * ConduitBodyOD, 0.5 * ConduitBodyOD
                oEnPoint.Set -parFacetoCenter + flangeThick, 0.4 * ConduitBodyOD, 0
            
                Set oLine4 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            
                Dim oCurvesColl As Collection
                Set oCurvesColl = New Collection
                oCurvesColl.Add oLine1
                oCurvesColl.Add oLine2
                oCurvesColl.Add oLine3
                oCurvesColl.Add oLine4
                
                Dim oBox As IngrGeom3D.ComplexString3d
                oStPoint.Set -parFacetoCenter + flangeThick, 0.4 * ConduitBodyOD, 0
                Set oBox = PlaceTrCString(oStPoint, oCurvesColl)
            
                CP.Set -parFacetoCenter + flangeThick, parFacetoCenter - flangeThick, 0
                CV.Set 0, 0, 1
         
                Dim ObjElbow As Object
                Set ObjElbow = PlaceRevolution(m_outputColl, oBox, _
                                CV, CP, PI / 2, True)
        
                'Set the output
                m_outputColl.AddOutput "elbow", ObjElbow
                Set ObjElbow = Nothing
                Set oLine1 = Nothing
                Set oLine2 = Nothing
                Set oLine3 = Nothing
                Set oLine4 = Nothing
                Set oBox = Nothing
                Set oCurvesColl = Nothing
            End If
            Set CP = Nothing
            Set CV = Nothing
                
        Case 9000, 9001, 9002
            parLength = arrayOfInputs(3)
            parHeight = arrayOfInputs(4)
            parWidth = arrayOfInputs(5)
            parCoverLength = arrayOfInputs(6)
            parCoverWidth = arrayOfInputs(7)
            
            'Cover length is assumed to be 60% of length if not provided.
            If CmpDblLessThanOrEqualTo(parCoverLength, 0) Then parCoverLength = 0.6 * parLength
            'Cover width is assumed to be 80% of Width if not provided.
            If CmpDblLessThanOrEqualTo(parCoverWidth, 0) Then parCoverWidth = 0.8 * parHeight
            
            'Assumption: Body length is 85% of length.
            Dim dBodyLength As Double
            dBodyLength = 0.85 * parLength
                     
            'Assumption: Body Height is 90% of Height.
            Dim dBodyHeight As Double
            dBodyHeight = 0.9 * parHeight
            
            'Nozzle locations
            oPlacePoint1.Set -(parLength - ConduitOD2 / 2) - cptOffset1 + depth1, 0, 0
            oPlacePoint2.Set 0, parHeight + cptOffset2 - depth2, 0
            
            'Create Elbow body
            Dim oLineString    As IngrGeom3D.ComplexString3d
            Dim oCurves         As Collection
            Set oLineString = New IngrGeom3D.ComplexString3d
            Set oCurves = New Collection
                  
            Dim oLine As IngrGeom3D.Line3d
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                -(parWidth / 2 - ConduitOD2 / 2), -dBodyHeight / 2, parWidth / 2, _
                -(dBodyLength - ConduitOD2 / 2 - parWidth / 2), -dBodyHeight / 2, parWidth / 2)
            oCurves.Add oLine
            
            Dim oArc As IngrGeom3D.Arc3d
            Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                -(dBodyLength - ConduitOD2 / 2 - parWidth / 2), -dBodyHeight / 2, 0, _
                0, -1, 0, _
                -(dBodyLength - ConduitOD2 / 2 - parWidth / 2), -dBodyHeight / 2, parWidth / 2, _
                -(dBodyLength - ConduitOD2 / 2 - parWidth / 2), -dBodyHeight / 2, -parWidth / 2)
            oCurves.Add oArc
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                -(dBodyLength - ConduitOD2 / 2 - parWidth / 2), -dBodyHeight / 2, -parWidth / 2, _
                -(parWidth / 2 - ConduitOD2 / 2), -dBodyHeight / 2, -parWidth / 2)
            oCurves.Add oLine
            Set oLine = Nothing
            
            Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                -(parWidth / 2 - ConduitOD2 / 2), -dBodyHeight / 2, 0, _
                0, -1, 0, _
                -(parWidth / 2 - ConduitOD2 / 2), -dBodyHeight / 2, -parWidth / 2, _
                -(parWidth / 2 - ConduitOD2 / 2), -dBodyHeight / 2, parWidth / 2)
            oCurves.Add oArc
            Set oArc = Nothing
            
            oStPoint.Set -(parWidth / 2 - ConduitOD2 / 2), -dBodyHeight / 2, parWidth / 2
            Set oLineString = PlaceTrCString(oStPoint, oCurves)
            
            Dim oElbowBody As Object
            Set oElbowBody = oGeomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, oLineString, _
                                0, 1, 0, dBodyHeight, True)
            
            Set oCurves = Nothing
            Set oLineString = Nothing
            
            'Create cover
            Set oLineString = New IngrGeom3D.ComplexString3d
            Set oCurves = New Collection
            Dim oBodyCenter As AutoMath.DPosition  'A location at center of the base of body.
            Set oBodyCenter = New AutoMath.DPosition
            oBodyCenter.Set dBodyLength / 2 - ConduitOD2 / 2, -dBodyHeight / 2, 0
    
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                -(oBodyCenter.x - (parCoverLength - parCoverWidth) / 2), oBodyCenter.y, parCoverWidth / 2, _
                -(oBodyCenter.x + (parCoverLength - parCoverWidth) / 2), oBodyCenter.y, parCoverWidth / 2)
            oCurves.Add oLine
            
            Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                -(oBodyCenter.x + (parCoverLength - parCoverWidth) / 2), oBodyCenter.y, 0, _
                0, -1, 0, _
                -(oBodyCenter.x + (parCoverLength - parCoverWidth) / 2), oBodyCenter.y, parCoverWidth / 2, _
                -(oBodyCenter.x + (parCoverLength - parCoverWidth) / 2), oBodyCenter.y, -parCoverWidth / 2)
            oCurves.Add oArc
            
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                -(oBodyCenter.x + (parCoverLength - parCoverWidth) / 2), oBodyCenter.y, -parCoverWidth / 2, _
                -(oBodyCenter.x - (parCoverLength - parCoverWidth) / 2), oBodyCenter.y, -parCoverWidth / 2)
            oCurves.Add oLine
            Set oLine = Nothing

            Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                -(oBodyCenter.x - (parCoverLength - parCoverWidth) / 2), -dBodyHeight / 2, 0, _
                0, -1, 0, _
                -(oBodyCenter.x - (parCoverLength - parCoverWidth) / 2), oBodyCenter.y, -parCoverWidth / 2, _
                -(oBodyCenter.x - (parCoverLength - parCoverWidth) / 2), oBodyCenter.y, parCoverWidth / 2)
            oCurves.Add oArc
            Set oArc = Nothing
            
            oStPoint.Set -(oBodyCenter.x - (parCoverLength - parCoverWidth) / 2), oBodyCenter.y, _
                                                                                parCoverWidth / 2
            Set oLineString = PlaceTrCString(oStPoint, oCurves)

            Dim oCover As Object
            Set oCover = oGeomFactory.Projections3d.CreateByCurve(m_outputColl.ResourceManager, _
                        oLineString, 0, -1, 0, COVERHEIGHT, True)
            
            Set oCurves = Nothing
            Set oLineString = Nothing
            
            'create cylinder at port 1
            oStPoint.Set 0, 0, 0
            oEnPoint.Set ConduitOD2 / 2 - parLength, 0, 0
            
            Dim oCylinderatPort1 As Object
            Set oCylinderatPort1 = PlaceCylinder(m_outputColl, oStPoint, oEnPoint, ConduitOD1, True)
            
            'Set the output
            m_outputColl.AddOutput "Cylinder at Port 1", oCylinderatPort1
            Set oCylinderatPort1 = Nothing
            
            'create cylinder at port 2
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0, parHeight, 0
            
            Dim oCylinderatPort2 As Object
            Set oCylinderatPort2 = PlaceCylinder(m_outputColl, oStPoint, oEnPoint, ConduitOD1, True)
            
            'Set the output
            m_outputColl.AddOutput "Cylinder at Port 2", oCylinderatPort2
            Set oCylinderatPort2 = Nothing
            
            If lPartDataBasis = 9001 Then
                Dim oTransformationMatrix As IJDT4x4
                Set oTransformationMatrix = New DT4x4
                oTransformationMatrix.LoadIdentity

                Dim oVector As AutoMath.DVector
                Set oVector = New AutoMath.DVector
                oVector.Set 1, 0, 0
                oTransformationMatrix.Rotate (PI / 2), oVector
                oElbowBody.Transform oTransformationMatrix
                oCover.Transform oTransformationMatrix

                Set oVector = Nothing
                Set oTransformationMatrix = Nothing
            End If

            If lPartDataBasis = 9002 Then
                Set oTransformationMatrix = New DT4x4
                oTransformationMatrix.LoadIdentity
                Set oVector = New AutoMath.DVector
                oVector.Set -1, 0, 0
                oTransformationMatrix.Rotate PI / 2, oVector
                oElbowBody.Transform oTransformationMatrix
                oCover.Transform oTransformationMatrix
                
                Set oVector = Nothing
                Set oTransformationMatrix = Nothing
            End If
            
            'Set the output
            iOutput = iOutput + 1
            m_outputColl.AddOutput arrayOfOutputs(iOutput), oElbowBody
            Set oElbowBody = Nothing
            m_outputColl.AddOutput "Cover", oCover
            Set oCover = Nothing
            
          Case Else
                GoTo ErrorLabel
        
    End Select
        
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oGeomFactory = Nothing

    'Place Nozzle 1
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJConduitPortOcc
    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateConduitNozzle(oPlacePoint1, oDir, m_outputColl, oPartFclt, 1)
    
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint1 = Nothing
    
    'Place Nozzle 2
    oDir.Set 0, 1, 0
    Set objNozzle = CreateConduitNozzle(oPlacePoint2, oDir, m_outputColl, oPartFclt, 2)
    
    'Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint2 = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
