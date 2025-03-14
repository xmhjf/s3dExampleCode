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
'**********************************************************************************
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   File:   SimplePhysical.cls
'   Author: svsmylav
'   Date:   Tuesday, Sep 21 2004
'
'   Description:
'      This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      Symbol Model No. is: I16AZ Page No. 359 of PDS Piping Component Data Reference Guide pcd.pdf.
'  Change History:
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper
Private InchToMtr       As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    Set m_oGeomHelper = New SymbolServices
    InchToMtr = 25.4 / 1000
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs() As Variant, ByRef arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorHandler
    
    Dim oPart As IJDPart
    Dim oPipeComponent As IJDPipeComponent

    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Set oPart = arrayOfInputs(1)
'    Set oPipeComponent = oPart
    
    Dim oNozzleData() As NozzleData
    NozzleInitialize oPart, oNozzleData
    
    Dim oNozzleCol As Collection
    Set oNozzleCol = New Collection
    
    Dim oOutputCol As Collection
    Set oOutputCol = New Collection
    
    Dim oT4x4Current As IJDT4x4
    Set oT4x4Current = New DT4x4
    
    Dim oT4x4Temp As IJDT4x4
    Set oT4x4Temp = New DT4x4
    
    Dim oGeometryFactory As GeometryFactory
    Set oGeometryFactory = New GeometryFactory

    Dim oOriginPos As DPosition
    Set oOriginPos = New DPosition
    oOriginPos.Set 0, 0, 0
    
'**********************************************************************************
'**********************************************************************************
'   BEGIN
'**********************************************************************************
'**********************************************************************************
    
        Dim FacetoCenter As Variant
        Dim Face1toCenter As Variant

        FacetoCenter = arrayOfInputs(2)
        Face1toCenter = arrayOfInputs(3)



        NozzleDefineDirectionVector oNozzleData, NozzleGeometryType_Linear
        

        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP1, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
        

        MoveAlongAxis oT4x4Current, AxisDirection_Primary, oNozzleData(1).dCptOffset
        

        Dim diameter As Variant
        diameter = 0#

        Dim length_1 As Variant
        length_1 = FacetoCenter - oNozzleData(1).dFlangeThickness

        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = oNozzleData(1).dDepth
        
'       Dim oCylinderCapped1 As Object
'       Set oCylinderCapped1 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(1).dFlangeDiameter), True)
        
'       oCylinderCapped1.Transform oT4x4Current
        
'       oOutputCol.Add oCylinderCapped1
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = length_1
        
    Dim stPoint As IJDPosition
    Dim enPoint As IJDPosition
            
    Set stPoint = oOriginPos
    Set enPoint = oT4x4Temp.TransformPosition(oOriginPos)
    enPoint.z = enPoint.z
        Dim oCone1 As Object
        Set oCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, CDbl(oNozzleData(1).dPipeDiameter) / 2, CDbl(diameter) / 2, False)
        
        oCone1.Transform oT4x4Current
        
        oOutputCol.Add oCone1
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP0, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
        



        Dim length_2 As Variant
        length_2 = Face1toCenter - oNozzleData(2).dFlangeThickness

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = length_2
        
        Set stPoint = oT4x4Temp.TransformPosition(oOriginPos)
        Set enPoint = oOriginPos
        enPoint.z = enPoint.z
        Dim oCone2 As Object
        Set oCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, CDbl(oNozzleData(2).dPipeDiameter) / 2, CDbl(diameter) / 2, False)
        
        oCone2.Transform oT4x4Current
        
        oOutputCol.Add oCone2
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = oNozzleData(2).dDepth
        
'       Dim oCylinderCapped2 As Object
'       Set oCylinderCapped2 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(2).dFlangeDiameter), True)
        
'       oCylinderCapped2.Transform oT4x4Current
        
'       oOutputCol.Add oCylinderCapped2
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        MoveAlongAxis oT4x4Current, AxisDirection_Primary, oNozzleData(2).dCptOffset
        

        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP2, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
        

'        If (oPipeComponent.PriSizeNPDUnits = "in") Then
'
'                Dim two As Variant
'                two = 2#
'
'                Dim ten As Variant
'                ten = 10#
'
'        Else
'
'                two = 50#
'
'                ten = 250#
'
'        End If
    RetrieveParameters 1, oPart, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

        Dim dia As Double
        dia = pipeDiam
'        dia = oNozzleData(1).dNpd
'
'
'        If (dia <= two) Then
'
'                Dim factor As Variant
'                factor = 0.4
'
'        Else
'
'                If (dia > two And dia <= ten) Then
'
'                        factor = 0.45
'
'                Else
'
'                        factor = 0.5
'
'                End If
'
'        End If
        
    Dim factor As Double
    If (pipeDiam <= 2 * InchToMtr) Then
            factor = 0.4
    Else
            If (pipeDiam > 2 * InchToMtr And pipeDiam <= 10 * InchToMtr) Then
                    factor = 0.45
            Else
                    If (pipeDiam > 10 * InchToMtr) Then
                            factor = 0.5
                    End If
            End If
    End If

        Dim half_plug_length As Variant
        half_plug_length = dia * 0.5

        Dim cone_diameter As Variant
        cone_diameter = FacetoCenter * factor

        Set oT4x4Current = oNozzleData(NozzleId_CP0).oT4x4.Clone
        

        DefineActiveOrientation oNozzleData, oT4x4Current, oNozzleData(1).AxisDirection_CPSecondary, oNozzleData(1).AxisDirection_CPPrimary
        

        MoveAlongAxis oT4x4Current, AxisDirection_Primary, -half_plug_length
        

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = half_plug_length
        
        Set stPoint = oT4x4Temp.TransformPosition(oOriginPos)
        Set enPoint = oOriginPos
        Dim oCone3 As Object
        enPoint.z = enPoint.z
        Set oCone3 = PlaceCone(m_OutputColl, stPoint, enPoint, CDbl(cone_diameter) / 2, CDbl(diameter) / 2, False)
        
        oCone3.Transform oT4x4Current
        
        oOutputCol.Add oCone3
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = half_plug_length
        
        Set stPoint = oOriginPos
        Set enPoint = oT4x4Temp.TransformPosition(oOriginPos)
        enPoint.z = enPoint.z
        Dim oCone4 As Object
        Set oCone4 = PlaceCone(m_OutputColl, stPoint, enPoint, CDbl(cone_diameter) / 2, CDbl(diameter) / 2, False)
        
        oCone4.Transform oT4x4Current
        
        oOutputCol.Add oCone4
        
        oT4x4Current.MultMatrix oT4x4Temp
        



    
'**********************************************************************************
'**********************************************************************************
'   END
'**********************************************************************************
'**********************************************************************************
    
    Dim oNozzle As IJDNozzle
    Dim lNozzleId As Long
    Dim lNozzleIndex As Long
    
    Dim oNozzlePos As DPosition
    
    Dim oNozzleVec As DVector
    Set oNozzleVec = New DVector
    
    'place the nozzles & add to the output collection
    For lNozzleIndex = 1 To oNozzleCol.Count
        lNozzleId = oNozzleCol.Item(lNozzleIndex)
        
        'move the nozzle in the opposite direction by the flange thickness
'        oNozzleVec.Set -oNozzleData(lNozzleId).oDirectionVector.x, -oNozzleData(lNozzleId).oDirectionVector.y, -oNozzleData(lNozzleId).oDirectionVector.z
'        oNozzleVec.Length = oNozzleData(lNozzleId).dFlangeThickness
'
'        Set oNozzlePos = oNozzleData(lNozzleId).oPlacementPoint.Clone.Offset(oNozzleVec)
        Set oNozzlePos = oNozzleData(lNozzleId).oPlacementPoint.Clone
        ' The nozzle needs to be positioned away by (cptoffset-depth) value in the nozzle direction.
        oNozzlePos.x = oNozzlePos.x + (oNozzleData(lNozzleId).oDirectionVector.x) * (oNozzleData(lNozzleId).dCptOffset - oNozzleData(lNozzleId).dDepth)
        oNozzlePos.y = oNozzlePos.y + (oNozzleData(lNozzleId).oDirectionVector.y) * (oNozzleData(lNozzleId).dCptOffset - oNozzleData(lNozzleId).dDepth)
        oNozzlePos.z = oNozzlePos.z + (oNozzleData(lNozzleId).oDirectionVector.z) * (oNozzleData(lNozzleId).dCptOffset - oNozzleData(lNozzleId).dDepth)
        
        Set oNozzle = CreateNozzle(CInt(lNozzleId), oPart, m_OutputColl, oNozzleData(lNozzleId).oDirectionVector, oNozzlePos)
        
        oOutputCol.Add oNozzle
    Next lNozzleIndex
    
    Dim lOutputIndex As Long
    
    'add all outputs
    For lOutputIndex = 1 To oOutputCol.Count
        m_OutputColl.AddOutput arrayOfOutputs(lOutputIndex), oOutputCol.Item(lOutputIndex)
    Next lOutputIndex
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
'    Resume Next
End Sub

