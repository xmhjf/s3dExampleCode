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
'       Symbol Model No. is: S28AZ Page No. 347 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'  Change History:
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper


Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    Set m_oGeomHelper = New SymbolServices
    
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
    
    Dim FacetoCenter As Double
    Dim Face1toCenter As Double
    Dim PipingSpecialityLength As Double
    Dim PipingSpecialityLength1 As Double
    Dim PipingSpecialityLength2 As Double
    Dim CylinderDiameter As Double
    Dim PipingSpecialityLength3 As Double
    Dim CylinderDiameter1 As Double

' Inputs
    FacetoCenter = arrayOfInputs(2) 'P2
    Face1toCenter = arrayOfInputs(3) 'P3
    PipingSpecialityLength = arrayOfInputs(4) 'P4
    PipingSpecialityLength1 = arrayOfInputs(5) 'P5
    PipingSpecialityLength2 = arrayOfInputs(6) 'P6
    CylinderDiameter = arrayOfInputs(7) 'P7
    PipingSpecialityLength3 = arrayOfInputs(8) 'P8
    CylinderDiameter1 = arrayOfInputs(9) 'P9

        NozzleDefineDirectionVector oNozzleData, NozzleGeometryType_Linear

        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP1, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
        

        MoveAlongAxis oT4x4Current, AxisDirection_Primary, oNozzleData(1).dCptOffset
        

        Dim spool_1 As Variant
        spool_1 = FacetoCenter - oNozzleData(1).dFlangeThickness - PipingSpecialityLength2 - PipingSpecialityLength

        Dim spool_2 As Variant
        spool_2 = Face1toCenter - oNozzleData(2).dFlangeThickness - PipingSpecialityLength3 - PipingSpecialityLength1

        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = oNozzleData(1).dDepth
'
'        Dim oCylinderCapped1 As Object
'        Set oCylinderCapped1 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(1).dFlangeDiameter), True)
'
'        oCylinderCapped1.Transform oT4x4Current
'
'        oOutputCol.Add oCylinderCapped1
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = spool_1
        
        Dim oCylinderCapped2 As Object
        Set oCylinderCapped2 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(1).dPipeDiameter), True)
        
        oCylinderCapped2.Transform oT4x4Current
        
        oOutputCol.Add oCylinderCapped2
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = PipingSpecialityLength2
        
        Dim oCylinderCapped3 As Object
        Set oCylinderCapped3 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(CylinderDiameter), True)
        
        oCylinderCapped3.Transform oT4x4Current
        
        oOutputCol.Add oCylinderCapped3
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = PipingSpecialityLength
        
        Dim oCylinderCapped4 As Object
        Set oCylinderCapped4 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(1).dPipeDiameter), True)
        
        oCylinderCapped4.Transform oT4x4Current
        
        oOutputCol.Add oCylinderCapped4
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP0, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
        



        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = PipingSpecialityLength1
        
        Dim oCylinderCapped5 As Object
        Set oCylinderCapped5 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(1).dPipeDiameter), True)
        
        oCylinderCapped5.Transform oT4x4Current
        
        oOutputCol.Add oCylinderCapped5
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = PipingSpecialityLength3
        
        Dim oCylinderCapped6 As Object
        Set oCylinderCapped6 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(CylinderDiameter1), True)
        
        oCylinderCapped6.Transform oT4x4Current
        
        oOutputCol.Add oCylinderCapped6
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
        oT4x4Temp.IndexValue(12) = spool_2
        
        Dim oCylinderCapped7 As Object
        Set oCylinderCapped7 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(2).dPipeDiameter), True)
        
        oCylinderCapped7.Transform oT4x4Current
        
        oOutputCol.Add oCylinderCapped7
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        oT4x4Temp.LoadIdentity
'        oT4x4Temp.IndexValue(12) = oNozzleData(2).dDepth
'
'        Dim oCylinderCapped8 As Object
'        Set oCylinderCapped8 = PlaceCylinder(m_OutputColl, oOriginPos, oT4x4Temp.TransformPosition(oOriginPos), CDbl(oNozzleData(2).dFlangeDiameter), True)
'
'        oCylinderCapped8.Transform oT4x4Current
'
'        oOutputCol.Add oCylinderCapped8
        
        oT4x4Current.MultMatrix oT4x4Temp
        

        MoveAlongAxis oT4x4Current, AxisDirection_Primary, oNozzleData(2).dCptOffset
        

        Call NozzlePlace(m_OutputColl, oPart, oNozzleData, NozzleId_CP2, oOriginPos, oT4x4Current, oNozzleCol, oOutputCol)
    
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
        
'        'move the nozzle in the opposite direction by the flange thickness
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
End Sub

