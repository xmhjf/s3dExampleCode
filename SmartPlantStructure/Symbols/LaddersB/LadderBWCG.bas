Attribute VB_Name = "LadderBWCG"
Option Explicit

'*******************************************************************
'  Copyright (C) 1998, Intergraph Corporation.  All rights reserved.
'
'  File:  LadderBWCG.bas
'
'  Description: Weight and Center of Gravity calculator.  This uses the output objects
'       from the symbol to make the calculations.  This assumes that all out put objects
'       are projections, although that is not necessary.
'       The Paint Area is also calculated.  Paint area does not take into account the end of a projection. (capping surface)
'       The Weight, CG, and Paint Area are all slightly off due to the fact that the steps protrude through the side frame.  (vertical ladder)
'
'  History:
'           03/08/2000 - Jeff Machamer - Creation
'           03/21/2000 - Jeff Machamer - Added calculation of Paint Area (Surface Area)
'  09/12/2007 C.C.P. Created for CR113595 Ladder on Corrugated Bulkhead. Copied from GenericWCG.bas.
'  03/26/2009 - GG - DM#162555,162556, and 162557. Fixed the calculations for weight and COG
'  09/11/2009 - RP - TR168827 changed the way outputs are retreived for computing weight and CG
'******************************************************************
Private Const E_FAIL = -2147467259

Public Sub CalcWCG(oLadder As ISPSLadder, ByVal PartInfoCol As IJDInfosCol, _
                   ByRef weight As Double, ByRef COGX As Double, ByRef COGY As Double, ByRef COGZ As Double)
Const METHOD = "CalcWCG"
On Error GoTo ErrHandler
Dim oErrors As IJEditErrors

 
    Dim iMaterial As IJDMaterial
    Dim material As Variant, grade As Variant, density As Variant
    Dim oSymbol As IJDSymbol
    Dim oOutputProxy As IJDProxy
    Dim oOutputs As IJDTargetObjectCol
    Dim Count As Integer
    Dim oProj As Projection3d
    Dim oCurve As IJCurve
    Dim curveArea As Double
    Dim curveSA As Double
    Dim projVector As New DVector
    Dim projCG As New DPosition
    Dim projVolume As Double
    Dim projSA As Double
    Dim AccumCG As New DPosition
    Dim AccumVolume As Double
    Dim projVecx As Double, projVecy As Double, projVecz As Double
    Dim projVecLength As Double
    Dim oTargetCol As IJDTargetObjectCol
    Dim PartAttrs As IJDAttributes
    Dim OccAttrs As IJDAttributes
    Dim oPartOcc As IJPartOcc
    Dim oPart As IJDPart
    Dim iSymbolOccMisc As iSymbolOccMisc
    Dim oRelation As IJDAssocRelation
    Dim a As Integer
    Dim x As Double, y As Double, z As Double



    Set oErrors = New IMSErrorLog.JServerErrors
    Set oSymbol = oLadder
    Set iSymbolOccMisc = oLadder


    Set oOutputProxy = iSymbolOccMisc.BindToOCIfExists("Physical")
    
    If oOutputProxy Is Nothing Then GoTo SymbolNotComputedYet ' the wiegtCG sematic computed before the symbol, exit now
    'we will be called again after the symbol compute
    Set oRelation = oOutputProxy.Source

    Set oOutputs = oRelation.CollectionRelations(strIJDOutputCollectionUUID, strOutputName)
    Count = oOutputs.Count
    On Error Resume Next
    
    For a = 1 To Count
        Set oProj = oOutputs.Item(a)
        If Not oProj Is Nothing Then
            oProj.GetProjection projVecx, projVecy, projVecz
            Set oCurve = oProj.Curve
                    
            curveArea = oCurve.Area
            curveSA = oCurve.Length
            oCurve.Centroid x, y, z
              
            projCG.x = x + (0.5 * projVecx)
            projCG.y = y + (0.5 * projVecy)
            projCG.z = z + (0.5 * projVecz)
            projVecLength = oProj.Length
            projVolume = projVecLength * curveArea
            AccumVolume = AccumVolume + projVolume
            If AccumVolume > 0 Then
                AccumCG.x = AccumCG.x + (projCG.x - AccumCG.x) * projVolume / AccumVolume
                AccumCG.y = AccumCG.y + (projCG.y - AccumCG.y) * projVolume / AccumVolume
                AccumCG.z = AccumCG.z + (projCG.z - AccumCG.z) * projVolume / AccumVolume
            End If
        Else
            'something besides projection..
        End If
        Set oProj = Nothing
    
    Next
    On Error GoTo ErrHandler
    'Custom ladders could calculate different density for different output objects.
    'Maybe change the inclined ladder "step" to be more geometrically correct so that the weight will be closer to correct.
    'Need to subtract any overlap from the Volume.  (CG should be okay since we would be subtracting the same amount from each side of the CG)
    '  Shouldn't be done in this step since you would have to move to many geometry calculations in here. (would be easy for this ladder though)
    '  Perhaps Change the macro so that there is no overlap?
    
    Set oPartOcc = oLadder
    oPartOcc.GetPart oPart
    Set PartAttrs = oPart
    Set OccAttrs = oLadder
    material = GetAttribute1(OccAttrs, "Primary_SPSMaterial", PartInfoCol)
    grade = GetAttribute1(OccAttrs, "Primary_SPSGrade", PartInfoCol)
     
    Set iMaterial = GetMaterialObject(material, grade)
  
    
    If Not iMaterial Is Nothing Then
        density = iMaterial.density
    Else
        GoTo ErrHandler
    End If
    weight = AccumVolume * density
    COGX = AccumCG.x
    COGY = AccumCG.y
    COGZ = AccumCG.z

SymbolNotComputedYet:

    Exit Sub
ErrHandler:
    Err.Raise E_FAIL
    oErrors.Add Err.Number, METHOD, Err.Description

End Sub
