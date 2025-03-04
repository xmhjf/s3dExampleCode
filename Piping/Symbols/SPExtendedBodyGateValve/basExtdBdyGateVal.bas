Attribute VB_Name = "basExtdBdyGateVal"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   basExtndBodyGateVal.bas
'   Author:         MA
'   Creation Date:  Friday, May 16 2008
'   Description:
'   Source: For PDB value 988: Forged Steel Valves Catalog, Bonney Forge, www.bonneyforge.com
'   For PDB value 13: Vogt Catalog
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   16.May.2008     MA      CR-141770 Created the symbol
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Public Type InputType
    name        As String
    description As String
    properties  As IMSDescriptionProperties
    uomValue    As Double
End Type

Public Type OutputType
    name            As String
    description     As String
    properties      As IMSDescriptionProperties
    Aspect          As SymbolRepIds
End Type

Public Type AspectType
    name                As String
    description         As String
    properties          As IMSDescriptionProperties
    AspectId            As SymbolRepIds
End Type


'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub
Public Sub ReportUnanticipatedError(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

'CR-141770 - Function retrieves the wall thickness of the pipe.
Public Function RetrievePipeWallThick(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart, _
                                ByRef dPipeThick As Double)

    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    
    Set oCollection = partInput.GetNozzles()
    For pPortIndex = 1 To oCollection.Size
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = index Then
                dPipeThick = oPipePort.WallThicknessOrGrooveSetback
            End If
    Next pPortIndex
    
    Set oPipePort = Nothing
    
End Function


