Attribute VB_Name = "basCHalfCoupling"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   basCHalfCoupling.bas
'   Author:         MS
'   Creation Date:  Tuesday, Aug 27 2002
'   Description:
'   Symbol Model No. is: F31 Page No. D-48 of PDS Piping Component Data Reference Guide.
'   Symbol is created with Three Outputs
'   The Two physical aspect outputs are created as follows:
'   One ObjNozzle object by using 'CreateNozzle' function and another ObjNozzle by using CreateNozzleWithLength
'   The One Insulation aspect output ObjHalfCouplingIns is created 'PlaceCylinder'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'
'  08.SEP.2006     KKC       DI-95670  Replace names with initials in all revision history sheets and symbols
'  22.OCT.2007     svsmylav  CR-128137 Added 'RetrieveWallThickness' function
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

Public Function RetrieveWallThickness(index As Long, _
                                ByRef partInput As PartFacelets.IJDPart) As Double
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    
    Set oCollection = partInput.GetNozzles()
    
    For pPortIndex = 1 To oCollection.Size
        Set oPipePort = oCollection.Item(pPortIndex)
            If oPipePort.PortIndex = index Then
                RetrieveWallThickness = oPipePort.WallThicknessOrGrooveSetback
                Exit For
            End If
    Next pPortIndex
    Set oCollection = Nothing
    Set oPipePort = Nothing
End Function

