Attribute VB_Name = "PartClassHelper"
'*******************************************************************
'  Copyright (C) 2002 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    PartClassHelper.bas
'
'  Initial Creation : Thandur Raghuveer
'
'  History:
'******************************************************************

Option Explicit

Public Const PARTCLASS = "PartClass"

Public Function GetClassPropertyValue(ByVal oSpacePart As IJDPart, MODULE As String) As String
Const METHOD = "GetClassPropertyValue"
On Error GoTo ErrorHandler
    
    Dim oPartClass                  As IJDPartClass
    Dim oPartRelationHelper         As IMSRelation.DRelationHelper
    Dim oPartCollection             As IMSRelation.DCollectionHelper
    Dim strClass                    As String
    
    If oSpacePart Is Nothing Then
        GetClassPropertyValue = vbNullString
    Else
        If oSpacePart.PartDescription <> vbNullString Then
        
            'Get the Relation Helper
            Set oPartRelationHelper = oSpacePart
            
            'get the partclass
            Set oPartCollection = oPartRelationHelper.CollectionRelations(IID_IJDPart, PARTCLASS)
            
            Set oPartClass = oPartCollection.Item(1)
        End If
        
        
        If Not oPartClass Is Nothing Then
            'Compare the part class type and localize it
            Select Case oPartClass.PartClassType
                Case "ShipZoneClass"
                    strClass = "ShipZone"
                Case "VoidSpaceClass"
                    strClass = "VoidSpace"
                Case "CompartmentClass"
                    strClass = "Compartment"
                Case "CompartRegionClass"
                    strClass = "CompartRegion"
                Case "CompartIFCZoneClass"
                    strClass = "CompartIFCZone"
            End Select
            GetClassPropertyValue = strClass
            
        End If
    End If

    Set oPartClass = Nothing
    Set oPartCollection = Nothing
    Set oPartRelationHelper = Nothing
    
Exit Function
ErrorHandler:
     Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function
Public Property Get IID_IJDPart() As Variant
    IID_IJDPart = InitGuid(&H7D, &H1B, &HB1, &H9A, &H98, &H92, &HD1, &H11, &HBD, &HDC, &H0, &H60, &H97, &H3D, &H48, &H5)
End Property

'This InitGuid is different in ordering of bytes from the function InitGuid
'   of CommonApp Utilities.
Private Function InitGuid(a As Byte, b As Byte, c As Byte, d As Byte, _
                e As Byte, f As Byte, g As Byte, h As Byte, _
                i As Byte, j As Byte, k As Byte, l As Byte, _
                m As Byte, n As Byte, o As Byte, p As Byte) As Variant

    Dim Guid(0 To 15) As Byte
    
    Guid(0) = a
    Guid(1) = b
    Guid(2) = c
    Guid(3) = d
    Guid(4) = e
    Guid(5) = f
    Guid(6) = g
    Guid(7) = h
    Guid(8) = i
    Guid(9) = j
    Guid(10) = k
    Guid(11) = l
    Guid(12) = m
    Guid(13) = n
    Guid(14) = o
    Guid(15) = p
    
    InitGuid = Guid
End Function


