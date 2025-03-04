Attribute VB_Name = "CableTrayType"
'*******************************************************************
'  Copyright (C) 1998-2004, Intergraph Corporation.  All rights reserved.
'
'  File : K:\CommonRoute\Middle\Reports\cblFillCalculations\CableTrayType.bas
'
'  Description:
'   Public Consts for CableTray types
'   Contains a user definded data type for storing the
'   values of NEC Specification of article 392
'
'  Change History:
'  Author : Cable Team
'  Suman       27 Mar 2007 CR-CP 88535   Moved to \Rules so that this project gets delivered in Programming resources of end user build.
'                                        Made changes suitably that in code we donot refer to any client components at all.
'************************************************************************
'General Declarations
Option Explicit

Private Const MODULE = "CableTrayType"

Public Const CABLETRAY_LADDER               As Long = 5
Public Const CABLETRAY_TROUGH_VENTILATED    As Long = 10
Public Const CABLETRAY_TROUGH_SOLID_BOTTOM  As Long = 15
Public Const CABLETRAY_SOLID_BOTTOM         As Long = 20
Public Const CABLETRAY_VENTED_BOTTOM        As Long = 25
Public Const CABLETRAY_CHANNEL_VENTILATED   As Long = 30
Public Const CABLETRAY_CHANNEL_SOLID        As Long = 35
Public Const CABLETRAY_CENTER_SUPPORTED     As Long = 40
Public Const CABLETRAY_WIRE_BASKET          As Long = 45


Public Const CABLE_COMMUNICATION            As Long = 1
Public Const CABLE_CONTROL                  As Long = 2
Public Const CABLE_DATA                     As Long = 3
Public Const CABLE_FIRE_ALARM               As Long = 4
Public Const CABLE_LIGHTNING                As Long = 5
Public Const CABLE_MULTICONDUCTOR_POWER     As Long = 6
Public Const CABLE_POWER                    As Long = 7
Public Const CABLE_SIGNAL                   As Long = 8

Public Const NECTABLESCOUNT                 As Long = 34

Private Type udtNECTableFormat
    strTableName        As String
    dWidth              As Double
    dMaximumAllowedArea As Double
End Type



Private m_colAllTablesCollection(NECTABLESCOUNT) As udtNECTableFormat
Private m_intWidth As Integer

'log errors to the errors collection
Private m_oServerErrors As IJEditErrors
Private m_oServerError As IJEditError

'*********************************************************************
' Function Name :InitializeAllTables
' Description   :Used to initialize m_colAllTablesCollection array to fill
'                it with the specifications of NEC
'
'
'
'
'**********************************************************************

Public Sub InitializeAllTables()
    Const METHOD = "InitializeAllTables"
    On Error GoTo ErrorHandler

    Dim lngCount    As Long
    Dim lngNoofRows As Long

    lngCount = 0

    Dim i           As Integer
    'Table 392.10(A)
    'Table for single conductor cables passing through Ladder
    'or Ventilated trough whose KCMIL is always less than 1000

    lngNoofRows = lngCount + 6

    For i = lngCount To lngNoofRows
        m_colAllTablesCollection(i).strTableName = "TABLE39210A"
    Next i
    
    m_colAllTablesCollection(lngCount).dWidth = 6#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 6.5
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 9#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 9.5
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 12#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 13#
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 18#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 19.5
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 24#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 26#
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 30#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 32.5
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 36#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 39#
    lngCount = lngCount + 1

    '-------====--------------------===-------------------------------------
    'Table 392.9(E)1
    'Table for multi conductor cables passing through
    ' Ventilated channel having only one cable

    lngNoofRows = lngCount + 2

    For i = lngCount To lngNoofRows
        m_colAllTablesCollection(i).strTableName = "TABLE3929E1"
    Next i

    m_colAllTablesCollection(lngCount).dWidth = 3#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 2.3
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 4#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 4.5
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 6#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 7#
    lngCount = lngCount + 1

    '-------====--------------------===-------------------------------------
    'Table 392.9(E)2
    'Table for multi conductor cables passing through
    ' Ventilated channel having more than one cable

    lngNoofRows = lngCount + 2
    '*********************************************************************
    'Following was modified
    'Delivered code commented out
    'For i = 10 To lngNoofRows
    
    'Modified to correctly add table values
    For i = lngCount To lngNoofRows
    'End of modification
    '*********************************************************************
        m_colAllTablesCollection(i).strTableName = "TABLE3929E2"
    Next i

    m_colAllTablesCollection(lngCount).dWidth = 3#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 1.3
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 4#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 2.5
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 6#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 3.8
    lngCount = lngCount + 1
    '-------------------------------------------------------------------------
    'Table 392.9(f)1
    'Table for multi conductor cables passing through
    ' Solid channel having only one cable

    lngNoofRows = lngCount + 3

    For i = lngCount To lngNoofRows
        m_colAllTablesCollection(i).strTableName = "TABLE3929F1"
    Next i


    m_colAllTablesCollection(lngCount).dWidth = 2#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 1.3
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 3#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 2#
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 4#
    '*********************************************************************
    'Following was modified due to an inconsistency with NEC 392.9(f) Column 1 and
    'the delivered example code from Intergraph
    
    'Delivered code commented out
    'm_colAllTablesCollection(lngCount).dMaximumAllowedArea = 7#

    'Code modified to change allowable cable fill to 3.7 in^2 per NEC code
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 3.7
    'End of Modification
    '*********************************************************************
    lngCount = lngCount + 1


    m_colAllTablesCollection(lngCount).dWidth = 6#
    '*********************************************************************
    'Following was modified due to an inconsistency with NEC 392.9(f) Column 1 and
    'the delivered example code from Intergraph
    
    'Delivered code commented out
    'm_colAllTablesCollection(lngCount).dMaximumAllowedArea = 7#

    'Code modified to change allowable cable fill to 5.5 in^2 per NEC code
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 5.5
    'End of Modification
    '*********************************************************************
    lngCount = lngCount + 1


    '-----------------------------------------------------------------------
    'Table 392.9(f)2
    'Table for multi conductor cables passing through
    ' solid channel having more than one cable

    lngNoofRows = lngCount + 3

    For i = lngCount To lngNoofRows
        m_colAllTablesCollection(i).strTableName = "TABLE3929F2"
    Next i


    m_colAllTablesCollection(lngCount).dWidth = 2#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 0.8
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 3#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 1.1
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 4#
    '*********************************************************************
    'Following was modified due to an inconsistency with NEC 392.9(f) Column 2 and
    'the delivered example code from Intergraph
    
    'Delivered code commented out
    'm_colAllTablesCollection(lngCount).dMaximumAllowedArea = 3.8

    'Code modified to change allowable cable fill to 2.1 in^2 per NEC code
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 2.1
    'End of Modification
    '*********************************************************************
    lngCount = lngCount + 1
    
    m_colAllTablesCollection(lngCount).dWidth = 6#
    '*********************************************************************
    'Following was modified due to an inconsistency with NEC 392.9(f) Column 2 and
    'the delivered example code from Intergraph
    
    'Delivered code commented out
    'm_colAllTablesCollection(lngCount).dMaximumAllowedArea = 3.8

    'Code modified to change allowable cable fill to 3.2 in^2 per NEC code
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 3.2
    'End of Modification
    '*********************************************************************
    lngCount = lngCount + 1
    
    '-----------------------------------------------------------------------
    'Table 392.9(A)
    'Table for multi conductor cables passing through Ladder
    'or Ventilated trough or solid bottom whose AWG is more than 4 for some cables

    lngNoofRows = lngCount + 6

    For i = lngCount To lngNoofRows
        m_colAllTablesCollection(i).strTableName = "TABLE3929A"
    Next i

    m_colAllTablesCollection(lngCount).dWidth = 6#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 7#
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 9#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 10.5
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 12#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 14#
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 18#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 21#
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 24#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 28#
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 30#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 35#
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 36#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 42#
    lngCount = lngCount + 1
    
    '------------------------------------------------------------------------
    'Table 392.9(C)
    'Table for multi conductor cables passing through Ladder
    'or Ventilated trough or solid bottom whose AWG is less than 4 for all cables

    '*********************************************************************
    'Following was modified due to an inconsistency with NEC 392.9(c) Column 3
    'Delivered code commented out
    'lngNoofRows = lngCount + 5
    
    'Code modified to include a total of 7 rows
    lngNoofRows = lngCount + 6
    'End of Modification
    '*********************************************************************

    For i = lngCount To lngNoofRows
        m_colAllTablesCollection(i).strTableName = "TABLE3929C"
    Next i


    m_colAllTablesCollection(lngCount).dWidth = 6#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 5.5
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 9#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 8#
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 12#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 11#
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 18#
    '*********************************************************************
    'Following was modified due to an inconsistency with NEC 392.9(c) Column 3
    'Delivered code commented out
    'm_colAllTablesCollection(lngCount).dMaximumAllowedArea = 22#
    
    'Code modified to change allowable cable fill to 16.5 in^2 per NEC code
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 16.5
    'End of Modification
    '*********************************************************************
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 24#
    '*********************************************************************
    'Following was modified due to an inconsistency with NEC 392.9(c) Column 3
    'Delivered code commented out
    'm_colAllTablesCollection(lngCount).dMaximumAllowedArea = 27.5
    
    'Code modified to change allowable cable fill to 22 in^2 per NEC code
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 22#
    'End of Modification
    '*********************************************************************
    lngCount = lngCount + 1

    m_colAllTablesCollection(lngCount).dWidth = 30#
    '*********************************************************************
    'Following was modified due to an inconsistency with NEC 392.9(c) Column 3
    'Delivered code commented out
    'm_colAllTablesCollection(lngCount).dMaximumAllowedArea = 33#

    'Code modified to change allowable cable fill to 27.5 in^2 per NEC code
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 27.5
    lngCount = lngCount + 1
    
    'Added 7th row to include case of 36" tray
    m_colAllTablesCollection(lngCount).dWidth = 36#
    m_colAllTablesCollection(lngCount).dMaximumAllowedArea = 33#
    'End of Modification
    '*********************************************************************

    Exit Sub
ErrorHandler:
    Set m_oServerErrors = New JServerErrors
    m_oServerErrors.Clear
    Set m_oServerError = m_oServerErrors.AddFromErr(Err, "Error processing ", METHOD, MODULE)
    m_oServerError.Raise
    Set m_oServerErrors = Nothing
End Sub

'*********************************************************************
' Function Name :GetAllowedArea
' Description   :Gets the Maximum allowed area with reference to the
'                Table Name and width passes as input
'
'
'Parameters     :input : TableName as string
'               :Output: Double

'
'**********************************************************************

Public Function GetAllowedArea(strTableName As String) As Double
    Dim lngCounter  As Long
    Dim intTmpWidth As Integer
    Const METHOD = "IJDFillCalculations_CalcCableArea"
    On Error GoTo ErrorHandler

    For lngCounter = LBound(m_colAllTablesCollection) To UBound(m_colAllTablesCollection)
        If StrComp(m_colAllTablesCollection(lngCounter).strTableName, strTableName, vbTextCompare) = 0 Then
            intTmpWidth = m_colAllTablesCollection(lngCounter).dWidth
            If (intTmpWidth = m_intWidth) Then
                GetAllowedArea = m_colAllTablesCollection(lngCounter).dMaximumAllowedArea
                If Not (GetAllowedArea = 0) Then
                    Exit For
                End If

            End If
        End If
    Next lngCounter

    If (GetAllowedArea = 0) Then
        GetAllowedArea = m_intWidth
    End If

    Exit Function

ErrorHandler:
    Set m_oServerErrors = New JServerErrors
    m_oServerErrors.Clear
    Set m_oServerError = m_oServerErrors.AddFromErr(Err, "Error processing ", METHOD, MODULE)
    m_oServerError.Raise
    Set m_oServerErrors = Nothing
End Function

'*********************************************************************
' Function Name :SetWidth
' Description   :Sets the width of the cable tray
' Parameters     :input : Width As Integer
'

'
'**********************************************************************

Public Sub SetWidth(Width As Integer)
    m_intWidth = Width
End Sub

