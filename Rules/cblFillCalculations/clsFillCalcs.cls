VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFillCalcs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************
'Copyright (C) 2002, Intergraph Corporation. All rights reserved.                                     '
'
'File: K:\CommonRoute\Middle\Reports\cblFillCalculations\clsFillCalcs.cls
'Abstract:                                                                                            '
'    Command to generate reports for cable fill in conduit and cable trays                                                              '
'                                                                                                     '
'                                                                                      '
'                                                                                                   '
'                                                                                                     '
'Notes:                                                                                               '
'                                                                                                     '
'History
'Creation:'Cable Team
'
'Modification :
'Kalicharan  02nd Sep 2004  MaximumFill calculation for Cabletrays
'Kalicharan  22nd Feb 2004 TR-CP·70372  Cable Cannot be Routed Through Conduit if Fill Efficiency Set
'Kalicharan  26th Apr 2005 TR-CP·75819  Cable fill calculation reports error, when the cable run is deleted
'Kalicharan  19th Apr 2005 TR-CP·78054  Routing from Conduit branch EF causes errors under certain conditions.
'RamPrasad   09th Jun 2005 TR-CP.79345  Fill Status of Overfilled is given anytime a Percent fill is defined in a Report.
'RamPrasad   09th Jun 2005 TR-CP.79540  Cable Fill by Cableway element report reports wrong Traverse areas.
'RamPrasad   04th Aug 2005 TR-CP.82800  Reports display the wrong Cable Fill for a Cable that is routed through Conduit.
'Kali        9 Aug 2006 Kali TR-CP·103168  Memory Leak When Auto-Routing Cable
'Dheeraj     18 Oct 2006 TR-CP·102081  Modified Fill Calculations to use Rated Voltage vs. Operating Voltage
'Kali        13 Dec 2006 TR-CP·111890  Typemismatch app error after choosing via feature
'Suman       27 Mar 2007 CR-CP 88535   Moved to \Rules so that this project gets delivered in Programming resources of end user build.
'                                      Made changes suitably that in code we donot refer to any client components at all.
'Kali        21 Sep 2007 TR-CP·128014    The FillCalculations program does not use the Load Witdth/Depth
'Vishal      15 Oct 2013  TR-CP-241550  Fill Percentage on Cableway Feature Not Updating Correctly
'*****************************************************************************************************

Option Explicit
Private Const MODULE = "CableFillCalculation"

Private Const PI            As Double = 3.14159265358979
'====================================================================================================
' constants no longer used to determine conductor size
'Private Const KCMIL         As Double = 0.78539816 * 10 ^ -9    'http://cipco.apogee.net/foe/fbwmcm.asp
'Private Const AWG4Diameter  As Double = 0.232    'http://www.hardwarebook.net/table/AWG.html
'====================================================================================================
Private Const TABLE39210A   As String = "TABLE39210A"
Private Const TABLE3929E1   As String = "TABLE3929E1"
Private Const TABLE3929E2   As String = "TABLE3929E2"
Private Const TABLE3929F1   As String = "TABLE3929F1"
Private Const TABLE3929F2   As String = "TABLE3929F2"
Private Const TABLE3929A    As String = "TABLE3929A"
Private Const TABLE3929C    As String = "TABLE3929C"

'Private Const IID_IJRtePathFeat As Variant = "{845E9741-9D93-11D1-BE27-08003651A403}"
Private Const IID_IJRteSemanticFillInfo = "{ACE25E75-EA28-4916-827F-5167792293A1}"

    
Implements IJDFillCalculations
Implements IJRteCARRealTimeFill

Private m_blnKCMIL1000more                 As Boolean   'Flag to check if the cable area is more than 1000 KCMIL
Private m_blnKCMIL250TO1000                As Boolean   'Flag to check if the cable area is betwn 250 to 1000 KCMIL
Private m_blnKCMILLess250                  As Boolean   'Flag to check if the cable area is less than 250 KCMIL
Private m_blnAWGGreater4                   As Boolean   'Flag to check if the cable diameter is more than 4 AWG
Private m_blnAWGLess4                      As Boolean   'Flag to check if the cable diameter is less than 4 AWG
Private m_blnMulticonductor                As Boolean   'Flag to check if the cable is a multi conductor
Private m_blnSingleconductor               As Boolean   'Flag to check if the cable is a single conductor

Private m_intCableCount                    As Integer   'Total number of cable inside the Cabletray
Private m_intCableTray                     As Integer   'Indicates CabelTray type
Private m_intCableCountMultiCon            As Integer   'Total number of multi conductor cables inside the Cabletray

Private m_dblCableDiaSum                   As Double    'Sum of diameters of all cables inside the tray
Private m_dblCableTrayLoadWidth            As Double    'Width of the cable tray
Private m_dblCableTrayLoadDepth            As Double    'Depth of the cable tray
Private m_dblCableTrayNominalWidth         As Double    'Width of the cable tray
Private m_dblCableTrayNominalDepth         As Double    'Depth of the cable tray
Private m_dblCableTrayFillEff              As Double    'Fill efficiency
Private m_dblTotalWireArea                 As Double    'Total area of the wire
Private m_dblCabDia1000                    As Double    'Sum of diameter of cables more than 1000KCMIL
Private m_dblCabArealess1000               As Double    'Sum of area of cables less than 1000KCMIL
Private m_dblCabDia4AWG                    As Double    'Sum of diameter of cables more than 4 AWG
Private m_dblCabArealess4awg               As Double    'Sum of area of cables less than 4 AWG
Private m_dblAllowedArea                   As Double
Private m_dblRatedVoltage                  As Double
'====================================================================================================
' for the purpose of splitting area calculations between single and multiconductor
'cable to allow for NEC 392.11(c) calculations
Private m_bln1to4ought                     As Boolean   'Flag to check if single conductor cable is between 1/0 and 4/0AWG

Private m_dblMultiTotalWireArea            As Double    'Total area of multiconductor cables
Private m_dblSingleTotalWireArea           As Double    'Total area of single conductor cables
Private m_dblMultiCableDiaSum              As Double    'Total width of multiconductor cables
Private m_dblSingleCableDiaSum             As Double    'Total width of single conductor cables
Private m_dblCabArea250to900               As Double    'Sum of area of cables between 250KCMIL and 900KCMIL (NEC 392.10(2))
Private m_dblSizeOfConductors              As Double    'More accurately determine conductor size for NEC calculations
'====================================================================================================

Private m_intSignalType                    As Integer

Private m_bCommunicationCables              As Boolean
Private m_bControl                          As Boolean
Private m_bDataCables                       As Boolean
Private m_bFireAlarmCables                  As Boolean
Private m_bLightningCables                  As Boolean
Private m_bMultiConPowerCables              As Boolean
Private m_bPowerCables                      As Boolean
Private m_bSignalCables                     As Boolean
Private m_bRatedVoltage                     As Boolean
Private m_UOM                               As New UnitsOfMeasureServicesLib.UomVBInterface   ' To use UOM service
Private m_CableDef                          As IJCableDefinition
Private m_bIsCaseOfBarrier                  As Boolean

' Logging
Private m_strLogFilePath                    As String
Private m_oLogFile                          As FileSystemObject
Private m_oLogFileStream                    As TextStream
Private m_strCablewayName                   As String

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'log errors to the errors collection
Private m_oServerErrors                     As IJEditErrors
Private m_oServerError                      As IJEditError

Private Sub Class_Initialize()
    Set m_oServerErrors = New JServerErrors
    m_oServerErrors.Clear
End Sub

Private Sub Class_Terminate()
    Set m_oServerErrors = Nothing
End Sub

Private Function IJDFillCalculations_CalcCableArea(ByVal pDispObject As Object) As Double

    Const METHOD = "IJDFillCalculations_CalcCableArea"

    On Error GoTo ErrorHandler

    Dim oCable           As clsCable
    Dim oCableData       As clsCableData
    Dim colCables        As IJElements
    Dim dblTotalWireArea As Double

    Set oCableData = New clsCableData
    Set colCables = pDispObject

    If Not colCables Is Nothing Then
        If Not colCables.Count = 0 Then
            Set oCableData.CableRuns = colCables
            For Each oCable In oCableData.CableData
                dblTotalWireArea = dblTotalWireArea + PI * (oCable.CableDiameter / 2) ^ 2
            Next oCable
        Else
            dblTotalWireArea = 0#
        End If
    Else
        dblTotalWireArea = 0#
    End If

    IJDFillCalculations_CalcCableArea = dblTotalWireArea

    Exit Function

ErrorHandler:
    Set m_oServerError = m_oServerErrors.AddFromErr(Err, "Error processing ", METHOD, MODULE)
    m_oServerError.Raise
End Function

Private Sub IJDFillCalculations_GetConduitFillParams(ByVal pDispConduitFeatObject As Object, ByVal pDispConduitGenPart As Object, lNoCables As Long, dWireArea As Double, bTradeSize As String, dTotalTraverseArea As Double, dAllowableTraverseArea As Double, dPercentFull As Double, dAvailableTraverseArea As Double, bStatus As String)
    Const METHOD = "IJDFillCalculations_GetConduitFillParams"
    On Error GoTo ErrorHandler

    Dim oConduitFeat    As IJRteConduitPathFeat
    Dim oPathFeat       As IJRtePathFeat
    Dim oParent         As IJDesignParent
    Dim oChild          As IJDesignChild
    Dim oConduitRun     As IJRteConduitRun
    Dim oGenPart        As IJRtePathGenPart
    Dim oPartOcc        As IJDistribPartOccur
    Dim oPortCollection As IJElements
    Dim oPort           As IJDistribPort
    Dim oConduitPort    As IJConduitPortOcc

    Dim dblFillEff          As Double
    Dim strNCD              As String
    Dim strNCDUnits         As String
    Dim strConduitTradeSize As String
    Dim dblOutsideDiameter  As Double
    Dim dblWallThickness    As Double
    Dim dblTotalArea        As Double
    Dim dblAvailArea        As Double
    Dim dblAllowArea        As Double
    Dim strFillStatus       As String
    Dim dblPercFull         As Double
    Dim oConduitPathFeat    As IJRtePathFeat
    Dim lCount              As Long
    Dim oElements           As IJElements

    On Error GoTo ErrorHandler
    
    dWireArea = 0#
    Set oGenPart = Nothing
    Set oConduitFeat = pDispConduitFeatObject
    ' The EndPathFeat and BranchFeature does not have a PartOccurence hence do not calculate the fill
    If (TypeOf pDispConduitFeatObject Is IJRteEndPathFeat) Or (TypeOf pDispConduitFeatObject Is IJRteBranchPathFeat) Then
        Exit Sub
    End If

    ' get the part from the feature
    If oGenPart Is Nothing Then
        If Not oConduitFeat Is Nothing Then
            Set oConduitPathFeat = oConduitFeat
            Set oElements = oConduitPathFeat.GetBaseParts(lCount)
            If lCount > 0 Then
                Set oGenPart = oElements.Item(1)
                If Not oGenPart Is Nothing Then

                End If
            End If
        End If
    End If

    Set oPathFeat = pDispConduitFeatObject
    Set oChild = oPathFeat
    Set oParent = oChild.GetParent
    Set oConduitRun = oParent


    '        later it needs to get the fill efficiency from
    dblFillEff = GetFillEfficiencyFromFeature(oConduitRun)

    If dblFillEff = 0 Then
        bStatus = "Cables Not Allowed"
        Exit Sub
    End If


    strNCD = oConduitRun.Ncd
    strNCDUnits = oConduitRun.NcdUnitType

    strConduitTradeSize = CStr(strNCD) & " " & strNCDUnits

    If oGenPart Is Nothing Then
        Set oGenPart = pDispConduitGenPart
    End If

    Set oPartOcc = oGenPart
    If oPartOcc Is Nothing Then
        Exit Sub
    End If

    '--- Retrieve Ports From Part Occurrence In Order to Get Diameter and WallThickness

    Set oPortCollection = oPartOcc.GetPorts(DistribPortStatus_BASE, DistribPortType_CONDUIT)

    If Not oPortCollection Is Nothing Then
        For Each oPort In oPortCollection
            Set oConduitPort = oPort

            dblOutsideDiameter = oConduitPort.ConduitDiameter
            dblWallThickness = oConduitPort.WallThickness
            dblTotalArea = PI * ((dblOutsideDiameter - 2 * dblWallThickness) / 2) ^ 2

        Next oPort
    End If

    Set oConduitPort = Nothing
    Set oPort = Nothing
    Set oPortCollection = Nothing
    Set oPartOcc = Nothing
    Set oGenPart = Nothing
    Set oParent = Nothing
    Set oChild = Nothing
    Set oPathFeat = Nothing
    Set oConduitFeat = Nothing

    Dim oIJDAssocRelation As IJDAssocRelation
    Dim oTargetObjCol As IJDTargetObjectCol
    Dim nLoop As Integer
    Dim oSegmentPathFeat As IJRtePathFeat
    Dim oCableRun As IJRteCableRun
    Dim OCablePart As IJCablePart



    Set oIJDAssocRelation = pDispConduitFeatObject
    'Get the collection of segments in relation with the Cabletray
    Set oTargetObjCol = oIJDAssocRelation.CollectionRelations(IID_IJRteSemanticFillInfo, "Segment")
    lNoCables = oTargetObjCol.Count


    For nLoop = 1 To lNoCables
        Set oSegmentPathFeat = Nothing
        Set oCableRun = Nothing
        Set OCablePart = Nothing
        Set oSegmentPathFeat = oTargetObjCol.Item(nLoop)
        Set oCableRun = oSegmentPathFeat.GetPathRun
        Set OCablePart = oCableRun.GetSpecificCablePart
        dWireArea = dWireArea + PI * (OCablePart.CableDiameter / 2) ^ 2
    Next nLoop
    'if there is a non-commited cable path passing through then include the realtime fill
    If Not m_CableDef Is Nothing Then
        Dim oUnk As IMoniker
        Dim oPOM As IJDPOM  'To get Object from Moniker
        Set oCableRun = Nothing
        Set OCablePart = Nothing
        Set oUnk = m_CableDef.CableMoniker
        Set oPOM = GetPOM("model")
        Set oCableRun = oPOM.GetObject(oUnk)
        Set oPOM = Nothing
        Set OCablePart = oCableRun.GetSpecificCablePart
        dWireArea = dWireArea + PI * (OCablePart.CableDiameter / 2) ^ 2
    End If

    Set oIJDAssocRelation = Nothing
    Set oTargetObjCol = Nothing
    Set oSegmentPathFeat = Nothing
    Set oCableRun = Nothing
    Set OCablePart = Nothing
    '--- Base Information Has Been Retrieved - Perform Calculations

    Select Case lNoCables

    Case 0
        dblAllowArea = dblTotalArea * 0.4 * dblFillEff
    Case 1
        dblAllowArea = dblTotalArea * 0.53 * dblFillEff
    Case 2
        dblAllowArea = dblTotalArea * 0.31 * dblFillEff
    Case Else
        dblAllowArea = dblTotalArea * 0.4 * dblFillEff
    End Select

    If dblAllowArea > 0 Then
        dblPercFull = dWireArea / dblAllowArea
    Else
        dblPercFull = 0#
    End If

    dblAvailArea = dblAllowArea - dWireArea

    Select Case dblAvailArea
    Case Is = 0
        strFillStatus = "Full"
    Case Is > 0
        strFillStatus = "Partial"
    Case Is < 0
        strFillStatus = "OverFilled"
    End Select

    '--- Populate Return Arguments

    bTradeSize = strConduitTradeSize
    dTotalTraverseArea = dblTotalArea
    dPercentFull = dblPercFull

    '--- Don't Display a Negative Available Area

    If dblAvailArea < 0 Then
        dAvailableTraverseArea = 0#
    Else
        dAvailableTraverseArea = dblAvailArea
    End If
    dAllowableTraverseArea = dblAllowArea
    bStatus = strFillStatus

    Exit Sub

ErrorHandler:
    Set m_oServerError = m_oServerErrors.AddFromErr(Err, "Error processing ", METHOD, MODULE)
    m_oServerError.Raise
End Sub

Private Sub IJDFillCalculations_GetCwayFillParams(ByVal pDispCableWayFeatObject As Object, ByVal pDispCableWayGenPart As Object, lNoCables As Long, dWireArea As Double, bTradeSize As String, dTotalTraverseArea As Double, dAllowableTraverseArea As Double, dPercentFull As Double, dAvailableTraverseArea As Double, bStatus As String)

    Const METHOD = "IJDFillCalculations_GetCwayFillParams"

    On Error GoTo ErrorHandler

    Dim oCableWayFeat       As IJRteCablewayPathFeat
    Dim oParent             As IJDesignParent
    Dim oChild              As IJDesignChild
    Dim oCableWayPart       As IJCableTrayPart
    Dim oCableTrayPartObj   As IJPartOcc


    Dim oIJDAssocRelation   As IJDAssocRelation
    Dim oTargetObjCol       As IJDTargetObjectCol
    Dim nLoop               As Integer
    Dim oSegmentPathFeat    As IJRtePathFeat
    Dim oCableRun           As IJRteCableRun
    Dim OCablePart          As IJCablePart

    Dim oCableway           As IJRteCableway
    Dim oCbleTrayFeat       As IJRtePathFeat
    Dim oCbleWayFeat        As IJRtePathFeat
    Dim oCblwyPartCol       As IJDObjectCollection
    Dim oCblwyRun           As IJRtePathRun
    Dim oTempPart           As IJDPart
    Dim oCblPartObj         As Object
    Dim dTempWidth          As Double
    Dim dTempDepth          As Double
    Dim dblAvailArea        As Double
    Dim strFillStatus       As String
    Dim oCbleWayCrossSect  As IJRteCrossSectOccur
    Dim oCrossSecShape As CrossSectionShapeTypes
    Dim dWidth As Double
    Dim dDepth As Double
    Dim dCornerRad As Double
    Dim oObject As Object
    Dim oNamed As IJNamedItem

    On Error GoTo ErrorHandler
    
    strFillStatus = "Not Specified"
    dWireArea = 0#
    
    m_strLogFilePath = GetSystemTempPath
    m_strLogFilePath = m_strLogFilePath & "CableFillCalculationRule.log"
    
    Set m_oLogFile = CreateObject("Scripting.FileSystemObject")
    
    Dim bLogFileExists As Boolean
    bLogFileExists = m_oLogFile.FileExists(m_strLogFilePath)
    If bLogFileExists Then
        Set m_oLogFileStream = m_oLogFile.OpenTextFile(m_strLogFilePath, ForAppending, True, TristateUseDefault)
    Else
        Set m_oLogFileStream = m_oLogFile.CreateTextFile(m_strLogFilePath)
    End If
    
    InitializeAllTables

    Set oCableWayFeat = pDispCableWayFeatObject
    Set oCbleTrayFeat = pDispCableWayFeatObject
    Set oCbleWayFeat = pDispCableWayFeatObject
    Set oCblwyRun = oCbleTrayFeat.GetPathRun
    
    Set oNamed = oCblwyRun
    m_strCablewayName = oNamed.Name
    Set oNamed = Nothing
    
    Set oChild = oCbleTrayFeat
    Set oParent = oChild.GetParent
    Set oCableway = oParent
    
    m_dblCableTrayFillEff = GetFillEfficiencyFromFeature(oCableway)
    If m_dblCableTrayFillEff = 0 Then
        bStatus = "Cables Not Allowed"
        Exit Sub
    End If
    
    CheckIfIsBarrierAndDeriveTrayType oCbleWayFeat
    
    If m_bIsCaseOfBarrier = True Then
                
        Set oCbleWayCrossSect = oCbleWayFeat
        oCbleWayCrossSect.GetParameters oCrossSecShape, dWidth, dDepth, dCornerRad
        
        m_dblCableTrayNominalWidth = m_UOM.ConvertDbuToUnit(UNIT_DISTANCE, dWidth, DISTANCE_INCH)
        m_dblCableTrayNominalDepth = m_UOM.ConvertDbuToUnit(UNIT_DISTANCE, dDepth, DISTANCE_INCH)
        
        'For barrier nominal width and depth are taken as load width and depth respectively.
        m_dblCableTrayLoadWidth = m_dblCableTrayNominalWidth
        m_dblCableTrayLoadDepth = m_dblCableTrayNominalDepth
        
    Else
        
        Dim iC As Long
        Dim oCableElem As IJElements
        Dim oCrossUtil As IJRtePathCrossSectUtility
        Dim bOuterDim As Boolean
        Dim bHasCSShape As Boolean
        
        oCbleWayFeat.GetParts iC, oCableElem
        
        If iC >= 1 Then
        
            For Each oCblPartObj In oCableElem
                On Error Resume Next
                Set oCableWayPart = Nothing
                Set oCableTrayPartObj = oCblPartObj
                If Not oCableTrayPartObj Is Nothing Then
                    oCableTrayPartObj.GetPart oTempPart
                    Set oCableWayPart = oTempPart
                    Exit For
                End If
            Next oCblPartObj
            
            If Not oCableWayPart Is Nothing Then
            m_intCableTray = oCableWayPart.TrayType
            dTempWidth = oCableWayPart.NominalWidth
            dTempDepth = oCableWayPart.NominalDepth
            Else
                Set oCrossUtil = oCbleWayFeat
                oCrossUtil.GetCrossSectionDataForFeature oCbleWayFeat, oCrossSecShape, dTempWidth, dTempDepth, dCornerRad, bOuterDim, bHasCSShape
            End If
    
            m_dblCableTrayNominalWidth = dTempWidth
            m_dblCableTrayNominalDepth = dTempDepth
    
            m_dblCableTrayLoadWidth = 0
            m_dblCableTrayLoadDepth = 0
            
            GetLoadWidthAndDepth pDispCableWayFeatObject
            'Dont have LoadWidth and LoadDepth values.Use NominalWidth and Depth values instead.
            If (m_dblCableTrayLoadWidth = 0) Or (m_dblCableTrayLoadDepth = 0) Then
                m_dblCableTrayLoadWidth = dTempWidth
                m_dblCableTrayLoadDepth = dTempDepth
            End If
    
            Dim dTmpInchWidth As Double
            Dim dTmpInchDepth As Double
    
            m_dblCableTrayNominalWidth = m_UOM.ConvertDbuToUnit(UNIT_DISTANCE, m_dblCableTrayNominalWidth, DISTANCE_INCH)
            m_dblCableTrayNominalDepth = m_UOM.ConvertDbuToUnit(UNIT_DISTANCE, m_dblCableTrayNominalDepth, DISTANCE_INCH)
            
            m_dblCableTrayLoadWidth = m_UOM.ConvertDbuToUnit(UNIT_DISTANCE, m_dblCableTrayLoadWidth, DISTANCE_INCH)
            m_dblCableTrayLoadDepth = m_UOM.ConvertDbuToUnit(UNIT_DISTANCE, m_dblCableTrayLoadDepth, DISTANCE_INCH)
            
            Set oCableElem = Nothing
        
        End If
    End If
    
        dTotalTraverseArea = m_dblCableTrayNominalWidth * m_dblCableTrayNominalDepth
        dTotalTraverseArea = m_UOM.ConvertUnitToDbu(UNIT_AREA, dTotalTraverseArea, AREA_SQUARE_INCH)
        
        SetWidth (m_dblCableTrayNominalWidth)
        
        Set oIJDAssocRelation = oCableWayFeat
        'Get the collection of segments in relation with the Cabletray
        Set oTargetObjCol = oIJDAssocRelation.CollectionRelations(IID_IJRteSemanticFillInfo, "Segment")
        m_intCableCount = oTargetObjCol.Count
        lNoCables = m_intCableCount
        m_intCableCountMultiCon = 0

        For nLoop = 1 To m_intCableCount
            Set oSegmentPathFeat = Nothing
            Set oCableRun = Nothing
            Set OCablePart = Nothing
            Set oSegmentPathFeat = oTargetObjCol.Item(nLoop)
            Set oCableRun = oSegmentPathFeat.GetPathRun
            m_intSignalType = oCableRun.SignalType
            Set OCablePart = oCableRun.GetSpecificCablePart
            If Not OCablePart Is Nothing Then
                ProcessCableProperties OCablePart
            End If
        Next nLoop
        
        Dim iCableCountSingleCon As Integer
        iCableCountSingleCon = m_intCableCount - m_intCableCountMultiCon

        If Not m_CableDef Is Nothing Then
            Dim oUnk As IMoniker
            Dim oPOM As IJDPOM     'To get Object from Moniker
            Set oCableRun = Nothing
            Set OCablePart = Nothing
            Set oUnk = m_CableDef.CableMoniker
            Set oPOM = GetPOM("model")
            Set oCableRun = oPOM.GetObject(oUnk)
            Set oPOM = Nothing
            m_intSignalType = oCableRun.SignalType
            Set OCablePart = oCableRun.GetSpecificCablePart
            If Not OCablePart Is Nothing Then
                ProcessCableProperties OCablePart
            End If
        End If
        
        dPercentFull = CalculateMaximumFill
        
        If dPercentFull < 0 Then
            m_oLogFileStream.WriteLine " "
            m_oLogFileStream.WriteLine " Cable fill can not be calculated by S3D for the given CableTray width, because the Maximum Allowable Fill Area calculated through NEC Table is a negative value. "
            m_oLogFileStream.WriteLine " Please increase the Cabletray width or use small diameter cables, Please check the Cableway:" & m_strCablewayName
        End If
        
        If (m_dblAllowedArea < 0.0001) Then
            'dTotalTraverseArea is already in DBU
            dAllowableTraverseArea = dTotalTraverseArea * m_dblCableTrayFillEff
        Else
            'CalculateMaximumFill would have computed m_dblAllowedArea in Square inches
            'Hence convert it back to DBU
            dAllowableTraverseArea = m_UOM.ConvertUnitToDbu(UNIT_AREA, m_dblAllowedArea, AREA_SQUARE_INCH)
        End If
        m_dblTotalWireArea = m_UOM.ConvertUnitToDbu(UNIT_AREA, m_dblTotalWireArea, AREA_SQUARE_INCH)
        dblAvailArea = (1 - dPercentFull) * dAllowableTraverseArea

        Select Case dblAvailArea
        Case Is = 0
            strFillStatus = "Full"
        Case Is > 0
            strFillStatus = "Partial"
        Case Is < 0
            strFillStatus = "OverFilled"
        End Select

        If dblAvailArea < 0 Then
            dAvailableTraverseArea = 0#
        Else
            dAvailableTraverseArea = dblAvailArea
        End If

        bStatus = strFillStatus
        
        dWireArea = m_dblTotalWireArea
        
        Set m_oLogFileStream = Nothing
        Set m_oLogFile = Nothing

    Exit Sub

ErrorHandler:
    Set m_oServerError = m_oServerErrors.AddFromErr(Err, "Error processing ", METHOD, MODULE)
    m_oServerError.Raise
End Sub

' ######### Function Description ###########
' Purpose : Below four methods are the various methods to calculate the
' Maximum fill percentage
Private Function GetPercentFill1(oArea As Double, sumofarea As Double) As Double
'Private Function GetPercentFill1(oArea As Double) As Double
    GetPercentFill1 = (sumofarea / m_dblCableTrayFillEff) / oArea
    'GetPercentFill1 = (m_dblTotalWireArea / m_dblCableTrayFillEff) / oArea

End Function

Private Function GetPercentFill2(oArea As Double) As Double

    GetPercentFill2 = (m_dblCabArealess1000 / m_dblCableTrayFillEff) / (oArea - (1.1 * m_dblCabDia1000))

End Function

Private Function GetPercentFill3(oArea As Double, dPercent As Double) As Double

    GetPercentFill3 = ((m_dblCabArealess4awg) + (dPercent * m_dblCabDia4AWG)) / (oArea * m_dblCableTrayFillEff)
    'Old formula was GetPercentFill3 = (m_dblCabArealess4awg / m_dblCableTrayFillEff) / (oArea - (dPercent * m_dblCabDia4AWG))

End Function
Private Function FillCalcUsingDiameters(ratio As Double, sumofdiameter As Double) As Double
'Private Function FillCalcUsingDiameters(ratio As Double) As Double
    FillCalcUsingDiameters = (sumofdiameter / m_dblCableTrayFillEff) / (ratio * m_dblCableTrayLoadWidth)
    'FillCalcUsingDiameters = (m_dblCableDiaSum / m_dblCableTrayFillEff) / (ratio * m_dblCableTrayLoadWidth)

End Function

Private Function GetFillEfficiencyFromFeature(oFeatureObj As Object) As Double


    Dim oRteCommonUtility       As IJRteCommonUtility
    Dim varUserInterfaceName    As Variant
    Dim strPropName             As String
    Dim varValue                As Variant


    ' set the user interface name, property for which the value to be retrieved from the cable way feature
    varUserInterfaceName = "IJRteFillInfo"
    strPropName = "FillEfficiency"

    Const METHOD = "GetFillEfficiencyFromFeature"

    Set oRteCommonUtility = Nothing
    '  Query Interface for CommonUtility interface and call the method to get the value.
    Set oRteCommonUtility = oFeatureObj
    Call oRteCommonUtility.GetCustomAttribute(varUserInterfaceName, strPropName, varValue)
    GetFillEfficiencyFromFeature = varValue


End Function

Private Sub GetLoadWidthAndDepth(oCabWyObj As Object)
    Const METHOD = "GetLoadWidthAndDepth"
    On Error GoTo ErrorHandler
    
    Dim oCabWyPathFeat  As IJRtePathFeat
    Dim lCount          As Long
    Dim oElements       As IJElements
    Dim oGenPart        As IJRtePathGenPart
    Dim oPartOcc        As IJDistribPartOccur
    Dim oPortCol        As IJElements
    Dim oCableTrayPortOcc As IJCableTrayPortOcc
    Dim oPortObj        As Object
    Dim oPathFeatUtility As IJRtePathFeatUtility
    Dim oPort1 As IJDistribPort
    Dim oPort2 As IJDistribPort
    
    If Not oCabWyObj Is Nothing Then Exit Sub
    
    Set oPathFeatUtility = oCabWyObj
    If Not oPathFeatUtility Is Nothing Then
        oPathFeatUtility.GetInlineCompBasePortsOfFeature oPort1, oPort2
        If Not oPort1 Is Nothing Then
            Set oCableTrayPortOcc = oPort1
            If Not oCableTrayPortOcc Is Nothing Then
                m_dblCableTrayLoadWidth = oCableTrayPortOcc.LoadWidth
                m_dblCableTrayLoadDepth = oCableTrayPortOcc.LoadDepth
            End If
            Set oCableTrayPortOcc = Nothing
        End If
        Set oPort1 = Nothing
        
        If Not oPort2 Is Nothing Then
            Set oCableTrayPortOcc = oPort2
            If Not oCableTrayPortOcc Is Nothing Then
                If oCableTrayPortOcc.LoadWidth < m_dblCableTrayLoadWidth Then
                    m_dblCableTrayLoadWidth = oCableTrayPortOcc.LoadWidth
                End If
                If oCableTrayPortOcc.LoadDepth < m_dblCableTrayLoadDepth Then
                    m_dblCableTrayLoadDepth = oCableTrayPortOcc.LoadDepth
                End If
            End If
            Set oCableTrayPortOcc = Nothing
        End If
        Set oPort2 = Nothing
    Else
        Set oCabWyPathFeat = oCabWyObj
        Set oElements = oCabWyPathFeat.GetBaseParts(lCount)
        If lCount > 0 Then
            Set oGenPart = oElements.Item(1)
            If Not oGenPart Is Nothing Then
                Set oPartOcc = oGenPart
                Set oPortCol = oPartOcc.GetPorts(DistribPortStatus_BASE, DistribPortType_CABLETRAY)
                Set oPartOcc = Nothing
                For Each oPortObj In oPortCol
                    Set oCableTrayPortOcc = oPortObj
                    'This happens for the first time when the values are 0
                    If (m_dblCableTrayLoadWidth = 0) And (m_dblCableTrayLoadDepth = 0) Then
                        m_dblCableTrayLoadWidth = oCableTrayPortOcc.LoadWidth
                        m_dblCableTrayLoadDepth = oCableTrayPortOcc.LoadDepth
                    End If
                    If oCableTrayPortOcc.LoadWidth < m_dblCableTrayLoadWidth Then
                        m_dblCableTrayLoadWidth = oCableTrayPortOcc.LoadWidth
                    End If
                    If oCableTrayPortOcc.LoadDepth < m_dblCableTrayLoadDepth Then
                        m_dblCableTrayLoadDepth = oCableTrayPortOcc.LoadDepth
                    End If
                    Set oCableTrayPortOcc = Nothing
                    Set oPortObj = Nothing
                Next oPortObj
            End If
            Set oGenPart = Nothing
        End If
        Set oElements = Nothing
        Set oCabWyPathFeat = Nothing
    End If
    Set oPathFeatUtility = Nothing
    Exit Sub
ErrorHandler:
    Set m_oServerError = m_oServerErrors.AddFromErr(Err, "Error processing ", METHOD, MODULE)
    m_oServerError.Raise
End Sub


' ######### Function Description ###########
' Purpose : Initialises the flags required to get the allowable area from
'NEC tables

Private Function ProcessCableProperties(ByVal OCablePart As IJCablePart)

    Const METHOD = "ProcessCableProperties"

    On Error GoTo ErrorHandler

    Dim dKcmil              As Double
    Dim dCurrentWireArea    As Double
    Dim dCableDiameter      As Double
    Dim nConductors         As Long
    
    Dim oAttributes         As IJDAttributes
    Dim oAttributeColl      As IJDAttributesCol
    Dim oAttribute          As IJDAttribute

    On Error Resume Next
    Set oAttributes = OCablePart
    Set oAttributeColl = oAttributes.CollectionOfAttributes("IJCableVCapacity")

    If Not oAttributeColl Is Nothing Then
        Set oAttribute = oAttributeColl.Item("RatedVoltage")
        m_dblRatedVoltage = oAttribute.Value
        Set oAttributeColl = Nothing
    End If


    On Error GoTo ErrorHandler

    If m_dblRatedVoltage >= 2000 Then
        m_bRatedVoltage = True
    End If

    'Find the nature of cables inside the tray or conduit

    Select Case m_intSignalType
    Case 0
        m_bPowerCables = True
    Case 1
        m_bCommunicationCables = True
    Case 2
        m_bControl = True
    Case 3
        m_bDataCables = True
    Case 4
        m_bFireAlarmCables = True
    Case 5
        m_bLightningCables = True
    Case 6
        m_bMultiConPowerCables = True
    Case 7
        m_bPowerCables = True
    Case 8
        m_bSignalCables = True
    End Select

    '=========================================================================
    'to more accurately determine size of conductor for calculations
    'Would be better to use UOM transformation and then to compare AWG value to table in this code.
    
    dCableDiameter = OCablePart.CableDiameter
    dCableDiameter = m_UOM.ConvertDbuToUnit(UNIT_DISTANCE, dCableDiameter, DISTANCE_INCH)
    
    On Error Resume Next
   
    Set oAttributes = OCablePart
    Set oAttributeColl = oAttributes.CollectionOfAttributes("IJCableSize")

    If Not oAttributeColl Is Nothing Then
        Set oAttribute = oAttributeColl.Item("SizeOfConductors")
        m_dblSizeOfConductors = oAttribute.Value
        Set oAttributeColl = Nothing
    End If

    On Error GoTo ErrorHandler
        
    nConductors = OCablePart.NoOfConductors

    If nConductors > 1 Then
        m_blnMulticonductor = True
        
        m_intCableCountMultiCon = m_intCableCountMultiCon + 1
        
        If m_intSignalType = 7 Then
            m_bMultiConPowerCables = True
        End If

        'Sums the total wire area of multiconductor cables
        m_dblMultiTotalWireArea = m_dblMultiTotalWireArea + PI * (dCableDiameter / 2) ^ 2
        
        'Sums the diameters of all the single conductor cables
        m_dblMultiCableDiaSum = m_dblMultiCableDiaSum + dCableDiameter
        
        'Size of 4/0 AWG in database represented by 459.99 and SizeOfConductors_UT is 1009 representing AWG
        If m_dblSizeOfConductors >= 459.99 Then
        'Sums the diameters of all the cables more than 4/0 AWG
            m_blnAWGGreater4 = True
            m_dblCabDia4AWG = m_dblCabDia4AWG + dCableDiameter
        Else
            m_blnAWGLess4 = True
            'Sums up areas of all cables less than 4/0 AWG
            m_dblCabArealess4awg = m_dblCabArealess4awg + PI * (dCableDiameter / 2) ^ 2
        End If

    Else
        m_blnSingleconductor = True

        'Sums the total wire area of single conductor cables
        m_dblSingleTotalWireArea = m_dblSingleTotalWireArea + PI * (dCableDiameter / 2) ^ 2
        
        'Sums the diameters of all the single conductor cables
        m_dblSingleCableDiaSum = m_dblSingleCableDiaSum + dCableDiameter

        If m_dblSizeOfConductors >= 1000 Then
            m_blnKCMIL1000more = True
            'Sums the diameters of all the cables more than 1000 KCMIL
            m_dblCabDia1000 = m_dblCabDia1000 + dCableDiameter
        ElseIf m_dblSizeOfConductors >= 500 And m_dblSizeOfConductors < 1000 Then
            m_blnKCMIL250TO1000 = True
            'Sums the areas of all the cables between 250 KCMIL and 900 KCMIL
            m_dblCabArea250to900 = m_dblCabArea250to900 + PI * (dCableDiameter / 2) ^ 2
            'Sums the areas of all the cables between 250 KCMIL and 900 KCMIL for NEC 392.10(A)(3)
            m_dblCabArealess1000 = m_dblCabArealess1000 + PI * (dCableDiameter / 2) ^ 2
        ElseIf m_dblSizeOfConductors < 500 And m_dblSizeOfConductors >= 324.85 Then
            m_bln1to4ought = True
            m_blnKCMILLess250 = True
            'Continues Summing the areas of all the cables less than 1000 KCMIL for NEC 392.10(A)(3)
            m_dblCabArealess1000 = m_dblCabArealess1000 + PI * (dCableDiameter / 2) ^ 2
        Else
            m_blnKCMILLess250 = True
            'Finishes Summing the areas of all the cables less than 1000 KCMIL for NEC 392.10(A)(3)
            m_dblCabArealess1000 = m_dblCabArealess1000 + PI * (dCableDiameter / 2) ^ 2
        End If
        
    End If
    
    'Sums up areas of all cables
    m_dblTotalWireArea = m_dblTotalWireArea + PI * (dCableDiameter / 2) ^ 2
    
    'Sums the diameters of all the cables
    m_dblCableDiaSum = m_dblCableDiaSum + dCableDiameter
    
    'Area of the current cable
    dCurrentWireArea = PI * (dCableDiameter / 2) ^ 2
    
    'End of Addition
    '===============================================================================
    
    

    '===============================================================================
    ' due to incorrect way to determine cable conductor size as it relates to NEC
    'dCableDiameter = OCablePart.CableDiameter

    'dCableDiameter = m_UOM.ConvertDbuToUnit(UNIT_DISTANCE, dCableDiameter, DISTANCE_INCH)

    'If dCableDiameter > AWG4Diameter Then
    '    m_blnAWGGreater4 = True

        'Sums the diameters of all the cables more than 4AWG
    '    m_dblCabDia4AWG = m_dblCabDia4AWG + dCableDiameter

    'Else
    '    m_blnAWGLess4 = True
        'Sums up areas of all cables less than 4 AWG
    '    m_dblCabArealess4awg = m_dblCabArealess4awg + PI * (dCableDiameter / 2) ^ 2
    'End If
    
    'Sums up areas of all cables
    'm_dblTotalWireArea = m_dblTotalWireArea + PI * (dCableDiameter / 2) ^ 2
    'Area of the current cable
    'dCurrentWireArea = PI * (dCableDiameter / 2) ^ 2

    'dKcmil = dCurrentWireArea / KCMIL

    'If dKcmil > 1000 Then
    '    m_blnKCMIL1000more = True
        'Sums the diameters of all the cables more than 1000 KCMIL
    '    m_dblCabDia1000 = m_dblCabDia1000 + dCableDiameter
    'ElseIf dKcmil > 250 Then
    '    m_blnKCMIL250TO1000 = True
        'Sums the areas of all the cables less than 1000 KCMIL
    '    m_dblCabArealess1000 = m_dblCabArealess1000 + PI * (dCableDiameter / 2) ^ 2
    'Else
    '    m_blnKCMILLess250 = True
        'Sums the areas of all the cables less than 1000 KCMIL
    '    m_dblCabArealess1000 = m_dblCabArealess1000 + PI * (dCableDiameter / 2) ^ 2

    'End If
    
    'Sums the diameters of all the cables
    'm_dblCableDiaSum = m_dblCableDiaSum + dCableDiameter
    '===============================================================================

    Exit Function

ErrorHandler:
    Set m_oServerError = m_oServerErrors.AddFromErr(Err, "Error processing ", METHOD, MODULE)
    m_oServerError.Raise
End Function

' ######### Function Description ###########
' Purpose : Algorithm to calculate the appropiate allowed are from the Nec tables

Private Function CalculateMaximumFill() As Double
'The followed code is according to the NEC Specification
'in the Article 392 of cable trays

    Const METHOD = "CalculateMaximumFill"
    
    ' to separate single and multiconductor fill % for NEC 2008 requirements
    Dim CalculateMaxFillSingle  As Double
    Dim CalculateMaxFillMulti   As Double
    
    CalculateMaxFillSingle = 0#
    CalculateMaxFillMulti = 0#

    On Error GoTo ErrorHandler

    If m_bRatedVoltage Then
        m_dblAllowedArea = m_dblCableTrayLoadWidth    'LoadWidth
        CalculateMaximumFill = FillCalcUsingDiameters(1#, m_dblCableDiaSum)
        'CalculateMaximumFill = FillCalcUsingDiameters(1#)
    Else
        If m_blnSingleconductor Then
        'If (m_blnSingleconductor And Not m_blnMulticonductor) Then
            
            'Process single conductor cables
            If (m_intCableTray = CABLETRAY_LADDER) Or (m_intCableTray = CABLETRAY_TROUGH_VENTILATED) Then
                If m_blnKCMIL1000more And Not m_blnKCMIL250TO1000 And Not m_bln1to4ought And Not m_blnKCMILLess250 Then
                'If m_blnKCMIL1000more And Not m_blnKCMIL250TO1000 And Not m_blnKCMILLess250 Then
                    CalculateMaxFillSingle = FillCalcUsingDiameters(1#, m_dblCabDia1000)   'LoadWdth
                    'CalculateMaximumFill = FillCalcUsingDiameters(1#)    'LoadWdth
                ElseIf m_blnKCMIL250TO1000 And Not m_blnKCMIL1000more And Not m_bln1to4ought And Not m_blnKCMILLess250 Then
                'ElseIf m_blnKCMIL250TO1000 And Not m_blnKCMIL1000more And Not m_blnKCMILLess250 Then
                    'Need to retrieve Area from NominalWidth
                    m_dblAllowedArea = GetAllowedArea(TABLE39210A)    'NominalWidth
                    CalculateMaxFillSingle = GetPercentFill1(m_dblAllowedArea, m_dblCabArea250to900)
                    'CalculateMaximumFill = GetPercentFill1(m_dblAllowedArea)
                Else
                    If m_blnKCMILLess250 And m_bln1to4ought Then
                    'If m_blnAWGLess4 Then
                        CalculateMaxFillSingle = FillCalcUsingDiameters(1#, m_dblSingleCableDiaSum)   'LoadWidth
                    Else
                        m_dblAllowedArea = GetAllowedArea(TABLE39210A)    'from NominalWidth
                        CalculateMaxFillSingle = GetPercentFill2(m_dblAllowedArea)
                    End If
                End If

            Else

                If (m_intCableTray = CABLETRAY_CHANNEL_VENTILATED) Then
                    CalculateMaxFillSingle = FillCalcUsingDiameters(1#, m_dblSingleCableDiaSum)
                Else
                    'CreateToDoRecord
                    m_oLogFileStream.WriteLine " "
                    m_oLogFileStream.WriteLine " Please check with the cableway: " & m_strCablewayName
                    m_oLogFileStream.WriteLine " Cable Fill calculation is not supported by S3D for Single Conductor cables if Cabletray type is Solid Bottom or Solid Channel"
                End If
            End If
        End If
        
        If m_blnMulticonductor Then
            'Process Multi conductor cables
            If (m_intCableTray = CABLETRAY_LADDER) Or (m_intCableTray = CABLETRAY_SOLID_BOTTOM) Or (m_intCableTray = CABLETRAY_TROUGH_VENTILATED) Then
            'If (m_intCableTray = CABLETRAY_LADDER) Or (m_intCableTray = CABLETRAY_TROUGH_SOLID_BOTTOM) Or (m_intCableTray = CABLETRAY_TROUGH_VENTILATED) Then

                'check the cable signal type
                If (m_bMultiConPowerCables Or m_bLightningCables) Then
                    If m_blnAWGGreater4 Then
                        If Not m_blnAWGLess4 Then
                            If (m_intCableTray = CABLETRAY_SOLID_BOTTOM) Then
                            'If (m_intCableTray = CABLETRAY_TROUGH_SOLID_BOTTOM) Then
                                CalculateMaxFillMulti = FillCalcUsingDiameters(0.9, m_dblMultiCableDiaSum)
                            Else
                                CalculateMaxFillMulti = FillCalcUsingDiameters(1#, m_dblMultiCableDiaSum)
                            End If
                        Else
                            If (m_intCableTray = CABLETRAY_SOLID_BOTTOM) Then
                            'If (m_intCableTray = CABLETRAY_TROUGH_SOLID_BOTTOM) Then
                                'retrive column4 of 392.9
                                m_dblAllowedArea = GetAllowedArea(TABLE3929C)    'Nominal Width
                                CalculateMaxFillMulti = GetPercentFill3(m_dblAllowedArea, 1)
                            Else
                                'retrive column2 of 392.9
                                m_dblAllowedArea = GetAllowedArea(TABLE3929A)    'Nominal Width
                                CalculateMaxFillMulti = GetPercentFill3(m_dblAllowedArea, 1.2)
                            End If
                        End If
                    Else
                        If (m_intCableTray = CABLETRAY_SOLID_BOTTOM) Then
                        'If (m_intCableTray = CABLETRAY_TROUGH_SOLID_BOTTOM) Then
                            'retrive column3 of 392.9
                            m_dblAllowedArea = GetAllowedArea(TABLE3929C)    'Nominal
                        Else
                            'retrive column1 of 392.9
                            m_dblAllowedArea = GetAllowedArea(TABLE3929A)    'Nominal
                        End If
                        CalculateMaxFillMulti = GetPercentFill1(m_dblAllowedArea, m_dblMultiTotalWireArea)
                    End If
                Else
                    CalculateMaxFillMulti = ProcessSignalCables
                End If

            Else
                If (m_intCableTray = CABLETRAY_CHANNEL_VENTILATED) Then
                    If (m_intCableCountMultiCon = 1) Then
                        m_dblAllowedArea = GetAllowedArea(TABLE3929E1)    'Nominal
                    Else
                        m_dblAllowedArea = GetAllowedArea(TABLE3929E2)    'Nominal
                    End If
                    CalculateMaxFillMulti = GetPercentFill1(m_dblAllowedArea, m_dblMultiTotalWireArea)
                Else
                    If (m_intCableTray = CABLETRAY_CHANNEL_SOLID) Then
                        If (m_intCableCountMultiCon = 1) Then
                            m_dblAllowedArea = GetAllowedArea(TABLE3929F1)    'Nominal
                        Else
                            m_dblAllowedArea = GetAllowedArea(TABLE3929F2)    'Nominal
                        End If
                        CalculateMaxFillMulti = GetPercentFill1(m_dblAllowedArea, m_dblMultiTotalWireArea)
                    Else
                        'ErrorLog
                        m_oLogFileStream.WriteLine " "
                        m_oLogFileStream.WriteLine " Please check with the cableway: " & m_strCablewayName
                        m_oLogFileStream.WriteLine " Cable Tray Type is not handled by Fill Calculation module, please check Tray Type or update VB Module."
                    End If
                End If
            End If
        End If
    End If
    
    'Check to see if max fill was calculated for single or multiconductor power cables
    If (CalculateMaxFillSingle = 0# And CalculateMaxFillMulti = 0#) Then
    'Control/Signal Cable Fill
    CalculateMaximumFill = CalculateMaximumFill
    Else
    'Add both single and multiconductor fill % together for NEC 2008 compliance
    CalculateMaximumFill = CalculateMaxFillSingle + CalculateMaxFillMulti
    End If
    
    m_oLogFileStream.WriteLine " "
    m_oLogFileStream.WriteLine " Cable fill for the feature = " & CalculateMaximumFill * 100

    Exit Function

ErrorHandler:
    Set m_oServerError = m_oServerErrors.AddFromErr(Err, "Error processing ", METHOD, MODULE)
    m_oServerError.Raise
End Function

'
' ######### ProcessSignalCables ###########
' Purpose : Calculates the fill percentage for Signal cables

Private Function ProcessSignalCables() As Double

    Dim dblFillFactor   As Double
    Dim dblArea         As Double



    If (m_bControl Or m_bSignalCables) And Not (m_bCommunicationCables Or m_bDataCables Or m_bFireAlarmCables Or m_bLightningCables Or m_bMultiConPowerCables) Then

    '======================================================================
    'to match Cable Fill Logic Diagram and NEC 392.9(B)(D) language
            
    'Original Commented Out
        
        'If m_dblCableTrayLoadDepth <= 6 Then    ' LoadDepth

         '   If (m_intCableTray = CABLETRAY_LADDER) Or (m_intCableTray = CABLETRAY_TROUGH_VENTILATED) Then
          '      dblFillFactor = 0.5
           ' ElseIf (m_intCableTray = CABLETRAY_SOLID_BOTTOM) Then
            '    dblFillFactor = 0.4
            'Else
                'Error Log

            'End If
            'm_dblAllowedArea = dblFillFactor * m_dblCableTrayLoadDepth * m_dblCableTrayLoadWidth    'Load

        'Else
                
         
            'dblFillFactor = 6
            'm_dblAllowedArea = dblFillFactor * m_dblCableTrayLoadWidth    ' Load

        'End If
    '======================================================================
    
    '======================================================================
    'Added section
    
        If m_dblCableTrayLoadDepth >= 6 Then
            m_dblCableTrayLoadDepth = 6
        Else
            'Do Nothing.  Tray load depth is calculated as is.
        End If
        
        If (m_intCableTray = CABLETRAY_LADDER) Or (m_intCableTray = CABLETRAY_TROUGH_VENTILATED) Then
            dblFillFactor = 0.5
        ElseIf (m_intCableTray = CABLETRAY_SOLID_BOTTOM) Then
            dblFillFactor = 0.4
        Else
            'Error Log
            m_oLogFileStream.WriteLine " "
            m_oLogFileStream.WriteLine " Please check with the cableway: " & m_strCablewayName
            m_oLogFileStream.WriteLine " Cable Tray Type is not handled by Fill Calculation module, please check Tray Type or update VB Module"
        End If
        
        m_dblAllowedArea = dblFillFactor * m_dblCableTrayLoadDepth * m_dblCableTrayLoadWidth    'Load
        
        ProcessSignalCables = GetPercentFill1(m_dblAllowedArea, m_dblMultiTotalWireArea)

    Else
        'ErrorLog create
        m_oLogFileStream.WriteLine " "
        m_oLogFileStream.WriteLine " Please check with the cableway: " & m_strCablewayName
        m_oLogFileStream.WriteLine " Cable Tray Type is not handled by Fill Calculation module, please check Tray Type or update VB Module"
    End If
End Function


'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    Method:        IJRteCARRealTimeFill_CalculateRealTimeFill  (Created 28/04/06)     '
'                                                                                      '
'    Type:          private                                                            '
'                                                                                      '
'    Description:   Realtime fill is the fill calculated before the cable is commited
'                   through the feature.We use a global CableDef and find the realtime
'                   fill.Handled the Parallel Cables as well.

'                                                                                      '
'    Inputs:        ByVal CabDef As Object, ByVal pArc As Object
'                                                                                      '
'    OutPuts:       double                                                             '
'                                                                                      '
'                                                                                      '
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Function IJRteCARRealTimeFill_CalculateRealTimeFill(ByVal CabDef As Object, ByVal pArc As Object) As Double

    Dim oCarArc             As IJCARArc
    Dim oEle                As IJElements
    Dim oPathFeat           As IJRtePathFeat
    Dim strTradeSize        As String
    Dim lNoOfCable          As Long
    Dim dWireArea           As Double
    Dim dblTotTravArea      As Double
    Dim dblAllowTravArea    As Double
    Dim dblPercentFull      As Double
    Dim dblAvailTravArea    As Double
    Dim strFillStatus       As String
    Dim oObj                As Object
    Dim dFinalRealTimeFill  As Double
    Dim dCurrentRealTimeFill As Double
    Dim oRouteUtils         As IJRteCableUtility
    Dim dFillFactor         As Double
    Dim oCarNode            As IJCARNode

    Const METHOD = "IJRteCARRealTimeFill_CalculateRealTimeFill"

    On Error GoTo ErrorHandler


    Set oRouteUtils = New RouteUtility


    Set m_CableDef = CabDef
    Set oCarArc = pArc
    Set oEle = oCarArc.GetBO
    If oEle Is Nothing Then
        Exit Function
    End If


    On Error GoTo ErrorHandler
    dCurrentRealTimeFill = 0

    dFillFactor = oCarArc.FillFactor    'fill of the commmited cables in the feature
    dCurrentRealTimeFill = oCarArc.CurrentRealTimeFill    ' real time fill including the non commited cables

    dFinalRealTimeFill = 0
    For Each oObj In oEle
        ReInitializeVariables
        Set oPathFeat = Nothing
        'The collection of objects could be Nodes or features
        If TypeOf oObj Is IJRtePathFeat Then
            Set oPathFeat = oObj
        ElseIf TypeOf oObj Is IJCARNode Then
            Set oCarNode = Nothing
            Set oCarNode = oObj
            Set oPathFeat = oCarNode.GetBO
        End If
        If TypeOf oPathFeat Is IJRteCablewayPathFeat Then
            IJDFillCalculations_GetCwayFillParams oPathFeat, Nothing, _
                                                  lNoOfCable, dWireArea, "", dblTotTravArea, dblAllowTravArea, dblPercentFull, dblAvailTravArea, strFillStatus
        ElseIf TypeOf oPathFeat Is IJRteConduitPathFeat Then
            IJDFillCalculations_GetConduitFillParams oPathFeat, Nothing, _
                                                     lNoOfCable, dWireArea, "", dblTotTravArea, dblAllowTravArea, dblPercentFull, dblAvailTravArea, strFillStatus
        End If
        If dFinalRealTimeFill < dblPercentFull Then
            dFinalRealTimeFill = dblPercentFull
        End If
    Next
    If m_CableDef.bBelongsToParallelCable = True Then
        Dim lNoOfSiblings As Long
        lNoOfSiblings = m_CableDef.NumberOfSibs
        dFinalRealTimeFill = dFinalRealTimeFill - dFillFactor
        dFinalRealTimeFill = dFinalRealTimeFill * lNoOfSiblings
        dFinalRealTimeFill = dFinalRealTimeFill + dFillFactor
    End If
    IJRteCARRealTimeFill_CalculateRealTimeFill = dFinalRealTimeFill '+ dCurrentRealTimeFill

    Exit Function

ErrorHandler:
    Set m_oServerError = m_oServerErrors.AddFromErr(Err, "Error processing ", METHOD, MODULE)
    m_oServerError.Raise
End Function


'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    Method:        ReInitializeVariables  (Created 12/08/06)                   '
'                                                                                      '
'    Type:          Private                                                           '
'                                                                                      '
'    Description:   Reinitialize all the globals such that correct value is calculated for next
'                   iteration.

'                                                                                      '
'    Inputs:        None
'                                                                                      '
'    OutPuts:       None                                                   '
'                                                                                      '
'                                                                                      '
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Sub ReInitializeVariables()


    m_dblCableDiaSum = 0#                'Sum of diameters of all cables inside the tray
    m_dblCableTrayLoadWidth = 0#        'Width of the cable tray
    m_dblCableTrayLoadDepth = 0#        'Depth of the cable tray
    m_dblCableTrayNominalWidth = 0#     'Width of the cable tray
    m_dblCableTrayNominalDepth = 0#     'Depth of the cable tray
    m_dblCableTrayFillEff = 0#           'Fill efficiency
    m_dblTotalWireArea = 0#              'Total area of the wire
    m_dblCabDia1000 = 0#                 'Sum of diameter of cables more than 1000KCMIL
    m_dblCabArealess1000 = 0#            'Sum of area of cables less than 1000KCMIL
    m_dblCabDia4AWG = 0#                 'Sum of diameter of cables more than 4 AWG
    m_dblCabArealess4awg = 0#            'Sum of area of cables less than 4 AWG
    m_dblAllowedArea = 0#
    m_dblRatedVoltage = 0#

    m_intSignalType = 0

    m_bCommunicationCables = False
    m_bControl = False
    m_bDataCables = False
    m_bFireAlarmCables = False
    m_bLightningCables = False
    m_bMultiConPowerCables = False
    m_bPowerCables = False
    m_bSignalCables = False
    m_bRatedVoltage = False

    m_blnKCMIL1000more = False             'Flag to check if the cable area is more than 1000 KCMIL
    m_blnKCMIL250TO1000 = False            'Flag to check if the cable area is betwn 250 to 1000 KCMIL
    m_blnKCMILLess250 = False              'Flag to check if the cable area is less than 250 KCMIL
    m_blnAWGGreater4 = False               'Flag to check if the cable diameter is more than 4 AWG
    m_blnAWGLess4 = False                  'Flag to check if the cable diameter is less than 4 AWG
    m_blnMulticonductor = False            'Flag to check if the cable is a multi conductor
    m_blnSingleconductor = False           'Flag to check if the cable is a single conductor

    m_intCableCount = 0               'Total number of cable inside the Cabletray
    m_intCableTray = 0                  'Indicates CabelTray type


    '====================================================================================================
    ' for the purpose of splitting area calculations between single and multiconductor
    'cable to allow for NEC 392.11(c) calculations
    m_bln1to4ought = False                  'Flag to check if single conductor cable is between 1/0AWG and 4/0AWG
    
    m_dblMultiTotalWireArea = 0#            'Total area of multiconductor cables
    m_dblSingleTotalWireArea = 0#           'Total area of single conductor cables
    m_dblMultiCableDiaSum = 0#              'Total width of multiconductor cables
    m_dblSingleCableDiaSum = 0#             'Total width of single conductor cables
    m_dblCabArea250to900 = 0#               'Sum of area of cables between 250KCMIL and 900KCMIL (NEC 392.10(2))
    m_dblSizeOfConductors = 0#              'More accurately determine conductor size for NEC calculations
    '====================================================================================================

End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    Method:        GetPOM                                                      '
'                                                                                                '
'    Type:          Private                                                                      '
'                                                                                                '
'    Description:    Gets IJDPOM Object
'                                                                                                '
'                                                                                                '
'    Inputs:        strDbType As String (DatabaseType)
'    OutPuts:       IJDPOM                                                                         '
'                                                                                                '
'                                                                                                '
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


Private Function GetPOM(strDbType As String) As IJDPOM
    Const METHODNAME = "GetPOM"
    On Error GoTo ErrHandler
    
    Dim oContext            As IJContext
    Dim oAccessMiddle       As IJDAccessMiddle
    Dim strConnectMiddle    As String

    strConnectMiddle = "ConnectMiddle"
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService(strConnectMiddle)
    Set GetPOM = oAccessMiddle.GetResourceManagerFromType(strDbType)
    Exit Function
ErrHandler:
    Set m_oServerError = m_oServerErrors.AddFromErr(Err, "Failed getting POM ", METHODNAME, MODULE)
    m_oServerError.Raise
End Function

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'    Method:        CheckIfIsBarrier                                                      '
'                                                                                                '
'    Type:          Private                                                                      '
'                                                                                                '
'    Description:    Check if the input IJRtePathFeat type has an enclosing run.                 '
'
'    Inputs:        oCbleWayFeat As IJRtePathFeat
'    OutPuts:       m_bIsCaseOfBarrier,m_traytype are populated in this method.
'
'                                                                                                '
'                                                                                                '
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Sub CheckIfIsBarrierAndDeriveTrayType(oCbleWayFeat As IJRtePathFeat)
    Const METHODNAME = "CheckIfIsBarrierAndDeriveTrayType"
    On Error GoTo ErrHandler
    
    
    m_bIsCaseOfBarrier = False
    Dim oEnclosedRun As IJRteEnclosedRun
    Dim oEnclosingrun As IJRteEnclosingRun
    Dim oEnclosingRunColl As IJDObjectCollection
    Dim oCbleTrayFeat As IJRtePathFeat
   
    Dim oCableTrayPart As IJCableTrayPart
   
    Dim oEnclosingPathRun As IJRtePathRun
    Dim oEnclosingrunObj  As Object

    Dim oCblwyPartCol       As IJDObjectCollection
    Dim oCblwyRun           As IJRtePathRun
    Dim oTempPart           As IJDPart
    Dim oCblPartObj         As Object
    Dim oCableTrayPartObj   As IJPartOcc
    Dim eEnclosingType As Long
    
    Dim oPathRun As IJRtePathRun
    Set oPathRun = oCbleWayFeat.GetPathRun
    Set oEnclosedRun = oPathRun
    Set oEnclosingRunColl = oEnclosedRun.GetEnclosingRuns
    
    'Checks if the count of enclosing run collection to determine if it is a barrier
    If oEnclosingRunColl.Count > 0 Then
        For Each oEnclosingrunObj In oEnclosingRunColl
            On Error Resume Next
            If TypeOf oEnclosingrunObj Is IJRteEnclosingRun Then
                Set oEnclosingrun = oEnclosingrunObj
                Exit For
            End If
        Next oEnclosingrunObj
        On Error GoTo ErrHandler
        
        Set oCblwyRun = oEnclosingrun
            Set oCblwyPartCol = oCblwyRun.GetParts
            If oCblwyPartCol.Count >= 1 Then
               For Each oCblPartObj In oCblwyPartCol
                    On Error Resume Next
                    Set oCableTrayPart = Nothing
                    Set oCableTrayPartObj = oCblPartObj
                    If Not oCableTrayPartObj Is Nothing Then
                        oCableTrayPartObj.GetPart oTempPart
                        Set oCableTrayPart = oTempPart
                        Exit For
                    End If
                Next oCblPartObj
            End If
            If Not oCableTrayPart Is Nothing Then
                m_intCableTray = oCableTrayPart.TrayType
            Else
                m_intCableTray = CABLETRAY_LADDER 'This case should not arise.As of now deafaulting the tray type to ladder.
            End If
        
        m_bIsCaseOfBarrier = True
    End If
    Exit Sub
ErrHandler:
    Set m_oServerError = m_oServerErrors.AddFromErr(Err, "Failed CheckIfIsBarrierAndDeriveTrayType ", METHODNAME, MODULE)
    m_oServerError.Raise
End Sub

''*********************************************************************
''  Routine     : GetSystemTempPath
''  Description : Gets the 'Temp' path of the Local machine.
''*********************************************************************
Private Function GetSystemTempPath() As String
    Dim sTmpPath As String * 512
    Dim nRet As Long
    On Error GoTo ErrorHandler
    nRet = GetTempPath(512, sTmpPath)
    GetSystemTempPath = Left$(sTmpPath, InStr(1, sTmpPath, vbNullChar) - 1)
    
    Exit Function
ErrorHandler:
    MsgBox "Unexpected Error : " & Err.Description
End Function
