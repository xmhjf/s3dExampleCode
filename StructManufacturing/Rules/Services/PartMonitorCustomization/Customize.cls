VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Customize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''***********************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : StrMfgPartMonitorInfo
''  File        : Customize.cls
''
''  Description : This class has the functionality to Start the commnad,listening to events of Ribbonbars
''
''  Author      : Kristian Kamph
''
''  History     :
''                16Jun06       Kristian Kamph
''                                  - Creation
''                05Oct06       Kristian Kamph
''                                  DI-CP·96742 - update part monitor to support profile xml
''***********************************************************************************
Option Explicit

Implements IJMfgPartMonitorInfo
Private m_oUnitsofMeasure As IJUomVBInterface

'List of Symbol name constants:
Private Const BEVEL_01_SYMBOL = "\Bevel1.sym" 'No depths or angles: Really means no bevel  OR 'Flat nose, single bevel
Private Const BEVEL_02_SYMBOL = "\Bevel2.sym" 'Angled nose, single bevel
Private Const BEVEL_03_SYMBOL = "\Bevel3.sym" 'Flat nose with bevel on one side
Private Const BEVEL_04_SYMBOL = "\Bevel4.sym" 'Angled nose with bevel on one side
Private Const BEVEL_05_SYMBOL = "\Bevel5.sym" 'Flat nose with bevels on both sides
Private Const BEVEL_06_SYMBOL = "\Bevel6.sym" 'Angled nose with bevels on both sides
Private Const BEVEL_07_SYMBOL = "\Bevel7.sym" 'Flat nose with 1 bevel on one side and 2 on the other side
Private Const BEVEL_08_SYMBOL = "\Bevel8.sym" 'Angled nose with 1 bevel on one side and 2 on the other side
Private Const BEVEL_09_SYMBOL = "\Bevel9.sym" '5 section bevel with flat nose
Private Const BEVEL_10_SYMBOL = "\Bevel10.sym" '5 section bevel with angled nose
Private Const BEVEL_UNSPECIFIED_SYMBOL = "\UnSpecBevel.sym" 'This should not happen. We return default symbol.
Private Const GRIND_SYMBOL = "\Grind.sym"
Private Const KNUCKLE_ANGLE_SYMBOL = "\KnuckleAngle.sym"
Private Const MARGIN_SYMBOL = "\Margin.sym"
Private Const MARGIN_AS_MARGIN_SYMBOL = "\MarginAsMargin.sym"
Private Const MARGIN_AS_SHRINKAGE_SYMBOL = "\MarginAsShrinkage.sym"
Private Const MARGIN_AS_REFERENCE_SYMBOL = "\MarginAsReference.sym"
Private Const PART_DIRECTION_LR_SYMBOL = "\PartDirLLFlangeRR.sym" 'Part direction L, Flange on Right side of line
Private Const PART_DIRECTION_L_SYMBOL = "\PartDirLL.sym" 'Part direction Left, Flange on same side if present
Private Const PART_DIRECTION_RL_SYMBOL = "\PartDirRRFlangeLL.sym" 'Part direction R, Flange on Left side of line
Private Const PART_DIRECTION_R_SYMBOL = "\PartDirRR.sym" 'Part direction Right, Flange on same side if present
Private Const PART_DIRECTION_CR_SYMBOL = "\PartDirCR.sym" 'Part direction centered, no flange
Private Const PART_DIRECTION_CL_SYMBOL = "\PartDirCL.sym" 'Part direction centered, Flange on Left side of line
Private Const PART_DIRECTION_C_SYMBOL = "\PartDirC.sym" 'Part direction centered, Flange on Right side of line
Private Const SHRINKAGE_SYMBOL = "\Shrinkage.sym"

'Path for symbols
Private Const PATH_SYMBOL_SHARE_EXTENSION = "\StructManufacturing\SMS_SCHEMA\RadSymbols"
Private Const PATH_DEV_SYMBOL_EXTENSION = "\Data\Symbols\StructManufacturing\SMS_SCHEMA\RadSymbols"

Private SymPath As String

Private Const MODULE = "Customize"

Private Sub Class_Initialize()
    Set m_oUnitsofMeasure = New UnitsOfMeasureServicesLib.UomVBInterface
    SymPath = GetSymbolSharePath()
End Sub

Private Sub Class_Terminate()
    Set m_oUnitsofMeasure = Nothing
End Sub


Public Function IJMfgPartMonitorInfo_DetermineSymbol(ByVal oAttributeSets As Object, strSymbolType As String) As String
    Const METHOD = "DetermineSymbol"
    On Error GoTo ErrorHandler
        
    Select Case UCase(strSymbolType)
        Case "MARKING"
            IJMfgPartMonitorInfo_DetermineSymbol = DetermineMarkingSymbol(oAttributeSets)
        Case "MARGIN"
            IJMfgPartMonitorInfo_DetermineSymbol = DetermineMarginSymbol(oAttributeSets)
        Case "SHRINKAGE"
            IJMfgPartMonitorInfo_DetermineSymbol = DetermineShrinkageSymbol(oAttributeSets)
        Case "BEVEL"
            IJMfgPartMonitorInfo_DetermineSymbol = DetermineBevelSymbol(oAttributeSets)
        Case "GRIND"
            IJMfgPartMonitorInfo_DetermineSymbol = DetermineGrindSymbol(oAttributeSets)
    End Select
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

Private Function DetermineShrinkageSymbol(oAttributeSets As AttributeSets) As String
    Const METHOD = "DetermineShrinkageSymbol"
    On Error GoTo ErrorHandler
    DetermineShrinkageSymbol = SymPath & SHRINKAGE_SYMBOL
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

Private Function DetermineMarginSymbol(oAttributeSets As AttributeSets) As String
    Const METHOD = "DetermineMarginSymbol"
    On Error GoTo ErrorHandler
    Dim strGeometry As String
    strGeometry = GetAttributeValue(oAttributeSets, "GEOMETRY")
    If UCase(strGeometry) = "MARGIN" Then
        DetermineMarginSymbol = SymPath & MARGIN_AS_MARGIN_SYMBOL
    ElseIf UCase(strGeometry) = "SHRINKAGE" Then
        DetermineMarginSymbol = SymPath & MARGIN_AS_SHRINKAGE_SYMBOL
    ElseIf UCase(strGeometry) = "REFERENCE" Then
        DetermineMarginSymbol = SymPath & MARGIN_AS_REFERENCE_SYMBOL
    Else
        DetermineMarginSymbol = SymPath & MARGIN_SYMBOL
    End If
Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

Private Function DetermineMarkingSymbol(oAttributeSets As AttributeSets) As String
    Const METHOD = "DetermineMarkingSymbol"
    On Error GoTo ErrorHandler
    
    Dim strPartName As String, strPartDirection As String, strMkgType As String
    Dim strFlangeDirection As String, strThicknessDirection As String
    Dim bDifferentDirections As Boolean
    Dim strFlangeDirection2D As String
    strMkgType = GetAttributeValue(oAttributeSets, "SMS_MARKING||TYPE") '"Marking Type")
            
'    Select Case strMkgType
'    Case "knuckle", "Knuckle"
    If InStr(strMkgType, "knuckle") Or InStr(strMkgType, "Knuckle") Then
        DetermineMarkingSymbol = SymPath & KNUCKLE_ANGLE_SYMBOL
    ElseIf InStr(strMkgType, "location") Or InStr(strMkgType, "Location") Or _
           InStr(strMkgType, "connection") Or InStr(strMkgType, "Connection") Then
'    Case "location", "Location"
        strPartName = GetAttributeValue(oAttributeSets, "SMS_MARKING_PART_LOCATION||LOCATED_PART_NAME") '"Located Part Name")
        strPartDirection = GetAttributeValue(oAttributeSets, "SMS_EDGE||PART_DIR") '"Part Direction")
        strFlangeDirection = GetAttributeValue(oAttributeSets, "SMS_MARKING_PART_LOCATION||FLANGE_DIR") '"Flange Direction")
        strFlangeDirection2D = GetAttributeValue(oAttributeSets, "SMS_MARKING_PART_LOCATION||FLANGE_DIR_2D") '"Flange Direction 2D")
        strThicknessDirection = GetAttributeValue(oAttributeSets, "SMS_MARKING_PART_LOCATION||THICKNESS_DIR") '"Thickness Dir")
        bDifferentDirections = False
        If strFlangeDirection <> "" Then
            If strFlangeDirection <> strThicknessDirection Then
                bDifferentDirections = True
                If strPartDirection = "C" Then
                    strPartDirection = strPartDirection & strFlangeDirection2D
                End If
            End If
        End If
        If strPartDirection = "L" Then
            If bDifferentDirections Then
                DetermineMarkingSymbol = SymPath & PART_DIRECTION_LR_SYMBOL
            Else
                DetermineMarkingSymbol = SymPath & PART_DIRECTION_L_SYMBOL
            End If
        ElseIf strPartDirection = "R" Then
            If bDifferentDirections Then
                DetermineMarkingSymbol = SymPath & PART_DIRECTION_RL_SYMBOL
            Else
                DetermineMarkingSymbol = SymPath & PART_DIRECTION_R_SYMBOL
            End If
        ElseIf InStr(strPartDirection, "C") > 0 Then
            If InStr(strPartDirection, "R") > 0 Then
                DetermineMarkingSymbol = SymPath & PART_DIRECTION_CR_SYMBOL
            ElseIf InStr(strPartDirection, "L") > 0 Then
                DetermineMarkingSymbol = SymPath & PART_DIRECTION_CL_SYMBOL
            Else
                DetermineMarkingSymbol = SymPath & PART_DIRECTION_C_SYMBOL
            End If
        End If
    End If
'    End Select
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

Private Function DetermineBevelSymbol(oAttributeSets As AttributeSets) As String
    Const METHOD = "DetermineBevelSymbol"
    On Error GoTo ErrorHandler
    
    'Angle and Depth A
    Dim sDEPTH_A As String
    Dim dDEPTH_A As Double
    sDEPTH_A = GetAttributeValue(oAttributeSets, "SMS_BEVEL||DEPTH_A")
    If IsNumeric(sDEPTH_A) Then
        dDEPTH_A = CDbl(sDEPTH_A)
    Else
        dDEPTH_A = 0
    End If
    
    Dim sANGLE_A As String
    Dim dANGLE_A As Double
    sANGLE_A = GetAttributeValue(oAttributeSets, "SMS_BEVEL||ANGLE_A")
    If IsNumeric(sANGLE_A) Then
        dANGLE_A = CDbl(sANGLE_A)
    Else
        dANGLE_A = 0
    End If
    'Angle and Depth B
    Dim sDEPTH_B As String
    Dim dDEPTH_B As Double
    sDEPTH_B = GetAttributeValue(oAttributeSets, "SMS_BEVEL||DEPTH_B")
    If IsNumeric(sDEPTH_B) Then
        dDEPTH_B = CDbl(sDEPTH_B)
    Else
        dDEPTH_B = 0
    End If
    
    Dim sANGLE_B As String
    Dim dANGLE_B As Double
    sANGLE_B = GetAttributeValue(oAttributeSets, "SMS_BEVEL||ANGLE_B")
    If IsNumeric(sANGLE_B) Then
        dANGLE_B = CDbl(sANGLE_B)
    Else
        dANGLE_B = 0
    End If
    
    'Angle and Depth N
    Dim sDEPTH_N As String
    Dim dDEPTH_N As Double
    sDEPTH_N = GetAttributeValue(oAttributeSets, "SMS_BEVEL||DEPTH_N")
    If IsNumeric(sDEPTH_N) Then
        dDEPTH_N = CDbl(sDEPTH_N)
    Else
        sDEPTH_N = GetAttributeValue(oAttributeSets, "SMS_BEVEL||DEPTH_C")
        If IsNumeric(sDEPTH_N) Then
            dDEPTH_N = CDbl(sDEPTH_N)
        Else
            dDEPTH_N = 0
        End If
    End If
    
    Dim sANGLE_N As String
    Dim dANGLE_N As Double
    sANGLE_N = GetAttributeValue(oAttributeSets, "SMS_BEVEL||ANGLE_N")
    If IsNumeric(sANGLE_N) Then
        dANGLE_N = CDbl(sANGLE_N)
    Else
        dANGLE_N = 0
    End If
    
    'Angle and Depth D
    Dim sDEPTH_D As String
    Dim dDEPTH_D As Double
    sDEPTH_D = GetAttributeValue(oAttributeSets, "SMS_BEVEL||DEPTH_D")
    If IsNumeric(sDEPTH_D) Then
        dDEPTH_D = CDbl(sDEPTH_D)
    Else
        dDEPTH_D = 0
    End If
    
    Dim sANGLE_D As String
    Dim dANGLE_D As Double
    sANGLE_D = GetAttributeValue(oAttributeSets, "SMS_BEVEL||ANGLE_D")
    If IsNumeric(sANGLE_D) Then
        dANGLE_D = CDbl(sANGLE_D)
    Else
        dANGLE_D = 0
    End If
    
    'Angle and Depth E
    Dim sDEPTH_E As String
    Dim dDEPTH_E As Double
    sDEPTH_E = GetAttributeValue(oAttributeSets, "SMS_BEVEL||DEPTH_E")
    If IsNumeric(sDEPTH_E) Then
        dDEPTH_E = CDbl(sDEPTH_E)
    Else
        dDEPTH_E = 0
    End If
    
    Dim sANGLE_E As String
    Dim dANGLE_E As Double
    sANGLE_E = GetAttributeValue(oAttributeSets, "SMS_BEVEL||ANGLE_E")
    If IsNumeric(sANGLE_E) Then
        dANGLE_E = CDbl(sANGLE_E)
    Else
        dANGLE_E = 0
    End If
    
    'Logic for selecting a bevel symbol based on the Depths and angles
    'Delivered there will be 10 symbols
    Dim iResultBevelNumber As Integer
    iResultBevelNumber = 0
    If Abs(dANGLE_N) > 0 Then
        If dDEPTH_N > 0 Then
            iResultBevelNumber = iResultBevelNumber + 1
        Else
            iResultBevelNumber = iResultBevelNumber + 1000000           'this means error
        End If
    End If
    
    If dDEPTH_N > 0 Then iResultBevelNumber = iResultBevelNumber + 2
    If dDEPTH_B > 0 Then iResultBevelNumber = iResultBevelNumber + 4
    If dDEPTH_D > 0 Then iResultBevelNumber = iResultBevelNumber + 8
    If dDEPTH_A > 0 Then iResultBevelNumber = iResultBevelNumber + 16
    If dDEPTH_E > 0 Then iResultBevelNumber = iResultBevelNumber + 32

    If dDEPTH_N = 0 Then
        iResultBevelNumber = iResultBevelNumber / 2
        If (dDEPTH_E = 0 And (Abs(dANGLE_B) > 0 Or Abs(dANGLE_D) > 0)) Then
            iResultBevelNumber = iResultBevelNumber + 1
        End If
    End If
    
    Select Case iResultBevelNumber
    Case 0
        DetermineBevelSymbol = SymPath & BEVEL_01_SYMBOL 'No depths or angles: Really means no bevel
    Case 2, 4
        DetermineBevelSymbol = SymPath & BEVEL_01_SYMBOL 'Flat nose, single bevel
    Case 3, 5
        DetermineBevelSymbol = SymPath & BEVEL_02_SYMBOL 'Angled nose, single bevel
    Case 6, 10
        DetermineBevelSymbol = SymPath & BEVEL_03_SYMBOL 'Flat nose with bevel on one side
    Case 7, 11
        DetermineBevelSymbol = SymPath & BEVEL_04_SYMBOL 'Angled nose with bevel on one side
    Case 14
        DetermineBevelSymbol = SymPath & BEVEL_05_SYMBOL 'Flat nose with bevels on both sides
    Case 15
        DetermineBevelSymbol = SymPath & BEVEL_06_SYMBOL 'Angled nose with bevels on both sides
    Case 30, 46
        DetermineBevelSymbol = SymPath & BEVEL_07_SYMBOL 'Flat nose with 1 bevel on one side and 2 on the other side
    Case 31, 47
        DetermineBevelSymbol = SymPath & BEVEL_08_SYMBOL 'Angled nose with 1 bevel on one side and 2 on the other side
    Case 62
        DetermineBevelSymbol = SymPath & BEVEL_09_SYMBOL '5 section bevel with flat nose
    Case 63
        DetermineBevelSymbol = SymPath & BEVEL_10_SYMBOL '5 section bevel with angled nose
    Case Else
        DetermineBevelSymbol = SymPath & BEVEL_UNSPECIFIED_SYMBOL 'This should not happen. We return default symbol.
    End Select
Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

Private Function DetermineGrindSymbol(oAttributeSets As AttributeSets) As String
    Const METHOD = "DetermineGrindSymbol"
    On Error GoTo ErrorHandler
    DetermineGrindSymbol = SymPath & GRIND_SYMBOL
Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

Public Function IJMfgPartMonitorInfo_GenerateReferenceXML(ByVal oPart As Object) As Object
    Const METHOD = "GenerateReferenceXML"
    On Error GoTo ErrorHandler
    
    Dim oDomDoc         As DOMDocument
    Dim oDomElement     As IXMLDOMElement
    Dim oDomElement1    As IXMLDOMElement
    Dim oMfgPart        As Object
    'Determine the Global Direction of the Upside
    Dim strUpside       As String
    
    Set oDomDoc = New DOMDocument
    Set oDomElement = oDomDoc.createElement("PROPERTIES")
    oDomDoc.appendChild oDomElement
    
    If TypeOf oPart Is IJMfgPlatePart Or TypeOf oPart Is IJMfgProfilePart Then
        On Error Resume Next
        strUpside = GetGlobalUpsideDirection(oPart)
        On Error GoTo ErrorHandler
        
        Set oDomElement1 = oDomDoc.createElement("PROPERTY")
        oDomElement1.setAttribute "NAME", "SMS_REFERENCE||UPSIDE_DIR"
        oDomElement1.setAttribute "VALUE", strUpside
        oDomElement.appendChild oDomElement1
    End If
    
    If TypeOf oPart Is IJMfgNestData Then
        Dim oNestData As IJMfgNestData
        Set oNestData = oPart
        Set oMfgPart = oNestData.Parent
    Else
        Set oMfgPart = oPart
    End If
    
    If TypeOf oMfgPart Is IJMfgPlatePart Or TypeOf oMfgPart Is IJMfgProfilePart Or TypeOf oMfgPart Is IJDMfgTemplateSet Then
        Dim strPartStatus           As String
        Dim strMfgStatus            As String
        Dim strSurvivabilityStatus  As String
            
        strPartStatus = GetPartStatus(oMfgPart)
        strMfgStatus = GetMfgStatus(oMfgPart)
        
        If strPartStatus <> "Unknown" Then
            Set oDomElement1 = oDomDoc.createElement("PROPERTY")
            oDomElement1.setAttribute "NAME", "SMS_REFERENCE||PART_STATUS"
            oDomElement1.setAttribute "VALUE", strPartStatus
            oDomElement.appendChild oDomElement1
        End If
        
        If strMfgStatus <> "Unknown" Then
            Set oDomElement1 = oDomDoc.createElement("PROPERTY")
            oDomElement1.setAttribute "NAME", "SMS_REFERENCE||MFG_OBJECT_STATUS"
            oDomElement1.setAttribute "VALUE", strMfgStatus
            oDomElement.appendChild oDomElement1
        End If
        
        If TypeOf oMfgPart Is IJMfgPlatePart Or TypeOf oMfgPart Is IJMfgProfilePart Then
            
            strSurvivabilityStatus = GetSurvivabilityStatus(oMfgPart)
            
            If strSurvivabilityStatus <> "Unknown" Then
                Set oDomElement1 = oDomDoc.createElement("PROPERTY")
                oDomElement1.setAttribute "NAME", "SMS_REFERENCE||SURVIVABILITY"
                oDomElement1.setAttribute "VALUE", strSurvivabilityStatus
                oDomElement.appendChild oDomElement1
            End If
        End If
    End If
    
    Set IJMfgPartMonitorInfo_GenerateReferenceXML = oDomDoc
CleanUp:
    Set oDomDoc = Nothing
    Set oDomElement = Nothing
    Set oDomElement1 = Nothing
    Set oMfgPart = Nothing
    Set oNestData = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

'------------------------------------------------------------------------------
'Routine      :   GetAttributeValue
'Abstract     :   Given the AttributeSets and the attribute name, this function returns the attribute value
'                 In order to differentiate between the String and Double attributes, the attributes are created
'                 with igAttrTypeString for strings and igAttrTypeArray types for double values ( Just to make Units to be same as UOM units in Attribute viewer and
'                 Drawing Area. This is just a way to identify the attribute value units.
'
'                 The UBound of Array attribute is always two.
'                   Array(1) - Unit Type(igUnitDistance,igUnitArea,igUnitAngle,igUnitMass)'
'                   Array(2) - Double Value'
'
'Description:   used in  AddSymbol() method..basically for Multiple margin case
'------------------------------------------------------------------------------
Private Function GetAttributeValue(oAttributeSets As AttributeSets, sAttrName As String) As String
    Const METHOD = "GetAttributeValue"
    On Error GoTo ErrorHandler
    
    Dim oAttribute      As rad2d.Attribute
    Dim oAttributeSet   As AttributeSet
    Dim dblDbu          As Double
    Dim i               As Long
    Dim j               As Long
    Dim arrDblValues()  As Double
    
    If Not oAttributeSets Is Nothing Then
        For i = 1 To oAttributeSets.Count
            Set oAttributeSet = oAttributeSets.Item(i)
            If Not oAttributeSet Is Nothing Then
                For j = 1 To oAttributeSet.Count
                    Set oAttribute = oAttributeSet.Item(j)
                    
                    If Not oAttribute Is Nothing And oAttribute.Name = sAttrName Then
                        If oAttribute.ValueIsArray Then 'If values are stored in array
                            oAttribute.GetArrayValue arrDblValues
                            If UBound(arrDblValues) > 1 Then
                                Select Case arrDblValues(1)
                                Case DISTANCE_CENTIMETER, DISTANCE_CHAIN, DISTANCE_FOOT, DISTANCE_FURLONG, DISTANCE_HUNDREDTH, DISTANCE_INCH, DISTANCE_KILOMETER, DISTANCE_LINK, DISTANCE_METER, DISTANCE_MILE, DISTANCE_MILLIMETER, DISTANCE_NANOMETER, DISTANCE_POINT, DISTANCE_POLE, DISTANCE_ROD, DISTANCE_TENTH, DISTANCE_THOUSANDTH, DISTANCE_YARD
                                   dblDbu = m_oUnitsofMeasure.ConvertUnitToDbu(UNIT_DISTANCE, arrDblValues(2), arrDblValues(1))
                                   GetAttributeValue = "" & dblDbu
                                   'm_oUnitsofMeasure.FormatUnit UNIT_DISTANCE, dblDbu, GetAttributeValue
                                Case ANGLE_DEGREE, ANGLE_MINUTE, ANGLE_SECOND, ANGLE_GRADIAN, ANGLE_RADIAN, ANGLE_REVOLUTION
                                   dblDbu = m_oUnitsofMeasure.ConvertUnitToDbu(UNIT_ANGLE, arrDblValues(2), arrDblValues(1))
                                   GetAttributeValue = "" & dblDbu
                                   'm_oUnitsofMeasure.FormatUnit UNIT_ANGLE, dblDbu, GetAttributeValue
                                Case AREA_ACRE, AREA_HECTARE, AREA_SQUARE_ACRE, AREA_SQUARE_CENTIMETER, AREA_SQUARE_FOOT, AREA_SQUARE_INCH, AREA_SQUARE_KILOMETER, AREA_SQUARE_METER, AREA_SQUARE_MILE, AREA_SQUARE_MILLIMETER, AREA_SQUARE_YARD
                                   dblDbu = m_oUnitsofMeasure.ConvertUnitToDbu(UNIT_AREA, arrDblValues(2), arrDblValues(1))
                                   GetAttributeValue = "" & dblDbu
                                   'm_oUnitsofMeasure.FormatUnit UNIT_AREA, dblDbu, GetAttributeValue
                                Case MASS_GRAIN, MASS_GRAM, MASS_KILOGRAM, MASS_LONG_TON, MASS_MEGAGRAM, MASS_METRIC_TON, MASS_MILLIGRAM, MASS_OUNCE, MASS_POUND_MASS, MASS_SHORT_TON, MASS_SLINCH, MASS_SLUG
                                   dblDbu = m_oUnitsofMeasure.ConvertUnitToDbu(UNIT_MASS, arrDblValues(2), arrDblValues(1))
                                   GetAttributeValue = "" & dblDbu
                                   'm_oUnitsofMeasure.FormatUnit UNIT_MASS, dblDbu, GetAttributeValue
                                End Select
                            End If
                        Else 'String
                            GetAttributeValue = oAttribute.Value
                        End If
                        GoTo CleanUp
                    End If
                Next
            End If
        Next
    End If
CleanUp:
    Set oAttribute = Nothing
    Set oAttributeSet = Nothing
    Erase arrDblValues
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

''**************************************************************************************
'' Routine      : GetSymbolSharePath
'' Abstract     : This method will return the SymbolShare Path ( like M:\CatalogData\Symbols).
''                Assumption: There will be Structmanufacturing folder under symbols directory. If doesn't exist
''                Get the LogicalSymbolPath registry key (This will be applicable for developer build)
''**************************************************************************************
Public Function GetSymbolSharePath() As String
    Const METHOD = "GetSymbolLocationPath"
    On Error GoTo ErrorHandler
    
    Dim oContext            As IJContext
    Dim oFileSysObj         As New FileSystemObject
    Dim oInstallDir         As S3DInstallInfo.CS3DInstallInfo
    Dim strSymbolShare      As String
    Dim strContextString    As String
    
    strContextString = "OLE_SERVER"
    
    'Get IJContext
    Set oContext = GetJContext()
    Set oInstallDir = New S3DInstallInfo.CS3DInstallInfo
    
    If Not oContext Is Nothing Then
        'Get the Symbol Share
        strSymbolShare = oContext.GetVariable(strContextString)
    
        'Append StructManufacturing
        strSymbolShare = strSymbolShare & PATH_SYMBOL_SHARE_EXTENSION
    
        'Check whether the folder exists if not get the Registry key Value
        If oFileSysObj.FolderExists(strSymbolShare) Then
            GetSymbolSharePath = strSymbolShare
        Else
            GetSymbolSharePath = oInstallDir.GetInstallDir(GAPPNAME_STRUCTURALMANUFACTURING) & PATH_DEV_SYMBOL_EXTENSION
        End If
    Else
        GetSymbolSharePath = oInstallDir.GetInstallDir(GAPPNAME_STRUCTURALMANUFACTURING) & PATH_DEV_SYMBOL_EXTENSION
    End If
    Set oContext = Nothing
    Set oFileSysObj = Nothing
    Set oInstallDir = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

Private Function GetGlobalUpsideDirection(oObject As Object) As String
    Const METHOD = "GetGlobalUpsideDirection"
    On Error GoTo ErrorHandler
    
    If oObject Is Nothing Then
        Exit Function
    End If
    
    Dim oPartSupport As IJPartSupport

    Dim oPlatePart As IJPlatePart
    Dim oPlatePartSupport As IJPlatePartSupport
    Dim oMfgPlate    As IJMfgPlatePart
        
    Dim oProfilePart As IJProfilePart
    Dim oMfgProfile As IJMfgProfilePart
        
    Dim oSurfaceBody As IJSurfaceBody
    
    If TypeOf oObject Is IJMfgPlatePart Then
        Set oMfgPlate = oObject
        oMfgPlate.GetDetailedPart oPlatePart
        Set oPartSupport = New PlatePartSupport
        Set oPartSupport.Part = oPlatePart
        Set oPlatePartSupport = oPartSupport
        
        If oMfgPlate.Upside = BaseSide Then
            oPlatePartSupport.GetSurface PlateBaseSide, oSurfaceBody
        Else
            oPlatePartSupport.GetSurface PlateOffsetSide, oSurfaceBody
        End If
    ElseIf TypeOf oObject Is IJMfgProfilePart Then
        Set oMfgProfile = oObject
        Dim oDetailedPart As Object
        Dim eProfileUpside As JXSEC_CODE
        
        Set oDetailedPart = oMfgProfile.GetDetailedPart
        Set oProfilePart = oDetailedPart
        If TypeOf oDetailedPart Is IJProfilePart Then
            Dim oSDProfileWrapper As New StructDetailObjects.ProfilePart
            Set oSDProfileWrapper.Object = oDetailedPart
        ElseIf TypeOf oDetailedPart Is IJStructProfilePart Then
            Dim oSDMemberWrapper As New StructDetailObjects.MemberPart
            Set oSDMemberWrapper.Object = oDetailedPart
        End If
                
        eProfileUpside = oMfgProfile.Upside
        If TypeOf oDetailedPart Is IJProfilePart And Not oSDProfileWrapper Is Nothing Then
            If eProfileUpside = JXSEC_WEB_LEFT Then
                Set oSurfaceBody = oSDProfileWrapper.SubPort(JXSEC_WEB_LEFT).Geometry
            ElseIf eProfileUpside = JXSEC_WEB_RIGHT Then
                Set oSurfaceBody = oSDProfileWrapper.SubPort(JXSEC_WEB_RIGHT).Geometry
            End If
        ElseIf TypeOf oDetailedPart Is IJStructProfilePart And Not oSDMemberWrapper Is Nothing Then
            If eProfileUpside = JXSEC_WEB_LEFT Then
                Set oSurfaceBody = oSDMemberWrapper.SubPort(JXSEC_WEB_LEFT).Geometry
            ElseIf eProfileUpside = JXSEC_WEB_RIGHT Then
                Set oSurfaceBody = oSDMemberWrapper.SubPort(JXSEC_WEB_RIGHT).Geometry
            End If
        End If
    End If
    
    If Not oSurfaceBody Is Nothing Then

        Dim oSurface As IJSurface
        Dim oNormal As IJDVector
        Dim oNormal2 As IJDVector
        Set oNormal2 = New DVector
        Dim oCG As IJDPosition
        Dim oPoint As IJDPosition
                                                
        oSurfaceBody.GetCenterOfGravity oCG
        Dim oMfgMGHelper As IJMfgMGHelper
        Set oMfgMGHelper = New MfgMGHelper
        
        oMfgMGHelper.ProjectPointOnSurfaceBody oSurfaceBody, oCG, oPoint, oNormal2
        
        oSurfaceBody.GetNormalFromPosition oPoint, oNormal
           
        Dim strGlobalUpside As String
        strGlobalUpside = "unknown"
        If Abs(oNormal.x) > Abs(oNormal.y) And Abs(oNormal.x) > Abs(oNormal.Z) Then 'x-dir is largest
            If oNormal.x > 0 Then
                strGlobalUpside = "fore"
            Else
                strGlobalUpside = "aft"
            End If
        ElseIf Abs(oNormal.y) > Abs(oNormal.Z) Then ' y-dir is largest
            Dim oMin As IJDPosition
            Dim oMax As IJDPosition
            Dim dMinX As Double
            Dim dMinY As Double
            Dim dMinZ As Double
            Dim dMaxX As Double
            Dim dMaxY As Double
            Dim dMaxZ As Double
            
            If Not oPlatePart Is Nothing Then
                oMfgMGHelper.GetRange oPlatePart, oMin, oMax
            ElseIf Not oProfilePart Is Nothing Then
                oMfgMGHelper.GetRange oProfilePart, oMin, oMax
            End If
                    
            oMin.Get dMinX, dMinY, dMinZ
            oMax.Get dMaxX, dMaxY, dMaxZ
            If dMinY <= 0 And dMaxY >= 0 Then
                If oNormal.y > 0 Then
                    strGlobalUpside = "port"
                Else
                    strGlobalUpside = "starboard"
                End If
            ElseIf dMaxY <= 0 Then
                If oNormal.y > 0 Then
                    strGlobalUpside = "in"
                Else
                    strGlobalUpside = "out"
                End If
            ElseIf dMinY >= 0 Then
                If oNormal.y > 0 Then
                    strGlobalUpside = "out"
                Else
                    strGlobalUpside = "in"
                End If
            End If
        Else                                        ' z-dir is largest
            If oNormal.Z > 0 Then
                strGlobalUpside = "upper"
            Else
                strGlobalUpside = "lower"
            End If
        End If
        Set oSurface = Nothing
        Set oNormal = Nothing
        Set oPoint = Nothing
    End If
    Set oPlatePartSupport = Nothing
    Set oPartSupport = Nothing
    Set oPlatePart = Nothing
    
    GetGlobalUpsideDirection = strGlobalUpside
    
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    GetGlobalUpsideDirection = "Unknown"
    GoTo CleanUp
End Function
 
Private Function GetPartStatus(ByVal oPart As Object) As String
    On Error GoTo ErrorHandler
    
    GetPartStatus = "Unknown"

    Dim oMfgChild           As IJMfgChild
    Dim oMfgParent          As IUnknown
    Dim ePartToDoType       As ToDoType
    Dim oMfgUpdateService   As IJMfgUpdateService
    Dim bParentStatus       As Boolean
    
    Set oMfgUpdateService = New MfgUpdateService
    Set oMfgChild = oPart
                    
    Set oMfgParent = oMfgChild.GetParent
    bParentStatus = IsObjectInTODORecord(oMfgParent, ePartToDoType)
                         
    If bParentStatus Then
        If ePartToDoType = TODO_ERROR Then
            GetPartStatus = "In Error"
        ElseIf ePartToDoType = TODO_WARNING Then
            GetPartStatus = "Warning"
        ElseIf ePartToDoType = TODO_OUT_OF_DATE Then
            GetPartStatus = "Out of Date"
        End If
    Else
        If oMfgUpdateService.IsParentPartLight(oPart) Then
            GetPartStatus = "Light Part"
        Else
            GetPartStatus = "Up To Date"
        End If
    End If
CleanUp:
    Set oMfgChild = Nothing
    Set oMfgParent = Nothing
    Set oMfgUpdateService = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetMfgStatus(ByVal oPart As Object) As String
    On Error GoTo ErrorHandler
    
    GetMfgStatus = "Unknown"
    
    Dim eMfgToDoType    As ToDoType
    Dim bMfgObjStatus   As Boolean
    Dim lStatus         As Long
    bMfgObjStatus = IsObjectInTODORecord(oPart, eMfgToDoType)
    Dim eMfgObjectStatus As MfgObjectStatus
    If bMfgObjStatus = False Then
        Dim oMfgUpdateService As IJMfgUpdateService
        Set oMfgUpdateService = New MfgUpdateService
        eMfgObjectStatus = oMfgUpdateService.IsUptoDateWithDependents(oPart)
        If eMfgObjectStatus = OutOfDate Then
            GetMfgStatus = "Out of Date"
        ElseIf eMfgObjectStatus = UpToDate Then
            GetMfgStatus = "Up to Date"
        ElseIf eMfgObjectStatus = CanBeDeleted Then
            GetMfgStatus = "Can Be Deleted"
        End If
        Set oMfgUpdateService = Nothing
    Else
        GetMfgStatus = "In To Do List"
    End If
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function
 
Private Function GetSurvivabilityStatus(ByVal oPart As Object) As String
On Error GoTo ErrorHandler
    
    GetSurvivabilityStatus = "Unknown"

    Dim res As enumMfgState
    Dim oMfgSurvivability As IJMfgSurvivability
    Set oMfgSurvivability = oPart
    res = oMfgSurvivability.MfgState
    Select Case res
    Case eMfgGenerated
        GetSurvivabilityStatus = "Generated"
    Case eMfgAnnotated
        GetSurvivabilityStatus = "Annotated"
    Case eMfgEdited
        GetSurvivabilityStatus = "Edited"
    End Select
    Set oMfgSurvivability = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
End Function

''**************************************************************************************
'' Routine      : IsObjectInTODORecord
'' Abstract     : Returns True with To Do Type when input object is in TO Do List.
''**************************************************************************************
Private Function IsObjectInTODORecord(oObject As Object, eToDoType As ToDoType) As Boolean
    Const METHOD = "IsObjectInTODORecord"
    On Error GoTo ErrorHandler
  
    Dim oAssocRelation  As IJDAssocRelation
    Dim Collection      As Object 'IJDRelationshipCol
    Dim RelCol          As IJDRelationshipCol
    Dim oToDoRecord     As IJToDoRecord
    Dim oRelationShip   As IJDRelationship
 
    Const IID_IJDObject As String = "{05F65FD1-57B3-11D1-93FC-0060973D4777}"

    IsObjectInTODORecord = False
    Set oAssocRelation = oObject
    Set Collection = oAssocRelation.CollectionRelations(IID_IJDObject, "toErrorList")
    Set RelCol = Collection
    Dim i As Integer
    For i = 1 To RelCol.Count
        Set oRelationShip = RelCol.Item(i)
        Set oToDoRecord = oRelationShip.Origin
        
        If (oToDoRecord.ToDoType = TODO_ERROR Or oToDoRecord.ToDoType = TODO_WARNING Or oToDoRecord.ToDoType = TODO_OUT_OF_DATE) Then
            eToDoType = oToDoRecord.ToDoType
            IsObjectInTODORecord = True
            Exit For
        End If
    Next
    
CleanUp:
    Set RelCol = Nothing
    Set Collection = Nothing
    Set oAssocRelation = Nothing
    Set oRelationShip = Nothing
    Set oToDoRecord = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

''**************************************************************************************
'' Routine      : GetAllowablesMgr
'' Abstract     : Returns a comma delimited list of allowables obtained from the catalog.
''              : The is the rule implementation to get the Allowables. Acts as a wrapper to call
''              : the other functions needed to get the allowables for a given property.
'' NOTE         : The oMfgObj input Object is the current Manufacturing object. This would normally be required,
''              : but if the user loads from an XML file, then there will never be any manufacturing object.
''              : So even if this needs to be required, it will never be required, because it can always be nothing,
''              : unless we want to disable the ability to load from an XML file, which will also never happen because
''              : of end-user expectations.
''**************************************************************************************
Public Function IJMfgPartMonitorInfo_GetAllowablesMgr(ByVal oPropElem As Object, ByVal sPartType As String, ByVal oPartObj As Object, ByVal oParentAttrElem As Object, ByVal oParentParentAttrElem As Object) As String
    Const METHOD = "GetAllowablesMgr"
    On Error GoTo ErrorHandler
    
    Dim oMatTypeElement             As IXMLDOMElement
    Dim oMatGradeElement            As IXMLDOMElement
    Dim oXSectionTypeElement        As IXMLDOMElement
    Dim i                           As Long
    Dim sUnitAttr                   As String
    Dim sUnitTypeAttr               As String
    Dim sMatTypeQuery               As String
    Dim sMatGradeQuery              As String
    Dim sXSecTypeQuery              As String
    Dim sXSecSizeQuery              As String
    Dim sSpecA                      As String
    Dim sSpecB                      As String
    'Dim sSpecC                      As String
    'Dim sSpecD                      As String
    Dim sAttrName                   As String
    Dim sTempAllowables             As String
    Dim sRetAllowables              As String
    Dim sAllowable                  As String
    Dim sAllowables()               As String
    
    IJMfgPartMonitorInfo_GetAllowablesMgr = "" 'Default is an empty string, in case we have any problems.
    
    If oPropElem Is Nothing Then GoTo CleanUp
    If oPropElem.xml = "" Then GoTo CleanUp
    
    '//PROPERTY[@NAME="SMS_PART_INFO||MATERIAL_TYPE"]
    '//PROPERTY[@NAME="SMS_PART_INFO||MATERIAL_GRADE"]
    '//PROPERTY[@NAME="SMS_PART_INFO||XSECTION_TYPE" or @NAME="SMS_BUILTUP||XSECTION_TYPE"]
    '//PROPERTY[@NAME="SMS_PART_INFO||XSECTION_TYPE" or @NAME="SMS_BUILTUP||XSECTION_TYPE"]
    sMatTypeQuery = "//PROPERTY[@NAME=" & """SMS_PART_INFO||MATERIAL_TYPE""" & "]"
    sMatGradeQuery = "//PROPERTY[@NAME=" & """SMS_PART_INFO||MATERIAL_GRADE""" & "]"
    sXSecTypeQuery = "//PROPERTY[@NAME=" & """SMS_PART_INFO||XSECTION_TYPE""" & _
                        " or @NAME=" & """SMS_BUILTUP||XSECTION_TYPE""" & "]"
    sXSecSizeQuery = "//PROPERTY[@NAME=" & """SMS_PART_INFO||XSECTION_SIZE""" & _
                        " or @NAME=" & """SMS_BUILTUP||XSECTION_SIZE""" & "]"
    If InStr(1, oPropElem.xml, "NAME") > 0 Then
        sAttrName = oPropElem.getAttribute("NAME")
    End If
    'this is the case that no ProgID was provided, but the user specified that they wanted the allowables list
    'to come from the database, or the user specified that they wanted the allowables list to come from the Database,
    'and also to be accumulated with the list of allowables from the XML file as well.
    'We can check the Property System Name, against the list of internally implemented database allowables accessors.
    Select Case sAttrName
        Case "SMS_PART_INFO||MATERIAL_GRADE"
            'Error checking internal to the function for this input.
            'Required: Need to set sSpecA to the value of the Material Type from the oParentAttrElem!
            If Not oParentAttrElem Is Nothing Then
                Set oMatTypeElement = oParentAttrElem.selectSingleNode(sMatTypeQuery)
            End If
            If Not oMatTypeElement Is Nothing Then
                sSpecA = oMatTypeElement.getAttribute("VALUE")
            End If
            If sSpecA <> "" Then
                sTempAllowables = GetMaterialGradesMgr(sPartType, sSpecA)
            End If
        Case "SMS_PART_INFO||MATERIAL_TYPE"
            sTempAllowables = GetMaterialTypesMgr(sPartType)
        Case "SMS_PLATE||THICKNESS", "SMS_MARKING_PART_LOCATION||THICKNESS"
            'Required: Need to set sSpecA to the value of the Material Type from the oParentAttrElem!
            'Required: Need to set sSpecB to the value of the Material Grade from the oParentAttrElem!
            If Not oParentAttrElem Is Nothing Then
                Set oMatTypeElement = oParentAttrElem.selectSingleNode(sMatTypeQuery)
            End If
            If Not oMatTypeElement Is Nothing Then
                sSpecA = oMatTypeElement.getAttribute("VALUE")
            End If
            If Not oParentAttrElem Is Nothing Then
                Set oMatGradeElement = oParentAttrElem.selectSingleNode(sMatGradeQuery)
            End If
            If Not oMatGradeElement Is Nothing Then
                sSpecB = oMatGradeElement.getAttribute("VALUE")
            End If
            If sSpecA <> "" And sSpecB <> "" Then
                sTempAllowables = GetMaterialThicknessMgr(sPartType, sSpecA, sSpecB)
            End If
        Case "SMS_PART_INFO||XSECTION_TYPE", "SMS_BUILTUP||XSECTION_TYPE"
            'Required: Need to get the reference standard from the Manufacturing Object.
            'NOTE: If there is no manufacturing object, such as loaded from XML, then we will need to rely only
            'on the Allowables from the XML, if there is any. The default does not have any!
            If Not oPartObj Is Nothing Then
                sTempAllowables = GetXSectionTypesMgr(sPartType, oPartObj)
            End If
        Case "SMS_BUILTUP||XSECTION_SIZE", "SMS_PART_INFO||XSECTION_SIZE"
            'Required: Need to set sSpecA to the value of the XSection_Type from the oParentAttrElem!
            If Not oParentAttrElem Is Nothing Then
                Set oXSectionTypeElement = oParentAttrElem.selectSingleNode(sXSecTypeQuery)
            End If
            If Not oXSectionTypeElement Is Nothing Then
                sSpecA = oXSectionTypeElement.getAttribute("VALUE")
            End If
            'NOTE: If the user hasn't picked an XSection Type, or the value is not set,
            'then there is nothing we can do! The Allowables will be empty,
            'although the user can still type something in.
            If sSpecA <> "" Then
                sTempAllowables = GetXSectionSizesMgr(sPartType, sSpecA, oPartObj)
            End If
    End Select
    'Still need to check and see if there is a unit conversion impact. We've got access to the entire Property Node
    'Ensure that we have a list of allowables that we can work with, it could be empty after-all.
    If InStr(1, oPropElem.xml, "UNIT") > 0 And sTempAllowables <> "" Then
        sUnitAttr = oPropElem.getAttribute("UNIT")
        If InStr(1, oPropElem.xml, "UNIT_TYPE") > 0 Then
            sUnitTypeAttr = oPropElem.getAttribute("UNIT_TYPE")
        End If
        If sUnitAttr <> "345" And sUnitAttr <> "0" _
        And sUnitAttr <> "" And sUnitAttr <> vbNullString _
        And sUnitTypeAttr <> "0" And sUnitTypeAttr <> "" And sUnitTypeAttr <> vbNullString Then
            'We have some unit conversion to do!
            'First we need to parse the Allowables back out.
            sAllowables = Split(sTempAllowables, ",")
            'Make sure we have something we can work with!
            If UBound(sAllowables) > 0 And IsNumeric(sUnitTypeAttr) = True And IsNumeric(sUnitAttr) = True Then
                For i = LBound(sAllowables) To UBound(sAllowables)
                    'Go through each allowable and convert it from database units to the storage units.
                    'The grid view will be responsible for converting the allowable values from storage units
                    'to display units! Just like it does for everything else!
                    'Make sure that each of the allowables is a numerical value
                    If IsNumeric(sAllowables(i)) = True Then
                        sAllowable = ManageValueUnitCheck(CDbl(sUnitTypeAttr), sAllowables(i), , _
                                                        True, True, True, , CDbl(sUnitAttr), , False, False)
                    Else
                        sAllowable = sAllowables(i)
                    End If
                    If i = LBound(sAllowables) Then
                        sRetAllowables = sAllowable
                    Else
                        sRetAllowables = sRetAllowables & "," & sAllowable
                    End If
                    sAllowable = ""
                Next i
                sTempAllowables = sRetAllowables
            End If
        End If
    End If
    IJMfgPartMonitorInfo_GetAllowablesMgr = sTempAllowables
CleanUp:
    Set oMatTypeElement = Nothing
    Set oMatGradeElement = Nothing
    Set oXSectionTypeElement = Nothing
    Erase sAllowables
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

Private Function GetMaterialTypesMgr(Optional sPartType As String = "") As String
    Const METHOD = "GetMaterialTypesMgr"
    On Error GoTo ErrorHandler
    
    'M:\ShipRefData\Middle\Bin\Debug\MoldedFormSpecMgr.dll
    Dim oSpecMgr            As IJDMoldedFormSpecMgr
    Dim oMoldedFormSpecs    As IJDCollection
    Dim oMoldedFormSpec     As IJDMoldedFormSpec
    Dim oIJDStructQuery2    As IJDStructQuery2
    Dim filter              As eMaterialFilter
    Dim i                   As Long
    Dim strAttrType()       As String
    
    Set oSpecMgr = New JMoldedFormSpecMgr
    
    'Get all the specs in the catalog
    Set oMoldedFormSpecs = oSpecMgr.GetSpecs
    
    'Get the first spec i.e., Spec_A(default)
    Set oMoldedFormSpec = oMoldedFormSpecs.Item(1)
    
    'NOTE: We may not be loading for plate, so this might need to be adjusted in the future!
    'We may need an input to control this. Added a parameter to control it, just in case.
    If sPartType = "PLATE" Then
        filter = MaterialForPlate
    ElseIf sPartType = "PROFILE" Then
        filter = MaterialForProfile
    Else 'Default back to plate
        filter = MaterialForPlate
    End If
    
    Set oIJDStructQuery2 = oMoldedFormSpec
    
    'Get Material types of plate
    oIJDStructQuery2.GetMaterials filter, strAttrType
    
    For i = 0 To UBound(strAttrType)
        If i = 0 Then
            GetMaterialTypesMgr = strAttrType(i)
        Else
            GetMaterialTypesMgr = GetMaterialTypesMgr & "," & strAttrType(i)
        End If
    Next i
CleanUp:
    Set oSpecMgr = Nothing
    Set oMoldedFormSpecs = Nothing
    Set oMoldedFormSpec = Nothing
    Set oIJDStructQuery2 = Nothing
    Erase strAttrType
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

Private Function GetMaterialGradesMgr(sPartType As String, sMaterialType As String) As String
    Const METHOD = "GetMaterialGradesMgr"
    On Error GoTo ErrorHandler
    
    Dim oSpecMgr            As IJDMoldedFormSpecMgr
    Dim oMoldedFormSpecs    As IJDCollection
    Dim oMoldedFormSpec     As IJDMoldedFormSpec
    Dim oIJDStructQuery2    As IJDStructQuery2
    Dim filter              As eMaterialFilter
    Dim i                   As Long
    Dim strArrGrade()       As String
    
    Set oSpecMgr = New JMoldedFormSpecMgr
    
    'Get all the specs in the catalog
    Set oMoldedFormSpecs = oSpecMgr.GetSpecs
    
    'Get the first spec i.e., Spec_A(default)
    Set oMoldedFormSpec = oMoldedFormSpecs.Item(1)
    
    If sPartType = "PLATE" Then
        filter = MaterialForPlate
    ElseIf sPartType = "PROFILE" Then
        filter = MaterialForProfile
    Else 'Default back to plate
        filter = MaterialForPlate
    End If
    
    Set oIJDStructQuery2 = oMoldedFormSpec
    
    'Get Material Grade collection from Material Type
    oIJDStructQuery2.GetGradesForMaterial filter, sMaterialType, strArrGrade
    
    For i = 0 To UBound(strArrGrade)
        If i = 0 Then
            GetMaterialGradesMgr = strArrGrade(i)
        Else
            GetMaterialGradesMgr = GetMaterialGradesMgr & "," & strArrGrade(i)
        End If
    Next i
CleanUp:
    Set oIJDStructQuery2 = Nothing
    Set oMoldedFormSpec = Nothing
    Set oMoldedFormSpecs = Nothing
    Set oSpecMgr = Nothing
    Erase strArrGrade
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

Private Function GetMaterialThicknessMgr(sPartType As String, sMaterialType As String, sMaterialGrade As String) As String
    Const METHOD = "GetMaterialThicknessMgr"
    On Error GoTo ErrorHandler
    
    Dim oSpecMgr            As IJDMoldedFormSpecMgr
    Dim oMoldedFormSpecs    As IJDCollection
    Dim oMoldedFormSpec     As IJDMoldedFormSpec
    Dim oIJDStructQuery2    As IJDStructQuery2
    Dim filter              As eMaterialFilter
    Dim i                   As Long
    Dim strThickness        As String
    Dim dblArrThickness()   As Double
    
    Set oSpecMgr = New JMoldedFormSpecMgr
    
    'Get all the specs in the catalog
    Set oMoldedFormSpecs = oSpecMgr.GetSpecs
    
    'Get the first spec i.e., Spec_A(default)
    Set oMoldedFormSpec = oMoldedFormSpecs.Item(1)
    
    If sPartType = "PLATE" Then
        filter = MaterialForPlate
    ElseIf sPartType = "PROFILE" Then
        filter = MaterialForProfile
    Else 'Default back to Plate
        filter = MaterialForPlate
    End If
    
    Set oIJDStructQuery2 = oMoldedFormSpec
    
    'Get the plate thickness collection
    
    oIJDStructQuery2.GetPlateThicknesses sMaterialType, sMaterialGrade, dblArrThickness
    
    For i = 0 To UBound(dblArrThickness)
        'Always format the unit for the Allowables list in the Part Editor!
        'In the case of thickness we just validate a range in the Part Editor.
        'CHECK_ATTRIBUTE_TYPE="CHECK_VALUE_RANGE"
        'Therefore what gets displayed in the combo box is just something the user can select,
        'the user should still be able to type in whatever number they want. It's just that the combo box
        'will give them the option to select a value that comes from the catalog, in this case.
        'However, we cannot do the unit formatting here, because these are the actual string values
        'that will show up in the combo box, and the user may have specified another unit in the XML definition!
        
        'We cannot do that unit conversion here, we need to do it one level up
        'where we will only have to implement it once!
        'Overrides will need to be handled in the over-ride section, which goes through the same functionality actually.
        'Just accumulate the values without formatting here!
        If i = 0 Then
            GetMaterialThicknessMgr = CStr(dblArrThickness(i))
        Else
            GetMaterialThicknessMgr = GetMaterialThicknessMgr & "," & CStr(dblArrThickness(i))
        End If
    Next i
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

Private Function GetXSectionTypesMgr(sPartType As String, oPartObj As Object) As String
    Const METHOD = "GetXSectionTypesMgr"
    On Error GoTo ErrorHandler
    
    Dim xSecColl                    As IJDCollection
    Dim oRefDataQuery               As RefDataMiddleServices.RefdataSOMMiddleServices
    Dim oPartClass                  As IJDPartClass
    Dim j                           As Long
    Dim sPartCrossSectionStandard   As String
    Dim sRetAllowables              As String
    
    sPartCrossSectionStandard = GetCrossSectionReferenceStandard(sPartType, oPartObj)
    If sPartCrossSectionStandard <> "" Then
        Set oRefDataQuery = New RefDataMiddleServices.RefdataSOMMiddleServices
        If Not oRefDataQuery Is Nothing Then
            Set xSecColl = oRefDataQuery.GetSectionTypes(sPartCrossSectionStandard)
        End If
    End If
    If Not xSecColl Is Nothing Then
        If xSecColl.Size > 0 Then
            For j = 1 To xSecColl.Size
                Set oPartClass = xSecColl.Item(j)
                If Not oPartClass Is Nothing Then
                    If j = 1 Then
                        sRetAllowables = oPartClass.Name
                    Else
                        sRetAllowables = sRetAllowables & "," & oPartClass.Name
                    End If
                End If
            Next j
        End If
    End If
    GetXSectionTypesMgr = sRetAllowables
CleanUp:
    Set xSecColl = Nothing
    Set oRefDataQuery = Nothing
        Set oPartClass = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

Private Function GetXSectionSizesMgr(sPartType As String, sXSectionType As String, oPartObj As Object) As String
    Const METHOD = "GetXSectionSizesMgr"
    On Error GoTo ErrorHandler
    Dim xSecColl                    As IJDCollection
    Dim oRefDataQuery As RefDataMiddleServices.StructServices
    
    Dim oCrossSection               As IJCrossSection
    Dim oAttribute                  As IJDAttribute
    Dim j                           As Long
    Dim sPartCrossSectionStandard   As String
    Dim sRetAllowables              As String
    
    If oPartObj Is Nothing Then GoTo CleanUp
    If sXSectionType = "" Or sPartType = "" Then GoTo CleanUp
    
    sPartCrossSectionStandard = GetCrossSectionReferenceStandard(sPartType, oPartObj)
    If sPartCrossSectionStandard <> "" Then
        Set oRefDataQuery = New RefDataMiddleServices.StructServices
        If Not oRefDataQuery Is Nothing Then
            Set xSecColl = oRefDataQuery.GetCrossSections(sPartCrossSectionStandard, sXSectionType)
            
        End If
    End If
    If Not xSecColl Is Nothing Then
        If xSecColl.Size > 0 Then
            For j = 1 To xSecColl.Size
                Set oCrossSection = xSecColl.Item(j)
                If Not oCrossSection Is Nothing Then
                    Set oAttribute = oCrossSection.Attributes.Item("SectionName")
                    If Not oAttribute Is Nothing Then
                        If j = 1 Then
                            sRetAllowables = oAttribute.Value
                        Else
                            sRetAllowables = sRetAllowables & "," & oAttribute.Value
                        End If
                    End If
                End If
            Next j
        End If
    End If
    GetXSectionSizesMgr = sRetAllowables
CleanUp:
        Set xSecColl = Nothing
        Set oRefDataQuery = Nothing
        Set oCrossSection = Nothing
        Set oAttribute = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

Private Function GetCrossSectionReferenceStandard(sPartType As String, oPartObj As Object) As String
    Const METHOD = "GetCrossSectionReferenceStandard"
    On Error GoTo ErrorHandler
    
    Dim oMfgProfileObj              As IJMfgProfilePart
    Dim oDetailObj                  As Object
    Dim oMfgProfilePartSupportObj   As IJProfilePartSupport
    Dim oMfgMemberObj               As ISPSMemberPartPrismatic
    Dim oPartSupport                As GSCADSDPartSupport.IJPartSupport
    Dim oCrossSectionObj            As IJCrossSection
    Dim oPartClass                  As IJDPartClass
    Dim oPartClassCrossSectionObj   As IJDCrossSectionPartClass
    Dim sPartCrossSectionStandard   As String
    
    If UCase(sPartType) = "PROFILE" Then
        Set oMfgProfileObj = oPartObj
        If Not oMfgProfileObj Is Nothing Then
            Set oDetailObj = oMfgProfileObj.GetDetailedPart
            Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
            If Not oPartSupport Is Nothing And Not oDetailObj Is Nothing Then
                Set oPartSupport.Part = oDetailObj
                Set oMfgProfilePartSupportObj = oPartSupport
            End If
            If Not oMfgProfilePartSupportObj Is Nothing Then
                oMfgProfilePartSupportObj.GetCrossSection oCrossSectionObj
            End If
            If Not oCrossSectionObj Is Nothing Then
                Set oPartClass = oCrossSectionObj.GetPartClass
            End If
            If Not oPartClass Is Nothing Then
                Set oPartClassCrossSectionObj = oPartClass
            End If
            If Not oPartClassCrossSectionObj Is Nothing Then
                sPartCrossSectionStandard = oPartClassCrossSectionObj.ReferenceStandard
            End If
        End If
    ElseIf UCase(sPartType) = "MEMBER" Then
        Set oMfgMemberObj = oPartObj
        If Not oMfgMemberObj Is Nothing Then
            sPartCrossSectionStandard = oMfgMemberObj.CrossSection.SectionStandard
        End If
    End If
    GetCrossSectionReferenceStandard = sPartCrossSectionStandard
CleanUp:
    Set oMfgProfileObj = Nothing
    Set oDetailObj = Nothing
    Set oMfgProfilePartSupportObj = Nothing
    Set oMfgMemberObj = Nothing
    Set oPartSupport = Nothing
    Set oCrossSectionObj = Nothing
    Set oPartClass = Nothing
    Set oPartClassCrossSectionObj = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    GoTo CleanUp
End Function

Private Function ManageValueUnitCheck(dblUnitType As Double, strInputVal As String, _
Optional strShortUnit As String = "", Optional bFormatUnit As Boolean = False, _
Optional bConv2SpecUnit As Boolean = False, Optional bWorkUnitless As Boolean = False, _
Optional dblSpecUnitIn As Double = 0, Optional dblSpecUnitOut As Double = 0, _
Optional dblPrecision As Integer = 0, Optional bEnforcePrecision As Boolean = True, _
Optional bParseInputValue As Boolean = True) As String
    Const METHOD = "ManageValueUnitCheck"
    On Error GoTo ErrorHandler
    Dim oUnitsOfMeasure     As IJUomVBInterface
    Dim dblParsedValue      As Double
    Dim strFormattedString   As String
    Dim bHandlingParseError As Boolean
    Dim bCaughtParseError   As Boolean
    
    Set oUnitsOfMeasure = New UnitsOfMeasureServicesLib.UomVBInterface
    bHandlingParseError = False
    bCaughtParseError = False
    
    If Not oUnitsOfMeasure Is Nothing Then
        bHandlingParseError = True
        If bParseInputValue = True Then
            If IsNumeric(strInputVal) = True Then
                dblParsedValue = oUnitsOfMeasure.ConvertUnitToDbu(dblUnitType, strInputVal, dblSpecUnitIn)
            ElseIf IsNumeric(strInputVal) = False Then
                dblParsedValue = oUnitsOfMeasure.ParseUnit(dblUnitType, strInputVal)
            End If
        Else
            If IsNumeric(strInputVal) = True Then
                'Only using this in the event that we know it's already in database units.
                dblParsedValue = CDbl(strInputVal)
            Else 'Not a numerical number value.
                'We could clean the input string value, and return something that can be converted to a numerical value.
                'But the way this is setup, we shouldn't actually get here if we never pass a formatted value and unit
                'string to this function and the dParseInputValue flag is set to false!
                'Going to leave this blank for now, and implement later if we find that it's needed.
                strFormattedString = "Parse Error"
            End If
        End If
        If bCaughtParseError = True Then
            strFormattedString = "error"
        Else
            If bFormatUnit = True And bWorkUnitless = False Then
                If dblSpecUnitOut <> 0 Then
                    oUnitsOfMeasure.FormatUnit dblUnitType, dblParsedValue, strFormattedString, , dblSpecUnitOut
                ElseIf dblSpecUnitOut = 0 Then
                    oUnitsOfMeasure.FormatUnit dblUnitType, dblParsedValue, strFormattedString
                End If
            Else
                If bConv2SpecUnit = True Then
                    If dblSpecUnitIn <> 0 Then
                        dblParsedValue = oUnitsOfMeasure.ConvertUnitToUnit(dblUnitType, dblParsedValue, _
                                                                            dblSpecUnitIn, dblSpecUnitOut)
                    ElseIf dblSpecUnitIn = 0 Then
                        dblParsedValue = oUnitsOfMeasure.ConvertDbuToUnit(dblUnitType, dblParsedValue, dblSpecUnitOut)
                    End If
                Else
                    dblParsedValue = oUnitsOfMeasure.ConvertDbuToUnit(dblUnitType, dblParsedValue, dblSpecUnitOut)
                End If
                strShortUnit = GetUnitShortString(dblUnitType, dblParsedValue, dblSpecUnitOut, True)
                If bWorkUnitless = False Then
                    strFormattedString = FormatWithNZerosMgr(dblParsedValue, dblPrecision, bEnforcePrecision) & strShortUnit
                Else
                    strFormattedString = FormatWithNZerosMgr(dblParsedValue, dblPrecision, bEnforcePrecision)
                End If
            End If
        End If
        ManageValueUnitCheck = strFormattedString
        bHandlingParseError = False
    End If
ClenaUp:
    Set oUnitsOfMeasure = Nothing
    Exit Function
ErrorHandler:
    If bHandlingParseError = False Then
        Err.Raise Err.Number
    Else
        bCaughtParseError = True
    End If
    Resume Next
End Function

Private Function GetUnitShortString(MyUnitType As Double, inVal As Double, _
Optional SpecUnit As Double = 0, Optional GetSpecUnit As Boolean = False) As String
    Const METHOD = "GetUnitShortString"
    On Error GoTo ErrorHandler
    Dim oUnitsOfMeasure As IJUomVBInterface
    Dim unitpos As Integer
    Dim tempStr As String 'Needed as a proxy to get the short unit from the Formatted string
    
    Set oUnitsOfMeasure = New UnitsOfMeasureServicesLib.UomVBInterface
    If Not oUnitsOfMeasure Is Nothing Then
        If GetSpecUnit = False Then
            oUnitsOfMeasure.FormatUnit MyUnitType, inVal, tempStr
        Else
            oUnitsOfMeasure.FormatUnit MyUnitType, inVal, tempStr, , SpecUnit
        End If
    End If
    
    'Find the first space between the value and the unit.
    unitpos = InStr(1, tempStr, " ")
    GetUnitShortString = Mid(tempStr, unitpos)
CleanUp:
    Set oUnitsOfMeasure = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

''**************************************************************************************
'' Routine      : FormatWithNZerosMgr
'' Abstract     : This method will add any number of zeros to a value to get the required precision
''              : for an output string.
'' NOTE         : If the string is longer than the required precision, then we have some options based
''              : on the boolean flag bEnforcePrecision which decides if the string should be truncated
''              : to the required precision, or if false then we return the full value of the string!
''**************************************************************************************
Private Function FormatWithNZerosMgr(dval As Double, Optional iPrecision As Integer = 0, _
Optional bEnforcePrecision As Boolean = True) As String
    Const METHOD = "FormatWithNZerosMgr"
    On Error GoTo ErrorHandler
    Dim lPlace As Long
    Dim sText As String
    'If iPrecision = 0 Then default to 2 digits of precision
    If iPrecision = 0 Then iPrecision = 2
    If dval = 0 Then
        sText = CStr(dval)
        sText = AddNZeros2String(sText, ".", iPrecision)
    Else
        sText = CStr(dval)
        'Search for position of decimal point
        lPlace = InStr(1, StrReverse(sText), ".", vbTextCompare)
        If bEnforcePrecision = True Then
            If lPlace > iPrecision Then
                'Need to truncate the string to the required precision.
                sText = Truncate2NDecPlace(dval, iPrecision)
            ElseIf lPlace = iPrecision Then
                'Don't need to do anything! Just return the string as is.
                GoTo CleanUp
            Else
                'Need to add some zeros to the string to get the required precision.
                sText = FormatWithNZeros(lPlace, iPrecision, sText)
            End If
        Else
            If lPlace < iPrecision Then
                'This is the only case that we need to actually add zeros to the string to get the required precision.
                'The only difference is that if it's the other way, we don't truncate it.
                sText = FormatWithNZeros(lPlace, iPrecision, sText)
            End If
        End If
    End If
CleanUp:
    FormatWithNZerosMgr = sText
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

Private Function FormatWithNZeros(lPlace As Long, iPrecision As Integer, sTextIn As String) As String
    Const METHOD = "FormatWithNZeros"
    On Error GoTo ErrorHandler
    Dim sText As String
    If lPlace = 0 Then
        sText = AddNZeros2String(sTextIn, ".", iPrecision)
    ElseIf lPlace = 1 Then
        sText = AddNZeros2String(sTextIn, "", iPrecision)
    ElseIf lPlace >= 2 Then
        sText = AddNZeros2String(sTextIn, "", (iPrecision + 1) - (lPlace - 1))
    Else 'It must be a negative number or something crazy, this is an error condition.
        Err.Raise 0, MODULE & "::" & METHOD, "Failure to get the position of the decimal point."
    End If
    FormatWithNZeros = sText
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

'********************************************************************
' Routine: AddNZeros2String
' Abstract: Adds "N" number of Zeros to the end of a given string, First time through the loop another string
' can be appended to the string, before adding other zeros.
' If there is an invalid input for iNZeros2Add, then we just return the input string, with no changes.
'********************************************************************
Private Function AddNZeros2String(sIn As String, sFirstStr As String, iNZeros2Add As Integer) As String
    Const METHOD = "AddNZeros2String"
    On Error GoTo ErrorHandler
    Dim i As Integer
    If iNZeros2Add >= 1 Then
        For i = 1 To iNZeros2Add
            If i = 1 Then
                AddNZeros2String = sIn & sFirstStr
            Else
                AddNZeros2String = AddNZeros2String & "0"
            End If
        Next i
    Else
        'There was an invalid input, just return the string the same as it was passed in.
        AddNZeros2String = sIn
    End If
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function

'********************************************************************
' Routine: Truncate2NDecPlace
' Abstract: Truncates to n-th significant figures,
' ensures that we always return n-number of digits past the decimal point. No more, no less
'********************************************************************
Private Function Truncate2NDecPlace(dblInput As Double, iSigFigDesired As Integer) As String
    Const METHOD = "Truncate2NDecPlace"
    On Error GoTo ErrorHandler
    Dim iSigFig As Integer
    Dim decPos As Integer
    Dim iSigFigDiff As Integer
    Dim i As Integer
    Dim sSigFig As String
    'We always need to truncate to the n-th significant digit.
    If dblInput = 0 Then
        Truncate2NDecPlace = AddNZeros2String("0.0", "", iSigFigDesired)
    ElseIf InStr(1, CStr(dblInput), ".") <> 0 Then
        decPos = InStr(1, CStr(dblInput), ".")
        iSigFig = Len(CStr(dblInput)) - decPos
        sSigFig = Right(CStr(dblInput), iSigFig)
        iSigFigDiff = iSigFigDesired - iSigFig
        If iSigFigDiff > 0 Then
            'We need to add some zeros
            Truncate2NDecPlace = AddNZeros2String(CStr(dblInput), "0", iSigFigDiff)
        ElseIf iSigFigDiff = 0 Then
            'Perfect!
            Truncate2NDecPlace = CStr(dblInput)
        ElseIf iSigFigDiff < 0 Then
            'Truncate some numbers off
            Truncate2NDecPlace = Left(CStr(dblInput), decPos + iSigFigDesired) 'Example: x.xxx
        Else
            'Not really sure what to do in this case! Definately an error condition!
            Err.Raise 1, METHOD, "Error Truncating to the desired significant figure."
        End If
    ElseIf InStr(1, CStr(dblInput), ".") = 0 Then
        Truncate2NDecPlace = AddNZeros2String(CStr(dblInput), ".0", iSigFigDesired)
    End If
    Exit Function
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Function
