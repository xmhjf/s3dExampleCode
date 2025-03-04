Attribute VB_Name = "PF_HgrGeneralUtil"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   PF_HgrPrimitives.bas
'   ProgID:         PF_HgrGeneralUtil.bas
'   Author:         Pelican Forge
'   Creation Date:  NA

'   Description:
'       This class contains common functionality that bridges the differences between SP3D and SupportModeler Code
'       It is intended to make conversion from SupMod libraries easier as well as to modularize common functionality
'
'   Change History:
'   dd.mmm.yyyy       who            change description
'   03.May.2006       SS             Applied fixes for DI 86724
'   03.May.2006       JRM            Apply fixes from code review (TR 95926)
'   16.May.2006      Ramya           DI 86721  Parameterize Database Queries in Symbols
'   23.May.2006      Bharath         DI 86722  ReadParametricData Upgrade
'   11.Jul.2006      SS              TR 98162 - Added new function GetPipeODwithoutInsulation()
'   25.July.2006     Ramya           TR 102076  Remove redundant check in GetNParamUsingParametricQuery & GetSParamUsingParamet
'   31.Oct.2006      IRK             TR 100431  Added AddInput, AddOutput and SetupInputsAndOutputsEx functions
'   Nov 06, 2006      SS                   TR 104793 Rigid rods and spring rod supports give asserts if HSSR steel is used.
'                                           Added a new function GetSupportingPropertyByRule
'   Feb 02, 2007     Prakash         DI#109787 : GetNParam and GetSParam should do case insensitive compare
'                                    Modified GetNParam(), GetSParam(), GetNParamUsingParametricQuery(), GetSParamUsingParametricQuery().
'   Feb 02, 2007     Prakash         DI#109789 : Need to share code file in VSS.
'                                    Merged the code into this single file which will be used by Rules/Symbols.
'   May 21, 2008     SS              TR 127345 - Added DoesExistSymbolAttribute function
'   Nov 29, 2010     Ramya           CR-CP-190404  Modified GetSupportingTypes, GetSupportingPropertyByRule to work with Place By Reference command
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "PF_HgrGeneralUtil"  'Used for error messages

Public Type Parameters
    name As String
    Type As DataTypeEnum
    Direction As ParameterDirectionEnum
    Length As Integer
    Value As Variant
End Type

'Data type to hold Inputs information
Public Type tInputs
    iCount As Long
    sInputName() As String
    sDataType() As String
End Type

'Data type to hold Outputs information
Public Type tOutputs
    iCount As Long
    sOutputName() As String
End Type

Dim PrevRecord As String
'Define attribute list
Private Type AttributeList
    AttributeName As String
    AttributeValue As Variant
End Type

Dim oAttributeList() As AttributeList
Dim mParameters() As Parameters
Dim bFirst As Boolean


' ---------------------------------------------------------------------------
' Name: SetupInputsAndOutputs
' Description: This function was created to modularize the code needed to generate the inputs and outputs that are required
'              to develop a SP3D symbol
' Example: any Symbol class file
'
' Inputs - pSymbolDefinition - IMSSymbolEntities.IJDSymbolDefinition, NUM_INPUTS - Integer, NUM_OUTPUTS - Integer, InputName() - String Array,
'          DataType() - String Array, OutputName() - String Array, m_progID - String
' Outputs - none
' ---------------------------------------------------------------------------
Public Function SetupInputsAndOutputs(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, _
        ByVal NUM_INPUTS As Integer, ByVal NUM_OUTPUTS As Integer, InputName() As String, _
        DataType() As String, OutputName() As String, m_progID As String)
Const METHOD = "SetupInputsAndOutputs"

    Dim strName() As String
    strName = Split(m_progID, ".")
    
    ' DI#109789 ; removed checks for Lisega and PSL
    'No Caching - Remove this before final build
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    On Error GoTo ErrorHandler

    ' Create a new input by new operator
    ' Commented the ReDim statement to set the array size to 20 and added the rule - If NUM_INPUTS > 20 use ReDim statement -
    ' - Fot JTS DI 86724

    'ReDim SymbolInput(1 To NUM_INPUTS) As IMSSymbolEntities.IJDInput
    Dim SymbolInput() As IMSSymbolEntities.IJDInput

    Call AssignArraySizeForInput(SymbolInput, NUM_INPUTS)

    Dim PCNum As IMSSymbolEntities.IJDParameterContent
    Set PCNum = New IMSSymbolEntities.DParameterContent
    PCNum.Type = igValue
    Dim PCTxt As IMSSymbolEntities.IJDParameterContent
    Set PCTxt = New IMSSymbolEntities.DParameterContent
    PCTxt.Type = igString

    Dim ii As Integer
    For ii = 1 To NUM_INPUTS
      Set SymbolInput(ii) = New IMSSymbolEntities.DInput
    Next
    Dim libDesc As New DLibraryDescription

    ' DI#109789 ; removing the checks for Lisega and PSL
    ' set the CMCacheForPart method in the definition
    Dim mCookie As Long
    Dim libCookie As Long

    'Specify "HgrSupHgrAssemblySymbols.ThreadTop" as the library.
    'This module implements the method "CMCacheForHgrPart"
    libDesc.name = "mySelfAsLib"
    libDesc.Type = imsLIBRARY_IS_ACTIVEX
    libDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    libDesc.Source = "HgrSupportSymbolUtilities.CustomMethods"
    pSymbolDefinition.IJDUserMethods.SetLibrary libDesc
    ' Get the lib/method cookie
    libCookie = libDesc.Cookie
    mCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCacheForHgrPart", libCookie)

    SymbolInput(1).name = "Part"
    SymbolInput(1).Description = "Part"
    SymbolInput(1).IJDInputStdCustomMethod.SetCMCache libCookie, mCookie

    'Set the input parameters
    Dim N As Integer
    Dim bFound As Boolean
    For N = 2 To NUM_INPUTS
        SymbolInput(N).name = InputName(N)
        SymbolInput(N).Description = InputName(N)
        SymbolInput(N).Properties = igINPUT_IS_A_PARAMETER

        If DataType(N) = "Double" Then
            PCNum.UomValue = 0.999999
            SymbolInput(N).DefaultParameterValue = PCNum
        ElseIf DataType(N) = "String" Then
            PCTxt.String = "No Value"
            SymbolInput(N).DefaultParameterValue = PCTxt
        Else
     'MsgBox "Unknown type passed to SetupInputsAndOutputs: " + DataType(N)
        End If
    Next

    Dim inputsProp As IMSDescriptionProperties
    inputsProp = pSymbolDefinition.IJDInputs.Property
    pSymbolDefinition.IJDInputs.Property = inputsProp Or igCOLLECTION_VARIABLE

    ' DI#109789 ; removing the checks for Lisega and PSL
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED



    ' Set the input to the definition
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = pSymbolDefinition

    For ii = 1 To NUM_INPUTS
      oInputs.SetInput SymbolInput(ii), ii
    Next

    ' Create the Simple Physical Representation
    ' =========================================
    ' Define the representation "SimplePhysical"
    Dim oRep1 As IMSSymbolEntities.IJDRepresentation
    Set oRep1 = New IMSSymbolEntities.DRepresentation
    Dim oOutputs As IMSSymbolEntities.IJDOutputs
    Set oOutputs = oRep1

    ' Create the output
    ' Commented the ReDim statement to set the array size to 20 and added the rule - If NUM_INPUTS > 20 use ReDim statement -
    ' - Fot JTS DI 86724

    'ReDim SimpleSymbolOutput(1 To NUM_OUTPUTS) As IMSSymbolEntities.IJDOutput
    Dim SimpleSymbolOutput() As IMSSymbolEntities.IJDOutput

    Call AssignArraySizeForOutput(SimpleSymbolOutput, NUM_OUTPUTS)

    For ii = 1 To NUM_OUTPUTS
        Set SimpleSymbolOutput(ii) = New IMSSymbolEntities.DOutput
        SimpleSymbolOutput(ii).name = OutputName(ii)
        SimpleSymbolOutput(ii).Description = OutputName(ii)
        SimpleSymbolOutput(ii).Properties = 0
        oOutputs.SetOutput SimpleSymbolOutput(ii)
        Set SimpleSymbolOutput(ii) = Nothing
    Next
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    oRep1.name = "Symbolic"
    oRep1.Description = "Symbolic Represntation of the 3d flexible"
    oRep1.RepresentationId = SimplePhysical 'define a aspect 0 (Simple_physical)
    oRep1.Properties = igREPRESENTATION_ISVBFUNCTION

    ' Set the representation to definition
    Dim oReps As IMSSymbolEntities.IJDRepresentations
    Set oReps = pSymbolDefinition
    oReps.SetRepresentation oRep1

    Dim oVbFuncSymbolicRep As IJDRepresentationEvaluation
    Set oVbFuncSymbolicRep = New DRepresentationEvaluation
    oVbFuncSymbolicRep.name = "Symbolic"
    oVbFuncSymbolicRep.Description = "script for the Symbolic representation"
    oVbFuncSymbolicRep.Properties = igREPRESENTATION_HIDDEN
    oVbFuncSymbolicRep.Type = igREPRESENTATION_VBFUNCTION
    oVbFuncSymbolicRep.ProgId = m_progID

    Dim oScripts As IMSSymbolEntities.IJDRepresentationEvaluations
    Set oScripts = pSymbolDefinition

    oScripts.AddRepresentationEvaluation oVbFuncSymbolicRep

'===================================================================
' DEFINE DetailPhysical REPRESENTATION
'===================================================================
    oRep1.name = "Detailed"
    oRep1.Description = "Detailed Represntation of the 3d flexible"
    oRep1.RepresentationId = DetailPhysical  'Detailed Physical
    oRep1.Properties = igREPRESENTATION_ISVBFUNCTION

    ' Set the representation to definition
    oReps.SetRepresentation oRep1

    ' Set the script associated to the Detailed representation
    Dim oVbFuncDetailedRep As DRepresentationEvaluation
    Set oVbFuncDetailedRep = New DRepresentationEvaluation
    oVbFuncDetailedRep.name = "Detailed"
    oVbFuncDetailedRep.Description = "script for the Detailed representation"
    oVbFuncDetailedRep.Properties = igREPRESENTATION_HIDDEN
    oVbFuncDetailedRep.Type = igREPRESENTATION_VBFUNCTION
    oVbFuncDetailedRep.ProgId = m_progID

    oScripts.AddRepresentationEvaluation oVbFuncDetailedRep

    Set oReps = Nothing
    Set oRep1 = Nothing
    Set oScripts = Nothing
    Set oVbFuncSymbolicRep = Nothing
    Set oVbFuncDetailedRep = Nothing
    Set PCNum = Nothing
    Set PCTxt = Nothing
    Set libDesc = Nothing
    Set oInputs = Nothing
    Set oOutputs = Nothing
    'Set oVbFuncMaintenanceRep = Nothing

'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function SetupInputsAndOutputsForCache(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, _
        ByVal NUM_INPUTS As Integer, ByVal NUM_OUTPUTS As Integer, InputName() As String, _
        DataType() As String, OutputName() As String, m_progID As String)
Const METHOD = "SetupInputsAndOutputsForCache"
    
    Dim oSymbolCache As New CustomCache

    NUM_INPUTS = NUM_INPUTS - 1

    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

    On Error GoTo ErrorHandler

    ' Create a new input by new operator
    ' Commented the ReDim statement to set the array size to 20 and added the rule - If NUM_INPUTS > 20 use ReDim statement -
    ' - Fot JTS DI 86724

    'ReDim SymbolInput(1 To NUM_INPUTS) As IMSSymbolEntities.IJDInput
    Dim SymbolInput() As IMSSymbolEntities.IJDInput

    Call AssignArraySizeForInput(SymbolInput, NUM_INPUTS)

    Dim PCNum As IMSSymbolEntities.IJDParameterContent
    Set PCNum = New IMSSymbolEntities.DParameterContent
    PCNum.Type = igValue
    Dim PCTxt As IMSSymbolEntities.IJDParameterContent
    Set PCTxt = New IMSSymbolEntities.DParameterContent
    PCTxt.Type = igString

    Dim ii As Integer
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = pSymbolDefinition
    
    oSymbolCache.SetupCustomCache pSymbolDefinition

    'Set the input parameters
    Dim N As Integer
    Dim bFound As Boolean
    For N = 1 To NUM_INPUTS
        Set SymbolInput(N) = New IMSSymbolEntities.DInput
        SymbolInput(N).name = InputName(N + 1)
        SymbolInput(N).Description = InputName(N + 1)
        SymbolInput(N).Properties = igINPUT_IS_A_PARAMETER

        If DataType(N + 1) = "Double" Then
            PCNum.UomValue = 0.999999
            SymbolInput(N).DefaultParameterValue = PCNum
        ElseIf DataType(N + 1) = "String" Then
            PCTxt.String = "No Value"
            SymbolInput(N).DefaultParameterValue = PCTxt
        Else
     'MsgBox "Unknown type passed to SetupInputsAndOutputs: " + DataType(N)
        End If

        oInputs.SetInput SymbolInput(N), N + 1
        Set SymbolInput(N) = Nothing
    Next

    ' Create the Simple Physical Representation
    ' =========================================
    ' Define the representation "SimplePhysical"
    Dim oRep1 As IMSSymbolEntities.IJDRepresentation
    Set oRep1 = New IMSSymbolEntities.DRepresentation
    Dim oOutputs As IMSSymbolEntities.IJDOutputs
    Set oOutputs = oRep1

    ' Create the output
    ' Commented the ReDim statement to set the array size to 20 and added the rule - If NUM_INPUTS > 20 use ReDim statement -
    ' - Fot JTS DI 86724

    'ReDim SimpleSymbolOutput(1 To NUM_OUTPUTS) As IMSSymbolEntities.IJDOutput
    Dim SimpleSymbolOutput() As IMSSymbolEntities.IJDOutput

    Call AssignArraySizeForOutput(SimpleSymbolOutput, NUM_OUTPUTS)

    For ii = 1 To NUM_OUTPUTS
        Set SimpleSymbolOutput(ii) = New IMSSymbolEntities.DOutput
        SimpleSymbolOutput(ii).name = OutputName(ii)
        SimpleSymbolOutput(ii).Description = OutputName(ii)
        SimpleSymbolOutput(ii).Properties = 0
        oOutputs.SetOutput SimpleSymbolOutput(ii)
        Set SimpleSymbolOutput(ii) = Nothing
    Next
    'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
    'different repIDs available.
    oRep1.name = "Symbolic"
    oRep1.Description = "Symbolic Represntation of the 3d flexible"
    oRep1.RepresentationId = SimplePhysical 'define a aspect 0 (Simple_physical)
    oRep1.Properties = igREPRESENTATION_ISVBFUNCTION

    ' Set the representation to definition
    Dim oReps As IMSSymbolEntities.IJDRepresentations
    Set oReps = pSymbolDefinition
    oReps.SetRepresentation oRep1

    Dim oVbFuncSymbolicRep As IJDRepresentationEvaluation
    Set oVbFuncSymbolicRep = New DRepresentationEvaluation
    oVbFuncSymbolicRep.name = "Symbolic"
    oVbFuncSymbolicRep.Description = "script for the Symbolic representation"
    oVbFuncSymbolicRep.Properties = igREPRESENTATION_HIDDEN
    oVbFuncSymbolicRep.Type = igREPRESENTATION_VBFUNCTION
    oVbFuncSymbolicRep.ProgId = m_progID

    Dim oScripts As IMSSymbolEntities.IJDRepresentationEvaluations
    Set oScripts = pSymbolDefinition

    oScripts.AddRepresentationEvaluation oVbFuncSymbolicRep

'===================================================================
' DEFINE DetailPhysical REPRESENTATION
'===================================================================
    oRep1.name = "Detailed"
    oRep1.Description = "Detailed Represntation of the 3d flexible"
    oRep1.RepresentationId = DetailPhysical  'Detailed Physical
    oRep1.Properties = igREPRESENTATION_ISVBFUNCTION

    ' Set the representation to definition
    oReps.SetRepresentation oRep1

    ' Set the script associated to the Detailed representation
    Dim oVbFuncDetailedRep As DRepresentationEvaluation
    Set oVbFuncDetailedRep = New DRepresentationEvaluation
    oVbFuncDetailedRep.name = "Detailed"
    oVbFuncDetailedRep.Description = "script for the Detailed representation"
    oVbFuncDetailedRep.Properties = igREPRESENTATION_HIDDEN
    oVbFuncDetailedRep.Type = igREPRESENTATION_VBFUNCTION
    oVbFuncDetailedRep.ProgId = m_progID

    oScripts.AddRepresentationEvaluation oVbFuncDetailedRep

    Set oReps = Nothing
    Set oRep1 = Nothing
    Set oScripts = Nothing
    Set oVbFuncSymbolicRep = Nothing
    Set oVbFuncDetailedRep = Nothing
    Set PCNum = Nothing
    Set PCTxt = Nothing
    Set oInputs = Nothing
    Set oOutputs = Nothing

'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: ASin
' Description: Calculate the Arc Sin of the input value
'
' Inputs - A - Double
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function ASin(A As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "ASin"

    'Inverse Sine Function
    If Abs(A) <> 1 Then
        ASin = Atn(A / Sqr(1 - A * A))
    Else
        ASin = 1.5707963267949 * Sgn(A)
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: ACos
' Description: Calculate the Arc Cos of the input value
'
' Inputs - A - Double
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function ACos(A As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "ACos"

    'Inverse Cosine Function
    If Abs(A) <> 1 Then
        ACos = 1.5707963267949 - Atn(A / Sqr(1 - A * A))
    ElseIf A = -1 Then
        ACos = 3.14159265358979
    ElseIf A = 1 Then
        ACos = 0
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: FIFtoMeter
' Description: Convert a string measurement that is feet-inch-fraction to Meters
'
' Inputs - S - String
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function FIFtoMeter(S As String) As Double
On Error GoTo ErrorHandler
Const METHOD = "FIFtoMeter"

    Dim numerator As String
    Dim denominator As String

    ' parse the string
    ' if there are no spaces, give an error message
    If InStr(S, " ") = 0 Then
 ' check if there is a quote
        If InStr(S, Chr(34)) > 0 Then
            If InStr(S, "/") = 0 Then
         ' this means it's just an inch value
                FIFtoMeter = Val(S) * 25.4 / 1000#
            Else
         ' if there is a slash, it is a fraction of an inch
                'numerator = Left(S, InStr(S, "/"))
                numerator = Mid(S, 1, InStr(S, "/")) ' DI#109789 ; changed to Mid as Left was giving compilation error because of enumtype 'left'
                denominator = Mid(S, InStr(S, "/") + 1)
                FIFtoMeter = (Val(numerator) / Val(denominator)) * 25.4 / 1000#
            End If
        Else
     ' may need to check for a single quote or a slash in the future, but I doubt it
            FIFtoMeter = 0#
        End If
    Else
        Dim firstParam As String
        Dim secondParam As String
        Dim thirdParam As String

 ' there is a space, so parse the first parameter out
        'firstParam = Left(S, InStr(S, " "))
        firstParam = Mid(S, 1, InStr(S, " ")) ' DI#109789 ; changed to Mid as Left was giving compilation error because of enumtype 'left'
        secondParam = Mid(S, InStr(S, " ") + 1)

        If InStr(secondParam, " ") = 0 Then
            If InStr(secondParam, "/") = 0 Then
         ' if there is one space and no slash, it is feet and inches
                FIFtoMeter = (Val(firstParam) * 304.8 + Val(secondParam) * 25.4) / 1000#
            Else
         ' if there is one space and a slash, it is inches and fractions
                ' numerator = Left(secondParam, InStr(secondParam, "/"))
                numerator = Mid(secondParam, 1, InStr(secondParam, "/")) ' DI#109789 ; changed to Mid as Left was giving compilation error because of enumtype 'left'
                denominator = Mid(secondParam, InStr(secondParam, "/") + 1)
                FIFtoMeter = (Val(firstParam) + Val(numerator) / Val(denominator)) * 25.4 / 1000#
            End If
        Else
     ' if there are two spaces that means there are feet, inches, and fractions of an inch to convert
            thirdParam = Mid(secondParam, InStr(secondParam, " ") + 1)
            ' secondParam = Left(secondParam, InStr(secondParam, " "))
            secondParam = Mid(secondParam, 1, InStr(secondParam, " ")) ' DI#109789 ; changed to Mid as Left was giving compilation error because of enumtype 'left'

            ' numerator = Left(thirdParam, InStr(thirdParam, "/"))
            numerator = Mid(thirdParam, 1, InStr(thirdParam, "/")) ' DI#109789 ; changed to Mid as Left was giving compilation error because of enumtype 'left'
            denominator = Mid(thirdParam, InStr(thirdParam, "/") + 1)
            FIFtoMeter = (Val(firstParam) * 304.8 + (Val(secondParam) + Val(numerator) / Val(denominator)) * 25.4) / 1000#
        End If
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: ReadParametricData
' Description: Create the sql statement that will be used in data look-ups
'
' Inputs - tableName - String, whereClause- String
' Outputs - String
' ---------------------------------------------------------------------------
Public Function ReadParametricData(tableName As String, whereClause As String) As String
On Error GoTo ErrorHandler
Const METHOD = "ReadParametricData"

    ReadParametricData = " FROM " & tableName & " " & whereClause

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: GetNParam
' Description: Retrieves a number from the database
'
' Inputs - record - String, column- String
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function GetNParam(record As String, column As String) As Double
On Error GoTo ErrorHandler
Const METHOD = "GetNParam"

    Dim oIJPartSelHlpr As IJHgrPartSelectionDBHlpr
    Set oIJPartSelHlpr = New PartSelectionHlpr
    
    Dim strQuery As String
    Dim i As Integer
    
    If PrevRecord <> record Then
        Dim partObjColl As IJDCollection
        'clear existing AttributeList
        Erase oAttributeList
        strQuery = "select * " & record
        Set partObjColl = oIJPartSelHlpr.GetPartCollectionFromDBQuery(strQuery)
        PrevRecord = record
        'copy partobjcoll into AttributeList
        SetAttributeList partObjColl.Item(1)
        Set partObjColl = Nothing
    End If
    Set oIJPartSelHlpr = Nothing
    
    'search the attribute (or column) value in attributelist
    For i = 0 To UBound(oAttributeList)
        If UCase(oAttributeList(i).AttributeName) = UCase(column) Then ' DI#109787
            GetNParam = oAttributeList(i).AttributeValue
            Exit Function
        End If
    Next i
    'couldn't find the column in attributelist simply return 0
    GetNParam = 0
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "Query: " + record + ", Column: " + column).Number
End Function

' ---------------------------------------------------------------------------
' Name: GetSParam
' Description: Retrieves a string from the database
'
' Inputs - record - String, column- String
' Outputs - string
' ---------------------------------------------------------------------------
Public Function GetSParam(record As String, column As String) As String
On Error GoTo ErrorHandler
Const METHOD = "GetSParam"

    Dim oIJPartSelHlpr As IJHgrPartSelectionDBHlpr
    Set oIJPartSelHlpr = New PartSelectionHlpr
    
    Dim strQuery As String
    Dim i As Integer
    
    If PrevRecord <> record Then
        Dim partObjColl As IJDCollection
        'clear existing AttributeList
        Erase oAttributeList
        strQuery = "select * " & record
        Set partObjColl = oIJPartSelHlpr.GetPartCollectionFromDBQuery(strQuery)
        
        PrevRecord = record
        'copy partobjcoll into AttributeList
        SetAttributeList partObjColl.Item(1)
        Set partObjColl = Nothing
    End If
    Set oIJPartSelHlpr = Nothing
    
    'search the attribute (or column) value in attributelist
    For i = 0 To UBound(oAttributeList)
        If UCase(oAttributeList(i).AttributeName) = UCase(column) Then ' DI#109787
            GetSParam = oAttributeList(i).AttributeValue
            Exit Function
        End If
    Next i
    'couldn't find the column in attributelist simply return null string
    GetSParam = ""
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "Query: " + record + ", Column: " + column).Number
End Function

' ---------------------------------------------------------------------------
' Name: CosDeg
' Description: Calculate the cosine of a angle that is in degrees
'
' Inputs - ANGLE - Double
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function CosDeg(angle As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "CosDeg"

    CosDeg = Cos(angle * 1.74532925199433E-02)

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: SinDeg
' Description: Calculate the sine of a angle that is in degrees
'
' Inputs - ANGLE - Double
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function SinDeg(angle As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "SinDeg"

    SinDeg = Sin(angle * 1.74532925199433E-02)

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: TanDeg
' Description: Calculate the tan of a angle that is in degrees
'
' Inputs - ANGLE - Double
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function TanDeg(angle As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "TanDeg"

    TanDeg = Tan(angle * 1.74532925199433E-02)

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: ACosDeg
' Description: Calculate the arc cosine of a angle that is in degrees
'
' Inputs - RATIO - Double
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function ACosDeg(RATIO As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "ACosDeg"

    ACosDeg = ACos(RATIO) * 57.2957795130823

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: ASinDeg
' Description: Calculate the arc sine of a angle that is in degrees
'
' Inputs - RATIO - Double
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function ASinDeg(RATIO As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "ASinDeg"

    ASinDeg = ASin(RATIO) * 57.2957795130823

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: ATanDeg
' Description: Calculate the arc tan of a angle that is in degrees
'
' Inputs - RATIO - Double
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function ATanDeg(RATIO As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "ATanDeg"

    ATanDeg = Atn(RATIO) * 57.2957795130823

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: Deg
' Description: Convert the value from radians to degrees
'
' Inputs - RadNumber - Double
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function Deg(RadNumber As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "Deg"

    Deg = RadNumber * 180 / 3.14159265358979

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: Rad
' Description: Convert the value from degrees to radians
'
' Inputs - RadNumber - Double
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function Rad(DegNumber As Double) As Double
On Error GoTo ErrorHandler
Const METHOD = "Rad"

    Rad = DegNumber / 180 * 3.14159265358979

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: ConvertValueToLongStringValue
' Description: This function can be used to look up data in a codelist or to convert units
'              optionally pass in bFIF, if true and the units are feet or inch, the string returned is in FIF format, 1/16 precision
'              optionally pass in iPrecision to control fractional precision for FIF, defaults to 16ths if not passed in
'              optionally pass in uPrimaryUnits to control the displayed units to override the custom interfaces sheet.  These
'              units must be one of the units as defined by IJUnitsOfMeasure in x:\doc\devHelp.chm
' bar Sept 16, 2005
'
' Inputs - oObject - object, strInterfaceName - string, strPropName - string, cvalue - variant, [bFIF] - boolean,
'          [iPrecision] - integer, [uPrimaryUnits] - units
' Outputs - String
' ---------------------------------------------------------------------------

Public Function ConvertValueToLongStringValue(ByVal oObject As Object, _
    strInterfaceName As String, strPropName As String, cvalue As Variant, _
    Optional bFIF As Boolean = False, _
    Optional iPrecision As Integer = 16, _
    Optional uPrimaryUnits As Units = TIME_FORTNIGHT) As String

    Const METHOD = "ConvertValueToLongStringValue"
    Dim sValue As String
    Dim longV As String
    Dim shortV As String
    Dim vInterfaceID As Variant
    Dim oAttribMetaData As IJDAttributeMetaData
    Dim oAttribInfo As IJDAttributeInfo
    Dim oCLMD As METADATALib.IJDCodeListMetaData
    Dim strTable As String
    Dim idUnit As Units

    On Error GoTo ErrHandler
    Set oAttribMetaData = oObject
    vInterfaceID = oAttribMetaData.iID(strInterfaceName)
    Set oAttribInfo = oAttribMetaData.AttributeInfo(vInterfaceID, strPropName)
    strTable = oAttribInfo.CodeListTableName
    If Not strTable = vbNullString Then
        Set oCLMD = oObject
        longV = oCLMD.LongStringValue(strTable, cvalue)
        If Not longV = vbNullString Then
            ConvertValueToLongStringValue = longV
        Else
            shortV = oCLMD.ShortStringValue(strTable, cvalue)
            If Not shortV = vbNullString Then
                ConvertValueToLongStringValue = shortV
            Else
                ConvertValueToLongStringValue = vbNullString
            End If
        End If
    Else
        Dim dblValue As Double      'TR40900 add unit
        Dim StrValue As String
        Dim oUomServices As UnitsOfMeasureServicesLib.UomVBInterface

        Set oUomServices = New UnitsOfMeasureServicesLib.UomVBInterface
        If (VarType(cvalue) = vbNull Or VarType(cvalue) = vbEmpty) Then
            cvalue = ""
        End If

        Dim NewUnits As Units
 ' If a special Primary Units was passed in then use it rather than what is in the custom interfaces sheet.
        If uPrimaryUnits <> TIME_FORTNIGHT Then  'so if it is not the function default
            NewUnits = uPrimaryUnits
        Else
            NewUnits = oAttribInfo.PrimaryUnits
        End If

        If VarType(cvalue) = vbDouble And oAttribInfo.UnitsType > 0 Then
         ' if units are feet or inch and fif is true, format to fif
            If (bFIF = True And oAttribInfo.UnitsType = UNIT_DISTANCE) And _
                (NewUnits = DISTANCE_INCH Or NewUnits = DISTANCE_FOOT) Then
                Dim xomFormat As IJUomVBFormat
                Set xomFormat = New UomVBFormat
                xomFormat.PrecisionType = PRECISIONTYPE_FRACTIONAL
                xomFormat.FractionalPrecision = iPrecision
                xomFormat.UnitsDisplayed = True
                xomFormat.ReduceFraction = True

                If cvalue > 0.305 Then 'if more than a foot, show in feet and inch
                    oUomServices.FormatUnit oAttribInfo.UnitsType, cvalue, StrValue, xomFormat, DISTANCE_FOOT, DISTANCE_INCH
                Else
                    oUomServices.FormatUnit oAttribInfo.UnitsType, cvalue, StrValue, xomFormat, NewUnits
                End If
            Else
                oUomServices.FormatUnit oAttribInfo.UnitsType, cvalue, StrValue, , NewUnits
            End If
            ConvertValueToLongStringValue = StrValue
        Else
            ConvertValueToLongStringValue = CStr(cvalue)
        End If

    End If
    Set oCLMD = Nothing
    Set oAttribInfo = Nothing
    Set oAttribMetaData = Nothing
    Set oUomServices = Nothing

    Exit Function

ErrHandler:
    'Call ErrHandler(Err, MODULE & METHOD)
    ConvertValueToLongStringValue = "<BAD VALUE>"
    Set oCLMD = Nothing
    Set oAttribInfo = Nothing
    Set oAttribMetaData = Nothing
    Set oUomServices = Nothing
End Function

' ---------------------------------------------------------------------------
' Name: ListAttributesOfObject
' Description: This function returns a list of all the attributes that the object has associated with it
'
' Inputs - pObject - object
' Outputs - String
' ---------------------------------------------------------------------------
Public Function ListAttributesOfObject(pObject As Object) As String
Const METHOD = "ListAttributesOfObject"
On Error GoTo ErrHandler

    Dim oAttributes             As IMSAttributes.IJDAttributes
    Dim oAttributesCol          As IJDAttributesCol
    Dim oAttribute              As IJDAttribute
    Dim oAttributeInfo          As IJDAttributeInfo
    Dim iID                     As Variant
    Dim iValue                  As Integer
    iValue = 0

    Dim strList As String

    Set oAttributes = pObject
    For Each iID In oAttributes
        Set oAttributesCol = oAttributes.CollectionOfAttributes(iID)
        If oAttributesCol.Count > 0 Then
            For Each oAttribute In oAttributesCol
                Set oAttributeInfo = oAttribute.AttributeInfo
                strList = strList + oAttributeInfo.name + " "
                Set oAttributeInfo = Nothing
                Set oAttribute = Nothing
            Next
        End If
        Set oAttributesCol = Nothing
    Next iID

    Set oAttributes = Nothing
    ListAttributesOfObject = strList
Exit Function
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Function

' ---------------------------------------------------------------------------
' Name: LogError
' Description: default Error logger
'
' Inputs - oErrObject - ErrObject, [strSourceFile] - string, [strMethod] - string, [strExtraInfo] - string
' Outputs - IJError
' ---------------------------------------------------------------------------
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError

Const METHOD = "LogError"
On Error GoTo ErrHandler

    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors

    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description

     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")

    ' add the error to the service : the error is also logged to the file
    ' specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/
    '      ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
Exit Function
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Function

' ---------------------------------------------------------------------------
' Name: LogErrorWithContext
' Description: default Error logger, same as LogError but adds the errors with proper error context
'
' Inputs - oErrObject - ErrObject, [strSourceFile] - string, [strMethod] - string, [strExtraInfo] - string
' Outputs - IJError
' ---------------------------------------------------------------------------
Public Function LogErrorWithContext(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "", _
                            Optional IsWarning As Boolean) As IJError ' TR#119636

Const METHOD = "LogErrorWithContext"
On Error GoTo ErrHandler

    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
    Dim strErrContext As String ' TR#119636

    ' TR#119636 ; set the error context based on whether it is a warning or error message
    If IsWarning = True Then
        strErrContext = UC_UserWarningMessage
    Else
        strErrContext = UC_UserErrorMessage
    End If
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description

     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")

    ' add the error to the service : the error is also logged to the file
    ' specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/
    '      ReportErrors_Log"
    Set LogErrorWithContext = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      strErrContext, _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
Exit Function
ErrHandler:
    Err.Raise Err.Number
End Function

' ---------------------------------------------------------------------------
' Name: PF_EventHandler
' Description: This procedure will allow us to use the same line of code to either log a message or raise an error when an expected event occurs
    ' WarnOnly = False --> We raise an error with the provided information and get kicked out of the current method
    ' WarnOnly = True --> We write the provided info to the log file but DO NOT raise an error, the method continues
'JRM 03.22.06
'
' Inputs - EventDescription - string, ErrorObject - ErrObject, SourceModule - String, SourceMethod - string
'          [WarnOnly] - boolean
' Outputs - String
' ---------------------------------------------------------------------------
Public Sub PF_EventHandler(EventDescription As String, ErrorObject As ErrObject, SourceModule As String, SourceMethod As String, Optional WarnOnly As Boolean = False)
'Const METHOD = "PF_EventHandler"
'On Error GoTo ErrHandler

    If WarnOnly = False Then
        Err.Raise LogErrorWithContext(ErrorObject, SourceModule, SourceMethod, "ERROR: " & EventDescription, False).Number
    Else
        LogErrorWithContext ErrorObject, SourceModule, SourceMethod, "WARNING: " & EventDescription, True
    End If

'Exit Sub
'ErrHandler:
'    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

'**************************************************************************************
' Function: RoundDimension
'
' This function will round the given dimension of the part based on the rule
' HgrSupDimensionRoundFactor.  This rule is from the HgrRules sheet of HS_System.xls
' This rule should return a long number corresponding to the number of decimal places
' to round the dimension to.
'
' Parameters:
'   oPart:      the part who's dimension is to be rounded
'   strName:    The name of the dimension to round. This should be a length dimension
'
' Author: BAPR, April 11, 2006
'**************************************************************************************
Public Sub RoundDimension(ByVal oPart As IJHgrSupportComponent, strName As String)
    Const METHOD = "RoundDimension"
    On Error GoTo ErrorHandler

    Dim oICH As HNGSUPSupportServices.IJHgrInputConfigHlpr ' need the ICH to get the rule value
    Dim vData() As Variant
    oPart.GetOccAssembly oICH
    If Not oICH Is Nothing Then ' The ICH will be null the first time thru until the symbol is computed the first time
        On Error GoTo Round 'in case the rule is not defined
        vData = oICH.GetDataByRule("HgrSupDimensionRoundFactor")  ' get the rounding factor from the rule
        On Error GoTo ErrorHandler
        Dim dLen As Double
        dLen = GetAttributeFromObject(oPart, strName) ' get the length before rounding
        oICH.SetAttributeValue strName, oPart, Round(dLen, vData(1)) ' set the new length to the rounded length
    End If
    Exit Sub 'normal exit
Round:     'just exit without trying to round
    Exit Sub
ErrorHandler:
        Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

' ---------------------------------------------------------------------------
' Name: AssignArraySizeForInput
' Description: Redimensions the Input array of a symbol
' Added this function for JTS DI 86724 SS
'
' Inputs - SymbolInput() - IMSSymbolEntities.IJDInput, I - Integer
' Outputs - NA
' ---------------------------------------------------------------------------
Public Sub AssignArraySizeForInput(ByRef SymbolInput() As IMSSymbolEntities.IJDInput, ByVal i As Integer)
Const METHOD = "AssignArraySizeForInput"
On Error GoTo ErrHandler

    If i > 20 Then
        Dim SymbolInput1() As IMSSymbolEntities.IJDInput
        ReDim SymbolInput1(1 To i)
        SymbolInput = SymbolInput1
    Else
        Dim SymbolInput2(20) As IMSSymbolEntities.IJDInput
        SymbolInput = SymbolInput2
    End If


Exit Sub
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

' ---------------------------------------------------------------------------
' Name: AssignArraySizeForOutput
' Description: Redimensions the Output array of a symbol
' Added this function for JTS DI 86724 SS
'
' Inputs - SimpleSymbolOutput() - IMSSymbolEntities.IJDOutput, I - Integer
' Outputs - NA
' ---------------------------------------------------------------------------
Public Sub AssignArraySizeForOutput(ByRef SimpleSymbolOutput() As IMSSymbolEntities.IJDOutput, ByVal i As Integer)
Const METHOD = "AssignArraySizeForOutput"
On Error GoTo ErrHandler

    If i > 20 Then
        Dim SymbolOutput1() As IMSSymbolEntities.IJDOutput
        ReDim SymbolOutput1(1 To i)
        SimpleSymbolOutput = SymbolOutput1
    Else
        Dim SymbolOutput2(20) As IMSSymbolEntities.IJDOutput
        SimpleSymbolOutput = SymbolOutput2
    End If

Exit Sub
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

'*************************************************************************************
' Method:       GetNParamUsingParametricQuery
'
' Description:

' Parameters:
'            record       -   Input (String)
'            column       -   Input(String)
'            pParameter   -   Input( Parameters)
'            Double       -(Ouput)
'*************************************************************************************
'Modified function based on PF_HgrGeneralUtil - GetNParamUsingParametricQuery(...)
'Lookup JHgrSupportJoint to get the information from specified column about SupportJoints xls sheet.
'return false, if no value is found
'July 7,2007 JL
'Fix for "Optional arrayN As Variant" is an array if multiple rows match the SQL query.

Public Function GetNParamUsingParametricQuery(record As String, column As String, pParameter() As Parameters, Optional arrayN As Variant) As Double
Const METHOD = "GetNParamUsingParametricQuery"
On Error GoTo ErrHandler
    Dim oIJPartSelHlpr As IJHgrPartSelectionDBHlpr
    Set oIJPartSelHlpr = New PartSelectionHlpr
    Dim query As String
    Dim J As Integer
    Dim bEnter As Boolean
    Dim ss(0) As Double
    arrayN = ss 'default value
    GetNParamUsingParametricQuery = 0 'set default value
    
    If Not bFirst Then
        mParameters = pParameter
        bEnter = False
        bFirst = True
    Else
        If UBound(mParameters) <> UBound(pParameter) Then
            bEnter = True
        Else
            Dim pCount As Integer
            Dim AllEqual As Integer
            AllEqual = 0
            For pCount = 0 To UBound(pParameter)
                If (mParameters(pCount).Value = pParameter(pCount).Value) Then
                    AllEqual = AllEqual + 1
                End If
            Next pCount
            If AllEqual = UBound(pParameter) + 1 Then
                bEnter = False
            Else
                bEnter = True
            End If
        End If
        Erase mParameters
        mParameters = pParameter
    End If
    If (PrevRecord <> record) Or bEnter Then
        Dim partObjColl As IJDCollection
        'clear existing AttributeList
        Erase oAttributeList
        'Get all the attributes of record
        query = "select * " & record
        For J = 0 To UBound(pParameter)
            Call oIJPartSelHlpr.AddParameter(pParameter(J).name, pParameter(J).Type, pParameter(J).Direction, pParameter(J).Length, pParameter(J).Value)
        Next J

        Set partObjColl = oIJPartSelHlpr.GetPartCollectionFromDBQuery(query)
        'Store all attributes information of given record into AttributeList object
        If partObjColl.Size > 0 Then
            ReDim dValue(partObjColl.Size) As Double
            Dim ii As Integer
            For ii = 1 To partObjColl.Size
                Erase oAttributeList
                SetAttributeList partObjColl.Item(ii)
                'search the attribute (or column) value in attributelist
                For J = 0 To UBound(oAttributeList)
                    If UCase(oAttributeList(J).AttributeName) = UCase(column) Then ' DI#109787
                        dValue(ii) = oAttributeList(J).AttributeValue
                        'GetNParamUsingParametricQueryNew22 = oAttributeList(j).AttributeValue
                    End If
                Next J
            Next
            If UBound(dValue) > 0 Then 'this is the only place where one or more rows were found from DB
                arrayN = dValue
                GetNParamUsingParametricQuery = dValue(1)
            Else
                PF_EventHandler "No data found", Err, MODULE, "GetNParamUsingParametricQuery", False
            End If
        Else
            PF_EventHandler "No data found", Err, MODULE, "GetNParamUsingParametricQuery", False
        End If

        'set the current record as previous record
        PrevRecord = record
        Set partObjColl = Nothing
    End If
    'search the attribute (or column) value in attributelist

    For J = 0 To UBound(oAttributeList)

        If UCase(oAttributeList(J).AttributeName) = UCase(column) Then ' DI#109787

            GetNParamUsingParametricQuery = oAttributeList(J).AttributeValue

            Exit Function

        End If

    Next J


    Set oIJPartSelHlpr = Nothing
    Exit Function
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "Query: " + record + ", Column: " + column).Number
End Function

'*************************************************************************************
' Method:       GetSParamUsingParametricQuery
'
' Description:

' Parameters:
'            record       -   Input (String)
'            column       -   Input(String)
'            pParameter   -   Input( Parameters)
'            String       -(Ouput)
'*************************************************************************************
Public Function GetSParamUsingParametricQuery(record As String, column As String, pParameter() As Parameters) As String
Const METHOD = "GetSParamUsingParametricQuery"
On Error GoTo ErrHandler

    Dim oIJPartSelHlpr As IJHgrPartSelectionDBHlpr
    Set oIJPartSelHlpr = New PartSelectionHlpr
    Dim query As String
    Dim J As Integer
    Dim bEnter As Boolean
    
    If Not bFirst Then
        mParameters = pParameter
        bEnter = False
        bFirst = True
    Else
        If UBound(mParameters) <> UBound(pParameter) Then
            bEnter = True
        Else
            Dim pCount As Integer
            Dim AllEqual As Integer
            AllEqual = 0
            For pCount = 0 To UBound(pParameter)
                If (mParameters(pCount).Value = pParameter(pCount).Value) Then
                    AllEqual = AllEqual + 1
                End If
            Next pCount
            If AllEqual = UBound(pParameter) + 1 Then
                bEnter = False
            Else
                bEnter = True
            End If
        End If
        Erase mParameters
        mParameters = pParameter

    End If
    If (PrevRecord <> record) Or bEnter Then
        
        Dim partObjColl As IJDCollection
        'clear existing AttributeList
        Erase oAttributeList
        'Get all the attributes of record
        query = "select * " & record
        For J = 0 To UBound(pParameter)
            Call oIJPartSelHlpr.AddParameter(pParameter(J).name, pParameter(J).Type, pParameter(J).Direction, pParameter(J).Length, pParameter(J).Value)
        Next J
    
        Set partObjColl = oIJPartSelHlpr.GetPartCollectionFromDBQuery(query)
    
        'Store all attributes information of given record into AttributeList object
        SetAttributeList partObjColl.Item(1)
        
        'set the current record as previous record
        PrevRecord = record
        Set partObjColl = Nothing
        
    End If
    Set oIJPartSelHlpr = Nothing
    
    'search the attribute (or column) value in attributelist
    For J = 0 To UBound(oAttributeList)
        If UCase(oAttributeList(J).AttributeName) = UCase(column) Then ' DI#109787
            GetSParamUsingParametricQuery = oAttributeList(J).AttributeValue
            Exit Function
        End If
    Next J
    'couldn't find the column in attributelist simply return null string
    GetSParamUsingParametricQuery = ""
    
    Exit Function
ErrHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "Query: " + record + ", Column: " + column).Number
End Function


''As parameterDirection will be adParamInput, this will be used by all FillParametersXXX()
Public Function ParamInput() As ParameterDirectionEnum
    ParamInput = adParamInput
End Function

'*************************************************************************************
' Method:       FillParameter
'
' Description:  Use this function to fill all the values for the Parameter

' Parameters:
'            Name    -   Input (String)
'            DataType-   Input(DataTypeEnum)
'            Value   -   Input( Variant)
'            Parameters-(Ouput)
'*************************************************************************************
Public Function FillParameter(name As String, DataType As DataTypeEnum, Value As Variant) As Parameters
    Const METHOD = "ParamInputs"
    On Error GoTo ErrorHandler
    Dim pParam As Parameters
        pParam.name = name
        pParam.Type = DataType
        pParam.Direction = ParamInput
        pParam.Value = Value
        pParam.Length = Len(pParam.Value)
        FillParameter = pParam
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:       FillParameterForDouble
'
' Description:  ''Use this function to fill Name and Value of a Parameter which is of type Double

' Parameters:
'            Name  -   Input (String)
'            Value  -   Input( Variant)
'            Parameters-(Ouput)
'*************************************************************************************
Public Function FillParameterForDouble(name As String, Value As Variant) As Parameters
    Const METHOD = "ParamInputs"
    On Error GoTo ErrorHandler
    Dim pParam As Parameters
        pParam.name = name
        pParam.Type = adDouble
        pParam.Direction = ParamInput
        pParam.Value = Value
        pParam.Length = Len(pParam.Value)
        FillParameterForDouble = pParam
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:       FillParameterForVariant
'
' Description:  ''Use this function to fill Name and Value of a Parameter which is of type Variant
' Parameters:
'            Name  -   Input (String)
'            Value  -   Input( Variant)
'            Parameters-(Ouput)
'*************************************************************************************
Public Function FillParameterForVariant(name As String, Value As Variant) As Parameters
    Const METHOD = "ParamInputs"
    On Error GoTo ErrorHandler
    Dim pParam As Parameters
        pParam.name = name
        pParam.Type = adVariant
        pParam.Direction = ParamInput
        pParam.Value = Value
        pParam.Length = Len(pParam.Value)
        FillParameterForVariant = pParam
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:       FillParameterForBSTR
'
' Description:  ''Use this function to fill Name and Value of a Parameter which is of type BSTR
' Parameters:
'            Name  -   Input (String)
'            Value  -   Input( Variant)
'            Parameters-(Ouput)
'*************************************************************************************
Public Function FillParameterForBSTR(name As String, Value As Variant) As Parameters
    Const METHOD = "ParamInputs"
    On Error GoTo ErrorHandler
    Dim pParam As Parameters
        pParam.name = name
        pParam.Type = adBSTR
        pParam.Direction = ParamInput
        pParam.Value = Value
        pParam.Length = Len(pParam.Value)
        FillParameterForBSTR = pParam
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:       FillParameterForVarChar
'
' Description:  ''Use this function to fill Name and Value of a Parameter which is of type VarChar
' Parameters:
'            Name  -   Input (String)
'            Value  -   Input( Variant)
'            Parameters-(Ouput)
'*************************************************************************************
Public Function FillParameterForVarChar(name As String, Value As Variant) As Parameters
    Const METHOD = "ParamInputs"
    On Error GoTo ErrorHandler
    Dim pParam As Parameters
        pParam.name = name
        pParam.Type = adVarChar
        pParam.Direction = ParamInput
        pParam.Value = Value
        pParam.Length = Len(pParam.Value)
        FillParameterForVarChar = pParam
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:       FillParameterForInteger
'
' Description:  ''Use this function to fill Name and Value of a Parameter which is of type Integer
'
' Parameters:
'            Name  -   Input (String)
'            Value  -   Input( Variant)
'            Parameters-(Ouput)
'*************************************************************************************
Public Function FillParameterForInteger(name As String, Value As Variant) As Parameters
    Const METHOD = "ParamInputs"
    On Error GoTo ErrorHandler
    Dim pParam As Parameters
        pParam.name = name
        pParam.Type = adInteger
        pParam.Direction = ParamInput
        pParam.Value = Value
        pParam.Length = Len(pParam.Value)
        FillParameterForInteger = pParam
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   FillParameterForBoolean
'
' Description:  ''Use this function to fill Name and Value of a Parameter which is of type Boolean
'
' Parameters:
'            Name  -   Input (String)
'            Value  -   Input( Variant)
'            Parameters-(Ouput)
'*************************************************************************************
Public Function FillParameterForBoolean(name As String, Value As Variant) As Parameters
    Const METHOD = "ParamInputs"
    On Error GoTo ErrorHandler
    Dim pParam As Parameters
        pParam.name = name
        pParam.Type = adBoolean
        pParam.Direction = ParamInput
        pParam.Value = Value
        pParam.Length = Len(pParam.Value)
        FillParameterForBoolean = pParam
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
Public Sub SetAttributeList(ByVal pObject As Object)
Const METHOD = "SetAttributeList"
On Error GoTo ErrHandler

    Dim oAttributes             As IMSAttributes.IJDAttributes
    Dim oAttributesCol          As IJDAttributesCol
    Dim oAttribute              As IJDAttribute
    Dim oAttributeInfo          As IJDAttributeInfo
    Dim iID                     As Variant
    Dim i                       As Integer
    i = 0

    Set oAttributes = pObject
    For Each iID In oAttributes
        Set oAttributesCol = oAttributes.CollectionOfAttributes(iID)
        If oAttributesCol.Count > 0 Then
            'it preserve the attributues of all IID
            ReDim Preserve oAttributeList(i + oAttributesCol.Count - 1)
            For Each oAttribute In oAttributesCol
                Set oAttributeInfo = oAttribute.AttributeInfo
                oAttributeList(i).AttributeName = oAttributeInfo.name
                oAttributeList(i).AttributeValue = oAttribute.Value
                i = i + 1
                Set oAttributeInfo = Nothing
                Set oAttribute = Nothing
            Next
        End If
        Set oAttributesCol = Nothing
    Next iID

        Exit Sub
ErrHandler:
     Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

' ---------------------------------------------------------------------------
' Aug 15 2006 - Jim McCrea -
' MultipleInterfaceDataLookup
' - this function does a complex query between 2 different interfaces that are on the SAME part
' - THe interfaces must have the oid property in common
' - The query this function was derived from (thanks Brant):
'       select * from juahgrmaxload650 where oid in (select oid from JCUHASAnvil_fig69 where oid in
'                           (select oid from juaHgrPipe_dia where pipe_dia >.02 and pipe_dia < .022))
' - Limitations: Right now the function just fails with a subscript out of rane error when query fails.  THis should be cleaned up
' - Sample function call:
'       xxx_NUCLEAR_COMMON_MultipleInterfaceDataLookup("MAX_LOAD_650", "juahgrmaxload650", "JCUHASAnvil_fig69", "pipe_dia", "juaHgrPipe_dia", "0.02", "0.022")
' or
'       xxx_NUCLEAR_COMMON_MultipleInterfaceDataLookup("MAX_LOAD_650", "juahgrmaxload650", "JCUHASAnvil_fig69", "pipe_dia", "juaHgrPipe_dia", "0.1143")
'
' sDestColName - the name of the column we are looking up
' sDestColInterface - the interface this column exists on (with the i removed of course ;))
' sPartInterface - the interface of the part we are looking up
' sRefColName - the name of the column we are referencing
' sRefInterfance - the name of the interface that the reference column exists on
' sRefValue - the value we are using for comparison purposes.  If the option sMaxRefValue is included then this is considered the Min value
' [sMaxRefValue] - the max value we are using for comparison purposes.
'
' Returns a Variant with the result of the query
'
' ---------------------------------------------------------------------------

Public Function MultipleInterfaceDataLookUp(sDestColName As String, sDestColInterface As String, sPartInterface As String, _
                                                                sRefColName As String, sRefInterfance As String, sRefValue As String, _
                                                                Optional sMaxRefValue As String = "None") As Variant
 
Const METHOD = "MultipleInterfaceDataLookup"
On Error GoTo ErrorHandler
 
    'Build the complex query
    Dim sComplexQuery As String
    
    If sMaxRefValue = "None" Then
        sComplexQuery = "Select [" & UCase(sDestColName) & "] from " & sDestColInterface & " where oid in " _
            & "(Select oid from " & sPartInterface & " where oid in " _
            & "(Select oid from " & sRefInterfance & " where [" & UCase(sRefColName) & "]= " & sRefValue & "))"
    Else
        sComplexQuery = "Select [" & UCase(sDestColName) & "] from " & sDestColInterface & " where oid in " _
            & "(Select oid from " & sPartInterface & " where oid in " _
            & "(Select oid from " & sRefInterfance & " where [" & UCase(sRefColName) & "]> " & sRefValue & " and [" & UCase(sRefColName) & "]< " & sMaxRefValue & "))"
    End If
 
    Dim oIJPartSelHlpr As IJHgrPartSelectionDBHlpr
    Set oIJPartSelHlpr = New PartSelectionHlpr
    
    Dim vAnswer() As Variant
    vAnswer = oIJPartSelHlpr.GetValuesFromDBQuery(sComplexQuery)
    
    MultipleInterfaceDataLookUp = vAnswer(0)
 
Exit Function
 
ErrorHandler:
        Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: GetSupportingTypes()
' Description: Return the type of structure being connected to
' May 29 2006 - JRK
'
' Inputs - None
' Outputs - Double
' ---------------------------------------------------------------------------
Public Function GetSupportingTypes(my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr) As String
 
Const METHOD = "GetSupportingTypes"
' Steel Section look up
    On Error GoTo ErrorHandler
    
    Dim eHgrCmdType As HgrCmdType
    Dim oIJHgrSupport As IJHgrSupport
    Set oIJHgrSupport = my_IJHgrInputConfigHlpr
    eHgrCmdType = oIJHgrSupport.CommandType
    
    'Determine whether connecting to Steel or a Slab
    Dim eStructShape1 As HgrPortShape
    Dim eStructShape2 As HgrPortShape
    Dim oBeamCollection As Object
    Dim oStruct As Object
    Dim sConnection As String
    
    If eHgrCmdType = HgrByReferencePointCmdType Then
        GetSupportingTypes = "Slab"
    Else
        my_IJHgrInputConfigHlpr.GetStructure oBeamCollection
        If Not oBeamCollection Is Nothing Then
        If oBeamCollection.Count > 1 Then
            GetSupportingTypes = "Steel"
            Set oStruct = oBeamCollection.Item(1)
            eStructShape1 = my_IJHgrInputConfigHlpr.GetStructureShape(oStruct)
            Set oStruct = Nothing
            Set oStruct = oBeamCollection.Item(2)
            eStructShape2 = my_IJHgrInputConfigHlpr.GetStructureShape(oStruct)
            Set oStruct = Nothing
        
            If eStructShape1 = 3 And eStructShape2 = 3 Then 'Two Slabs
                GetSupportingTypes = "Slab"
            End If
        
            If eStructShape1 = 3 And eStructShape2 = 4 Then 'Slab then Steel
                GetSupportingTypes = "Slab-Steel"
            End If
        
            If eStructShape1 = 4 And eStructShape2 = 3 Then 'Steel then Slab
                GetSupportingTypes = "Steel-Slab"
            End If
        Else
            Set oStruct = oBeamCollection.Item(1)
            eStructShape1 = my_IJHgrInputConfigHlpr.GetStructureShape(oStruct)
            If eStructShape1 = 4 Then
                GetSupportingTypes = "Steel"
            Else
                GetSupportingTypes = "Slab"
            End If
            Set oStruct = Nothing
        End If
        End If
    End If
        
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Nov 01 2006 - JRM
' IsStructureSlopedAcrossPipe
' - this function takes in the config helper and a string (slab or steel)
'
' Returns a boolean
'
' ---------------------------------------------------------------------------
 
Public Function IsStructureSlopedAcrossPipe(myIJHgrInputConfigHlpr As IJHgrInputConfigHlpr, sStructType As String, Optional bOnTurn As Boolean = False) As Boolean

Const METHOD = "IsStructureSlopedAcrossPipe"

' Steel Section look up

    On Error GoTo ErrorHandler

    Dim dRefAngle As Double

    dRefAngle = myIJHgrInputConfigHlpr.GetPortOrientAngle(ORIENT_GLOBAL_Z, "Structure", HGRPORT_Y)

    Dim dRefAngle2 As Double

    dRefAngle2 = myIJHgrInputConfigHlpr.GetPortOrientAngle(ORIENT_GLOBAL_Z, "Structure", HGRPORT_X)

    If sStructType = "Slab" Then
        If dRefAngle > PF_HgrGeneralUtil.Rad(89.99999999999) And dRefAngle < PF_HgrGeneralUtil.Rad(90.0000000001) Then
            IsStructureSlopedAcrossPipe = False
        Else
            IsStructureSlopedAcrossPipe = True
        End If
    Else
        If dRefAngle2 > PF_HgrGeneralUtil.Rad(89.99999999999) And dRefAngle2 < PF_HgrGeneralUtil.Rad(90.0000000001) Then
            IsStructureSlopedAcrossPipe = False
        Else
            If myIJHgrInputConfigHlpr.IsPlaceByStructure Or bOnTurn Then

                IsStructureSlopedAcrossPipe = True
            Else
                IsStructureSlopedAcrossPipe = False
            End If
        End If
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Nov 06 2006 - SS
' Name: GetSupportingPropertyByRule()
' - this function takes in the property name, index and config helper
'
' Returns the value of the property if the property is found in the database for that section
'
' ---------------------------------------------------------------------------
Public Function GetSupportingPropertyByRule(ByVal strProperty As String, Index As Long, my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr) As Double
Const METHOD = "GetSupportingPropertyByRule"

    On Error GoTo ErrorHandler
    
        Dim SectionName As Variant
        Dim oBeamCollection As Object
        Dim oStruct As Object

        Dim strQuery As String
        Dim oIJPartSelHlpr As IJHgrPartSelectionDBHlpr
        Set oIJPartSelHlpr = New PartSelectionHlpr
        Dim partObjColl As IJDCollection

        Dim vCount As Variant
        Dim oAttributes             As IMSAttributes.IJDAttributes
        Dim oAttributesCol          As IJDAttributesCol
        Dim oAttribute              As IJDAttribute
        Dim oAttributeInfo          As IJDAttributeInfo
        Dim iID                     As Variant
        
        GetSupportingPropertyByRule = 0
        
        Dim eHgrCmdType As HgrCmdType
        Dim oIJHgrSupport As IJHgrSupport
        Set oIJHgrSupport = my_IJHgrInputConfigHlpr
       
        eHgrCmdType = oIJHgrSupport.CommandType

        If eHgrCmdType = HgrByReferencePointCmdType Then
            GetSupportingPropertyByRule = 0.02
        Else
            my_IJHgrInputConfigHlpr.GetStructure oBeamCollection
            Set oStruct = oBeamCollection.Item(Index)
    
            SectionName = my_IJHgrInputConfigHlpr.GetPropertyFromSupporting("SectionName", oStruct)
    
            strQuery = " Select * from JSTRUCTFlangedSectDimensions where oid in (Select oid from JSTRUCTCrossSection" & _
                       " Where SectionName = '" & CStr(SectionName) & "')"
    
            Set partObjColl = oIJPartSelHlpr.GetPartCollectionFromDBQuery(strQuery)
    
            If partObjColl.Size = 0 Then
                Exit Function
            ElseIf partObjColl.Size > 0 Then
                
                Set oAttributes = partObjColl.Item(1)
                For Each iID In oAttributes
                    Set oAttributesCol = oAttributes.CollectionOfAttributes(iID)
                    If oAttributesCol.Count > 0 Then
                        For Each oAttribute In oAttributesCol
                            Set oAttributeInfo = oAttribute.AttributeInfo
                            If oAttributeInfo.name = strProperty Then
                                GetSupportingPropertyByRule = oAttribute.Value
                                Exit Function
                            End If
                            Set oAttributeInfo = Nothing
                            Set oAttribute = Nothing
                        Next
                    End If
                    Set oAttributesCol = Nothing
                Next iID
            End If
        End If
        Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: GetCableTrayCSDimensions
' Description: This function returns width and height of the selected cable tray.
'
' Inputs - sblOcc - object
' Outputs - dWidth - Double, dHeight - Double
' ---------------------------------------------------------------------------
Public Sub GetCableTrayCSDimensions(ByVal sblOcc As Object, ByRef dWidth As Double, ByRef dHeight As Double)
Const METHOD = "GetCableTrayWidth"
    On Error GoTo ErrorHandler

    Dim oSupportComp                As IJHgrSupportComponent
    Dim my_IJHgrInputConfigHlpr     As IJHgrInputConfigHlpr
    Dim inputObjColl                As Object
    Dim inputObj                    As Object
    Dim dPipeRadius                 As Double
    Dim strUnitType                 As String
    Dim csType                      As CrossSectionShapeTypes

    Set oSupportComp = sblOcc
    oSupportComp.GetOccAssembly my_IJHgrInputConfigHlpr
    Set oSupportComp = Nothing
    
    If my_IJHgrInputConfigHlpr Is Nothing Then
        dWidth = 0.1254 '6 in for now
        dHeight = 0.1016 '4 in for now
    Else
        my_IJHgrInputConfigHlpr.GetCableTray inputObjColl
        Set inputObj = inputObjColl.Item(1)
        my_IJHgrInputConfigHlpr.GetRteCrossSectParams inputObj, csType, strUnitType, dPipeRadius, dWidth, dHeight
    End If
    
Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

'*************************************************************************************
' Method:   GetPipeODwithoutInsulation
' Desc:     'GetPipeODwithoutInsulation
'            Function to get PipeOD without insulation - Takes dTotalOD as double
'            oRoute as Object as Input and returns the PipeOD without insulation
'
' Param:
'           dTotalOD                       -    dTotalOD as  double - input
'           oRoute                         -    oRoute as Object - input
'           Ouput                          -    returns the PipeOD without insulation
'*************************************************************************************
Public Function GetPipeODwithoutInsulation(dTotalOD As Double, oRoute As Object) As Double
Const METHOD = "GetPipeODwithoutInsulation"
    On Error GoTo ErrorHandler
    
    Dim oInsulationObject As IJRteInsulation
    Set oInsulationObject = oRoute
    
    Dim dInsulat As Double
    dInsulat = oInsulationObject.Thickness
    
    GetPipeODwithoutInsulation = dTotalOD - 2 * dInsulat
    Set oInsulationObject = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetUserAttribute
' Desc:
'            Get the given User attribute from the given interface.
'
' Param:
'           oPart                       -    IJDPart (a row in the catalog)
'           strInterface                -    The interface name to look on
'           strItem                     -    the attribute name form the interface to get
'                                            data from
' Example:
' Private Sub IJHgrSymbolBOMServices_EvaluateBOM(ByVal pSupportComp As Object, bstrBOMDesc As String)
'
'    Dim oPartOcc As PartOcc
'    Dim oPart As IJDPart
'    Set oPartOcc = pSupportComp ' The part occurence
'    oPartOcc.GetPart oPart      ' The associated catalog part
'
'    Dim vMatGrade As Variant
'    vMatGrade = GetUserAttribute(oPart, "IJHgrPartMaterial", "MaterialGrade")
'*************************************************************************************
Public Function GetUserAttribute(oPart As IJDPart, strInterface As String, strItem As String) As Variant
Const METHOD = "GetUserAttribute"
    On Error GoTo ErrorHandler
    
    Dim oAttributeHelper As IJDAttributes
    Dim oAttributeCol As IJDAttributesCol
    Dim vAttributeValue As Variant
     
    Set oAttributeHelper = oPart
        
    On Error Resume Next
    Set oAttributeCol = oAttributeHelper.CollectionOfAttributes(strInterface)
    On Error GoTo ErrorHandler
    
    If (oAttributeCol Is Nothing) Then
        '  Interface not found
        GetUserAttribute = strInterface & "." & strItem & " Not Found"
    Else
        On Error GoTo NoItem
        vAttributeValue = oAttributeCol.Item(strItem).Value
        On Error GoTo ErrorHandler
        GoTo KeepOn
NoItem: 'item not found
        vAttributeValue = strInterface & "." & strItem & " Not Found"
KeepOn: 'item was found
        GetUserAttribute = vAttributeValue
    End If
    
    Set oAttributeHelper = Nothing
    Set oAttributeCol = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: AddInput
' Description: Used in IJDUserSymbolServices_InitializeSymbolDefinition to
'              add inputs.
'
' Parameters:
' Inputs  - strName; the name to add
'         - strType; the data type of the name being added
' Outputs - tInputs; contains array of names/dataTypes and the count of items
'
' Example:
' AddInput Inputs, "DR", "Double"
'
' Note:
' Once all the Symbol Inputs are added, need to add another statement as below.That is te last statement should be
' as below. This statement is to make sure that the index of the array is set back to 0 for the next symbol
' AddInput Inputs, "AddInputs_Done", ""
'
' Revision History:
' Date          Author  Revision
' ~~~~~~~~~~~~~ ~~~~~~  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' June 23, 2006 BAPR    Created
'
' ---------------------------------------------------------------------------
Public Sub AddInput(Inputs As tInputs, strName As String, sTrType As String)
Const METHOD = "AddInput"
    On Error GoTo ErrorHandler
    
    Dim iSize As Long
    Static iIdx As Long
    
    If UCase(strName) = "ADDINPUTS_DONE" Then
        iIdx = 0
        Exit Sub
    End If
    
    If iIdx = 0 Then
        iIdx = 2
    End If
    iSize = -1
    On Error Resume Next
    iSize = UBound(Inputs.sDataType)
    On Error GoTo ErrorHandler
    If iSize = -1 Or iSize < iIdx Then
        ReDim Preserve Inputs.sDataType(iSize + 50)
        ReDim Preserve Inputs.sInputName(iSize + 50)
    End If
    
    Inputs.sDataType(iIdx) = sTrType
    Inputs.sInputName(iIdx) = strName
    Inputs.iCount = iIdx
    
    iIdx = iIdx + 1
    
Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub
' ---------------------------------------------------------------------------
' Name: AddOutput
' Description: Used in IJDUserSymbolServices_InitializeSymbolDefinition to
'              add Outputs.
'
' Parameters:
' Inputs  - strName; the name to add
' Outputs - tOutputs; contains array of names and the count of items
'
' Example:
' AddOutput Outputs, "Port1"
' In the above statement, the name of the output (i.e. "Port1") is case sensitive and should be the same
' when passing it to any of the add geometry functions like AddCylinder
'
' Note:
' Once all the Outputs are added, need to add another statement as below. That is te last statement should be
' as below. This statement is to make sure that the index of the array is set back to 0 for the next symbol
' AddOutput Outputs, "AddOutputs_Done"
'
' Revision History:
' Date          Author  Revision
' ~~~~~~~~~~~~~ ~~~~~~  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Oct 31, 2006 IRK    Created
'
' ---------------------------------------------------------------------------
Public Sub AddOutput(Outputs As tOutputs, strName As String)
Const METHOD = "AddOutput"
    On Error GoTo ErrorHandler
    
    Dim iSize As Long
    Static iIdx As Long
    
    If UCase(strName) = "ADDOUTPUTS_DONE" Then
        iIdx = 0
        Exit Sub
    End If
    
    If iIdx = 0 Then
        iIdx = 1
    End If
    iSize = -1
    On Error Resume Next
    iSize = UBound(Outputs.sOutputName)
    On Error GoTo ErrorHandler
    If iSize = -1 Or iSize < iIdx Then
        ReDim Preserve Outputs.sOutputName(iSize + 50)
    End If
    
    Outputs.sOutputName(iIdx) = strName
    Outputs.iCount = iIdx
    
    iIdx = iIdx + 1
    
Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub
' ---------------------------------------------------------------------------
' Name: SetupInputsAndOutputsEx
' Description: This function was some extension kind of code which in turn calls the SetupInputsAndOutputsForCache that modularize the code
' needed to generate the inputs and outputs that are required to develop a SP3D symbol
'
' Example: any Symbol class file
'
' Inputs - pSymbolDefinition - IMSSymbolEntities.IJDSymbolDefinition, Inputs - Type,
'          Outputs - Type, m_progID - String
' Outputs - none
' ---------------------------------------------------------------------------
Public Sub SetupInputsAndOutputsEx(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, _
        ByRef Inputs As tInputs, ByRef Outputs As tOutputs, m_progID As String, Optional bCached As Boolean = True)
Const METHOD = "SetupInputsAndOutputsEx"
    On Error GoTo ErrorHandler
    
    Dim InputData As tInputs
    InputData = Inputs
    
    Dim OutputData As tOutputs
    OutputData = Outputs
    
    'SetupInputsAndOutputsForCache pSymbolDefinition, InputData.iCount, Outputs.iCount, InputData.sInputName, InputData.sDataType, Outputs.sOutputName, m_progID
    
    If bCached = True Then

        SetupInputsAndOutputsForCache pSymbolDefinition, InputData.iCount, Outputs.iCount, InputData.sInputName, InputData.sDataType, Outputs.sOutputName, m_progID

    Else

        SetupInputsAndOutputs pSymbolDefinition, InputData.iCount, Outputs.iCount, InputData.sInputName, InputData.sDataType, Outputs.sOutputName, m_progID

    End If

Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

' ---------------------------------------------------------------------------
' Name: GetSectionDataFromDatabase()
' Description: Will get the data for the specified steel section type, standard and name.
' Date - Author: May 17 2006 - JRM
'
' Inputs: secType As String - Type of Section for example "WT"
'         SecStandard As String - Steel standard to be used for Section
'         SecName as String - The name of the Standard for example "L100X100X10
' Outputs: An array of doubles containing the data for the specified section size.
' ---------------------------------------------------------------------------
Public Function GetSectionDataFromDatabase(secType As String, SecStandard As String, SecName As String) As Double()
Const METHOD = "GetSectionDataFromDatabase"

    On Error GoTo ErrorHandler

    Dim CatalogDef As Object
    Dim oCatResMgr As IUnknown
    Dim dAnswer() As Double


    Set oCatResMgr = GetCatalogResourceManager

    Dim m_xService As SP3DStructGenericTools.CrossSectionServices
    Set m_xService = New SP3DStructGenericTools.CrossSectionServices

  'Use the secion type, section standard & section name that you want.

    m_xService.GetStructureCrossSectionDefinition oCatResMgr, _
                                                  SecStandard, secType, _
                                                  SecName, CatalogDef
    Dim Var As Variant
    ReDim dAnswer(4) As Double
    On Error Resume Next
    m_xService.GetCrossSectionAttributeValue CatalogDef, "width", Var
    dAnswer(1) = Var
    m_xService.GetCrossSectionAttributeValue CatalogDef, "tf", Var
    dAnswer(2) = Var
    m_xService.GetCrossSectionAttributeValue CatalogDef, "tw", Var
    dAnswer(3) = Var
    m_xService.GetCrossSectionAttributeValue CatalogDef, "depth", Var
    dAnswer(4) = Var

    GetSectionDataFromDatabase = dAnswer

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: DoesExistSymbolAttribute
' Description: This function checks for the attributename in the interface and returns
' a boolean value

'
' Inputs - strInterfaceName - string, strAttributeName String
' Outputs - True or False
' ---------------------------------------------------------------------------
Public Function DoesExistSymbolAttribute(ByVal strInterfaceName As String, ByVal strAttributeName As String) As Boolean
Const METHOD = "DoesExistSymbolAttribute"
    On Error GoTo ErrorHandler

    Dim strQuery As String

    Dim oIJPartSelHlpr As IJHgrPartSelectionDBHlpr
    Set oIJPartSelHlpr = New PartSelectionHlpr
    Dim partObjColl As IJDCollection

    Dim oAttributes             As IMSAttributes.IJDAttributes
    Dim oAttributesCol          As IJDAttributesCol
    Dim oAttribute              As IJDAttribute
    Dim oAttributeInfo          As IJDAttributeInfo
    Dim iID                     As Variant

    DoesExistSymbolAttribute = False
    
    strQuery = " Select * from " & strInterfaceName

    On Error Resume Next
    Set partObjColl = oIJPartSelHlpr.GetPartCollectionFromDBQuery(strQuery)
    
        On Error Resume Next
    
        If partObjColl Is Nothing Then Exit Function
        
        If partObjColl.Size = 0 Then
            Exit Function
        ElseIf partObjColl.Size > 0 Then
            On Error GoTo ErrorHandler
            Set oAttributes = partObjColl.Item(1)
            For Each iID In oAttributes
                Set oAttributesCol = oAttributes.CollectionOfAttributes(iID)
                If oAttributesCol.Count > 0 Then
                    For Each oAttribute In oAttributesCol
                        Set oAttributeInfo = oAttribute.AttributeInfo
                        If UCase(oAttributeInfo.name) = UCase(strAttributeName) Then
                            DoesExistSymbolAttribute = True
                        End If
                    Next
                End If
            Next iID
        End If
Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

