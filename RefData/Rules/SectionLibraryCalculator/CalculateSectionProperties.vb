'=====================================================================================================
'
'Copyright 2011 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
'
'File
'  CalculateSectionProperties.vb
'
'=====================================================================================================
Imports System
Imports System.Exception
Imports System.Collections.Generic
Imports Ingr.SP3D.Common.Middle '<InstallDir>\Core\Container\Bin\Assemblies\Release\CommonMiddle.dll
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle '<InstallDir>\Core\Container\Bin\Assemblies\Release\ReferenceDataMiddle.dll

'==========================================================================================================================
'DefinitionName/ProgID is "SectionLibraryCalculator,Ingr.SP3D.ReferenceData.Client.Services.CalculateSectionProperties"
'==========================================================================================================================

Namespace Ingr.SP3D.ReferenceData.Client.Services

    Public Class CalculateSectionProperties
        Implements ISectionCalculator

#Region "Private members"

        Private Const IStructCrossSectionDimensions_Width$ = "IStructCrossSectionDimensions:Width"
        Private Const IStructCrossSectionDimensions_Depth$ = "IStructCrossSectionDimensions:Depth"
        Private Const IStructCrossSectionDimensions_Area$ = "IStructCrossSectionDimensions:Area"
        Private Const IStructCrossSectionDimensions_Perimeter$ = "IStructCrossSectionDimensions:Perimeter"
        Private Const IStructFlangedSectionDimensions_tw$ = "IStructFlangedSectionDimensions:tw"
        Private Const IStructFlangedSectionDimensions_tf$ = "IStructFlangedSectionDimensions:tf"
        Private Const IStructFlangedSectionDimensions_kdesign$ = "IStructFlangedSectionDimensions:kdesign"
        Private Const IStructFlangedSectionDimensions_d$ = "IStructFlangedSectionDimensions:d"
        Private Const IStructFlangedSectionDimensions_bf$ = "IStructFlangedSectionDimensions:bf"
        Private Const IStructCrossSectionDesignProperties_IsHollow$ = "IStructCrossSectionDesignProperties:IsHollow"
        Private Const IStructCrossSectionDesignProperties_Ixx$ = "IStructCrossSectionDesignProperties:Ixx"
        Private Const IStructCrossSectionDesignProperties_Zxx$ = "IStructCrossSectionDesignProperties:Zxx"
        Private Const IStructCrossSectionDesignProperties_Sxx$ = "IStructCrossSectionDesignProperties:Sxx"
        Private Const IStructCrossSectionDesignProperties_Rxx$ = "IStructCrossSectionDesignProperties:Rxx"
        Private Const IStructCrossSectionDesignProperties_Iyy$ = "IStructCrossSectionDesignProperties:Iyy"
        Private Const IStructCrossSectionDesignProperties_Zyy$ = "IStructCrossSectionDesignProperties:Zyy"
        Private Const IStructCrossSectionDesignProperties_Syy$ = "IStructCrossSectionDesignProperties:Syy"
        Private Const IStructCrossSectionDesignProperties_Ryy$ = "IStructCrossSectionDesignProperties:Ryy"
        Private Const IStructCrossSectionDesignProperties_Rxy$ = "IStructCrossSectionDesignProperties:Rxy"
        Private Const IStructCrossSectionDesignProperties_J$ = "IStructCrossSectionDesignProperties:J"
        Private Const IStructCrossSectionDesignProperties_Cw$ = "IStructCrossSectionDesignProperties:Cw"
        Private Const IStructCrossSectionDesignProperties_Sw$ = "IStructCrossSectionDesignProperties:Sw"
        Private Const IStructCrossSectionDesignProperties_ro$ = "IStructCrossSectionDesignProperties:ro"
        Private Const IStructCrossSectionDesignProperties_H$ = "IStructCrossSectionDesignProperties:H"
        Private Const IStructCrossSectionDesignProperties_CentroidX$ = "IStructCrossSectionDesignProperties:CentroidX"
        Private Const IStructCrossSectionDesignProperties_CentroidY$ = "IStructCrossSectionDesignProperties:CentroidY"
        Private Const IStructCrossSectionDesignProperties_IsSymmetricAboutX$ = "IStructCrossSectionDesignProperties:IsSymmetricAboutX"
        Private Const IStructCrossSectionDesignProperties_IsSymmetricAboutY$ = "IStructCrossSectionDesignProperties:IsSymmetricAboutY"
        Private Const IStructCrossSectionUnitWeight_UnitWeight$ = "IStructCrossSectionUnitWeight:UnitWeight"
        Private Const IStructFlangedBoltGage_gf$ = "IStructFlangedBoltGage:gf"
        Private Const IStructFlangedBoltGage_gw$ = "IStructFlangedBoltGage:gw"
        Private Const IJUA2L_bb$ = "IJUA2L:bb"
        Private Const IJUAHSS_tnom$ = "IJUAHSS:tnom"
        Private Const IJUAHSS_tdes$ = "IJUAHSS:tdes"
        Private Const IJUAHSSR_b_t$ = "IJUAHSSR:b_t"
        Private Const IJUAHSSR_h_t$ = "IJUAHSSR:h_t"
        Private Const IJUAHSSC_D_t$ = "IJUAHSSC:D_t"
        Private Const IUABuiltUpTopFlange_TopFlangeWidth$ = "IUABuiltUpTopFlange:TopFlangeWidth"
        Private Const IUABuiltUpTopFlange_TopFlangeThickness$ = "IUABuiltUpTopFlange:TopFlangeThickness"
        Private Const IUABuiltUpTopFlange_TopFlangeWidthExt$ = "IUABuiltUpTopFlange:TopFlangeWidthExt"
        Private Const IUABuiltUpBottomFlange_BottomFlangeWidthExt$ = "IUABuiltUpBottomFlange:BottomFlangeWidthExt"
        Private Const IUABuiltUpBottomFlange_BottomFlangeWidth$ = "IUABuiltUpBottomFlange:BottomFlangeWidth"
        Private Const IUABuiltUpBottomFlange_BottomFlangeThickness$ = "IUABuiltUpBottomFlange:BottomFlangeThickness"
        Private Const IUABuiltUpWeb_WebThickness$ = "IUABuiltUpWeb:WebThickness"
        Private Const IUABuiltUpWeb_DepthExt$ = "IUABuiltUpWeb:DepthExt"
        Private Const IUABuiltUpIHaunch_LengthStart$ = "IUABuiltUpIHaunch:LengthStart"
        Private Const IUABuiltUpIHaunch_DepthStart$ = "IUABuiltUpIHaunch:DepthStart"
        Private Const IUABuiltUpIHaunch_LengthEnd$ = "IUABuiltUpIHaunch:LengthEnd"
        Private Const IUABuiltUpIHaunch_DepthEnd$ = "IUABuiltUpIHaunch:DepthEnd"
        Private Const IUABuiltUpIHaunch_DepthHaunch$ = "IUABuiltUpIHaunch:DepthHaunch"
        Private Const IUABuiltUpITaperWeb_DepthStart$ = "IUABuiltUpITaperWeb:DepthStart"
        Private Const IUABuiltUpITaperWeb_DepthEnd$ = "IUABuiltUpITaperWeb:DepthEnd"
        Private Const IUABuiltUpL_OffsetWeb$ = "IUABuiltUpL:OffsetWeb"
        Private Const IUABuiltUpC_OffsetWebBot$ = "IUABuiltUpC:OffsetWebBot"
        Private Const IUABuiltUpC_OffsetWebTop$ = "IUABuiltUpC:OffsetWebTop"
        Private Const IJUAC_eo_x$ = "IJUAC:eo_x"
        Private Const IJUAL_xp$ = "IJUAL:xp"
        Private Const IJUAL_yp$ = "IJUAL:yp"
        Private Const IUABuiltUpTube_TubeDiameter$ = "IUABuiltUpTube:TubeDiameter"
        Private Const IUABuiltUpTube_TubeThickness$ = "IUABuiltUpTube:TubeThickness"
        Private Const IUABuiltUpCan_DiameterStart$ = "IUABuiltUpCan:DiameterStart"
        Private Const IUABuiltUpCan_DiameterEnd$ = "IUABuiltUpCan:DiameterEnd"
        Private Const IUABuiltUpCone_DiameterStart$ = "IUABuiltUpCone:DiameterStart"
        Private Const IUABuiltUpCone_DiameterEnd$ = "IUABuiltUpCone:DiameterEnd"
        Private Const IUABuiltUpCone1_Cone1Thickness$ = "IUABuiltUpCone1:Cone1Thickness"
        Private Const IUABuiltUpCone_ConeThickness$ = "IUABuiltUpCone:ConeThickness"
        Private Const IUABuiltUpCone2_Cone2Thickness$ = "IUABuiltUpCone2:Cone2Thickness"
        Private Const IUABUBoxFlangeMajor_OffsetLeftWebTop$ = "IUABUBoxFlangeMajor:OffsetLeftWebTop"
        Private Const IUABUBoxFlangeMajor_OffsetRightWebTop$ = "IUABUBoxFlangeMajor:OffsetRightWebTop"
        Private Const IUABUBoxFlangeMajor_OffsetLeftWebBot$ = "IUABUBoxFlangeMajor:OffsetLeftWebBot"
        Private Const IUABUBoxFlangeMajor_OffsetRightWebBot$ = "IUABUBoxFlangeMajor:OffsetRightWebBot"
        Private Const IStructAngleBoltGage_lsg$ = "IStructAngleBoltGage:lsg"
        Private Const IStructAngleBoltGage_lsg1$ = "IStructAngleBoltGage:lsg1"
        Private Const IStructAngleBoltGage_lsg2$ = "IStructAngleBoltGage:lsg2"
        Private Const IStructAngleBoltGage_ssg$ = "IStructAngleBoltGage:ssg"
        Private Const IStructAngleBoltGage_ssg1$ = "IStructAngleBoltGage:ssg1"
        Private Const IStructAngleBoltGage_ssg2$ = "IStructAngleBoltGage:ssg2"
        Private Const IUABuiltUpCompute_IsModifiable$ = "IUABuiltUpCompute:IsModifiable"
        Private Const IUABuiltUpCompute_SectionProperties$ = "IUABuiltUpCompute:SectionProperties"
        Private Const IUABuiltUpLengthExt_LengthExt$ = "IUABuiltUpLengthExt:LengthExt"
        Private Const IUABuiltUpEndCan_IsTransitionAtStart$ = "IUABuiltUpEndCan:IsTransitionAtStart"
        Private Const IUABuiltUpEndCan_DiameterCone$ = "IUABuiltUpEndCan:DiameterCone"
        Private CSType$

#End Region

#Region "Public methods and functions"

        ''' <summary>
        ''' Calculates the section design properties.
        ''' The cases, according to the cross-section type are defined in which, all the required properties to define a cross section are taken as inputs 
        ''' and the same are used for the calculation of cross-section design properties.
        ''' </summary>
        ''' <param name="inputProperties">Input properties can be depth,width, etc which are necessary in calculating properties.</param>
        ''' <param name="crossSectionType">Cross-section type name.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Public Function CalculateCrossSectionProperties(ByVal inputProperties As Dictionary(Of String, Object), _
                ByVal crossSectionType$) As Dictionary(Of String, Double) Implements ISectionCalculator.CalculateSectionProperties

            'Checking the inputs
            If String.IsNullOrEmpty(crossSectionType) Then
                Throw New ArgumentNullException("crossSectionType")
            End If

            If inputProperties Is Nothing Then
                Throw New ArgumentNullException("inputProperties")
            End If

            Select Case crossSectionType
                Case "W", "M", "HP"
                    Return CalculateISectionProperties(inputProperties)
                Case "S"
                    Return CalculateSSectionProperties(inputProperties)
                Case "WT", "MT"
                    Return CalculateTSectionProperties(inputProperties)
                Case "ST"
                    Return CalculateSTSectionProperties(inputProperties)
                Case "C", "MC"
                    Return CalculateCSectionProperties(inputProperties)
                Case "L"
                    Return CalculateLSectionProperties(inputProperties)
                Case "2L"
                    Return Calculate2LSectionProperties(inputProperties)
                Case "HSSR"
                    Return CalculateHSSRSectionProperties(inputProperties)
                Case "HSSC", "PIPE"
                    Return CalculatePIPESectionProperties(inputProperties)
                Case "CS"
                    Return CalculateCSSectionProperties(inputProperties)
                Case "RS"
                    Return CalculateRSSectionProperties(inputProperties)
                Case "BUI"
                    Return CalculateBUISectionProperties(inputProperties)
                Case "BUIUE"
                    Return CalculateBUIUESectionProperties(inputProperties)
                Case "BUITapWeb"
                    Return CalculateBUITapWebSectionProperties(inputProperties)
                Case "BUIHaunch"
                    Return CalculateBUIHaunchSectionProperties(inputProperties)
                Case "BUL"
                    Return CalculateBULSectionProperties(inputProperties)
                Case "BUC"
                    Return CalculateBUCSectionProperties(inputProperties)
                Case "BUBoxFM"
                    Return CalculateBUBoxFMSectionProperties(inputProperties)
                Case "BUTube"
                    Return CalculateBUTubeSectionProperties(inputProperties)
                Case "BUCone"
                    Return CalculateBUConeSectionProperties(inputProperties)
                Case "BUCan"
                    Return CalculateBUCanSectionProperties(inputProperties)
                Case "BUEndCan"
                    Return CalculateBUEndCanSectionProperties(inputProperties)
                Case "BUFlat"
                    Return CalculateBUFlatSectionProperties(inputProperties)
                Case "BUTee"
                    Return CalculateBUTeeSectionProperties(inputProperties)
                Case Else
                    Throw New CmnException(SectionLibraryCalculatorLocalizer.GetString(SectionLibraryCalculatorResourceIDs.MSG_MISSING_CSTYPE1, "Section property calculations are not available for '") _
                                           + crossSectionType + SectionLibraryCalculatorLocalizer.GetString(SectionLibraryCalculatorResourceIDs.MSG_MISSING_CSTYPE2, "' cross-section part class. Please update calculator project to handle property calculations for '") _
                                           + crossSectionType + SectionLibraryCalculatorLocalizer.GetString(SectionLibraryCalculatorResourceIDs.MSG_MISSING_CSTYPE3, "' cross-section part class or specify a different calculator ProgID in <SYMBOL_SHARE>\Xml\Structure\CrossSectionTypeProperties.xml.") + " " _
                                           + SectionLibraryCalculatorLocalizer.GetString(SectionLibraryCalculatorResourceIDs.MSG_MISSING_CSTYPE4, "For more information on how to add a new cross-section type in the calculator project, please refer to the Structure Reference Data help."))
            End Select

        End Function

        ''' <summary>
        ''' Validates the minimum required input properties.
        ''' </summary>
        ''' <param name="inputProperties">The minimum required input properties.</param>
        ''' <param name="crossSectionType">Cross-section type name.</param>
        ''' <param name="errorMessage">Returned error message to indicate why the property is not valid.</param>
        '''<returns>True if properties are valid, False otherwise.</returns>
        Public Function ValidateInputProperties(ByVal inputProperties As Dictionary(Of String, Object), ByVal crossSectionType$, ByRef errorMessage$) As Boolean Implements ISectionCalculator.ValidateInputProperties

            'Checking the inputs
            If String.IsNullOrEmpty(crossSectionType) Then
                Throw New ArgumentNullException("crossSectionType")
            End If

            If inputProperties Is Nothing Then
                Throw New ArgumentNullException("inputProperties")
            End If

            CSType = crossSectionType

            Select Case crossSectionType
                Case "W", "M", "HP"
                    errorMessage = ValidateISectionProperties(inputProperties)
                Case "S"
                    errorMessage = ValidateSSectionProperties(inputProperties)
                Case "WT", "MT"
                    errorMessage = ValidateTSectionProperties(inputProperties)
                Case "ST"
                    errorMessage = ValidateSTSectionProperties(inputProperties)
                Case "C", "MC"
                    errorMessage = ValidateCSectionProperties(inputProperties)
                Case "L"
                    errorMessage = ValidateLSectionProperties(inputProperties)
                Case "2L"
                    errorMessage = Validate2LSectionProperties(inputProperties)
                Case "HSSR"
                    errorMessage = ValidateHSSRSectionProperties(inputProperties)
                Case "HSSC", "PIPE"
                    errorMessage = ValidatePIPESectionProperties(inputProperties)
                Case "CS"
                    errorMessage = ValidateCSSectionProperties(inputProperties)
                Case "RS"
                    errorMessage = ValidateRSSectionProperties(inputProperties)
                Case "BUI"
                    errorMessage = ValidateBUISectionProperties(inputProperties)
                Case "BUIUE"
                    errorMessage = ValidateBUIUESectionProperties(inputProperties)
                Case "BUITapWeb"
                    errorMessage = ValidateBUITapWebSectionProperties(inputProperties)
                Case "BUIHaunch"
                    errorMessage = ValidateBUIHaunchSectionProperties(inputProperties)
                Case "BUL"
                    errorMessage = ValidateBULSectionProperties(inputProperties)
                Case "BUC"
                    errorMessage = ValidateBUCSectionProperties(inputProperties)
                Case "BUBoxFM"
                    errorMessage = ValidateBUBoxFMSectionProperties(inputProperties)
                Case "BUTube"
                    errorMessage = ValidateBUTubeSectionProperties(inputProperties)
                Case "BUCone"
                    errorMessage = ValidateBUConeSectionProperties(inputProperties)
                Case "BUCan"
                    errorMessage = ValidateBUCanSectionProperties(inputProperties)
                Case "BUEndCan"
                    errorMessage = ValidateBUEndCanSectionProperties(inputProperties)
                Case "BUFlat"
                    errorMessage = ValidateBUFlatSectionProperties(inputProperties)
                Case "BUTee"
                    errorMessage = ValidateBUTeeSectionProperties(inputProperties)
                Case Else
                    errorMessage = String.Empty
            End Select

            'If we have an error, we should return false else true
            If errorMessage Is String.Empty Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Private methods and functions"

        ''' <summary>
        ''' Function to calculate I section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia and Warping constant
        ''' Properties like theoretical yield stress, normalized warping constant, Warping statical moment, 
        ''' Statical moment at point in flange and web mid-depth are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties can be depth, width, etc which are necessary in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateISectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))

            'In I section, flange width will be the width  
            Dim flangeWidth# = width
            outputs.Add(IStructFlangedSectionDimensions_bf, flangeWidth)
            outputs.Add(IStructFlangedSectionDimensions_d, depth)

            'Total area is calculated by summing up the individual areas of the I section (flanges(top and bottom), web)
            Dim area# = 2 * width * flangeThickness + webThickness * (depth - 2 * flangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Total perimeter is calculated by summing up the individual perimeters of the I section (flanges(top and bottom), web)   
            Dim perimeter# = 2 * width + 4 * flangeThickness + 2 * (depth - 2 * flangeThickness) + 2 * (width - webThickness)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'I section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is always symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Centroid about X and Y are calculated according to the symmetry of the section.
            'Since I section is symmetric about X and Y axis, centroid about X and Y can be calculated easily
            'In this case, they are width / 2, depth / 2 about X and Y respectively

            Dim centroid_X# = width / 2
            Dim centroid_Y# = depth / 2
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of moment of inertia about X axis
            Dim momentOfInertia_X# = ((width * Math.Pow(depth, 3)) - ((width - webThickness) * Math.Pow((depth - 2 * flangeThickness), 3))) / 12
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            Dim momentOfInertia_Y# = (2 * flangeThickness * Math.Pow(width, 3) + (depth - 2 * flangeThickness) * Math.Pow(webThickness, 3)) / 12
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'I section is symmetric about both X and Y axis, elastic section modulus is just the moment of inertia about particular axis (X or Y) upon
            'distance from edge to neutral axis respectively.
            Dim elasticSectionModulus_X# = 2 * (momentOfInertia_X / depth)
            Dim elasticSectionModulus_Y# = 2 * (momentOfInertia_Y / width)
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'The following are the relations/formulae used to calculate plastic section modulus (about X and Y axis)
            Dim plasticSectionModulus_X# = width * flangeThickness * (depth - flangeThickness) + webThickness * Math.Pow((0.5 * depth - flangeThickness), 2)
            Dim plasticSectionModulus_Y# = 0.5 * (width * width) * flangeThickness + 0.25 * (webThickness * webThickness * (depth - 2 * flangeThickness))
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'The following are the relations/formulae used to calculate torsional moment of inertia and warping constant
            Dim torsional_MomentOfInertia# = ((depth - 2 * flangeThickness) * Math.Pow(webThickness, 3) + 2 * width * Math.Pow(flangeThickness, 3)) / 3
            Dim warpingConstant# = ((depth - flangeThickness) * (depth - flangeThickness) * flangeThickness * Math.Pow(width, 3)) / 24
            outputs.Add(IStructCrossSectionDesignProperties_J, torsional_MomentOfInertia)
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)

            'The values of flange gage bolt and web gage bolt are edited by the user interactively
            'These default values are set to zero at present 
            Dim flangeGage# = 0, webGage# = 0
            outputs.Add(IStructFlangedBoltGage_gf, flangeGage)
            outputs.Add(IStructFlangedBoltGage_gw, webGage)

            CalculateISectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate S section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia and Warping constant
        ''' Properties like Theoretical yield stress, normalized warping constant, Warping statical moment, 
        ''' Statical moment at point in flange and web mid-depth are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties can be depth, width, etc which are necessary in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateSSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))
            Dim kdesign# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_kdesign))

            'In S section, flange width will be the width  
            Dim flangeWidth# = width
            outputs.Add(IStructFlangedSectionDimensions_bf, flangeWidth)
            outputs.Add(IStructFlangedSectionDimensions_d, depth)

            'Total area is calculated by summing up the individual areas of the S section (flanges(top and bottom), web, etc)  
            Dim area# = 2 * (0.5 * 2 * (kdesign - flangeThickness) * (webThickness + width)) + 2 * (width * (2 * flangeThickness - kdesign)) + (depth - (2 * kdesign)) * webThickness
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Total perimeter is calculated by summing up the individual perimeters of the S section (flanges(top and bottom), web etc)  
            Dim perimeter# = (2 * width) + (2 * (depth - (2 * kdesign))) + (8 * flangeThickness - 4 * kdesign) + 4 * Math.Sqrt(((0.5 * (width - webThickness) * 0.5 * (width - webThickness)) + (2 * (kdesign - flangeThickness)) * (2 * (kdesign - flangeThickness))))
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required 

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'S section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is always symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAbouty As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutX))

            'Slope is basically the ratio of (difference of height & length) to (difference of width & webThickness)
            Dim height# = (depth - (kdesign - (0.083 * (width - webThickness))))
            Dim length# = (depth - 2 * kdesign)
            Dim slope# = (height - length) / (width - webThickness)

            'Calculation of moment of inertia about X
            Dim momentOfInertia_X# = 1 / 12 * ((width * Math.Pow(depth, 3)) - ((0.25 / slope) * (Math.Pow(height, 4) - Math.Pow(length, 4))))
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y
            Dim momentOfInertia_Y# = 1 / 12 * ((Math.Pow(width, 3) * (depth - height)) - (length * Math.Pow(webThickness, 3)) + 0.25 * slope * (Math.Pow(width, 4) - Math.Pow(webThickness, 4)))
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'S-section is symmetric about both its axis, elastic section modulus is just the moment of inertia about particular axis(X or Y) upon
            'distance from edge to neutral axis respectively.
            Dim elasticSectionModulus_X# = 2 * (momentOfInertia_X / depth)
            Dim elasticSectionModulus_Y# = 2 * (momentOfInertia_Y / width)
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) upon the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'The following are the relations/formulae used to calculate plastic section modulus (about X and Y axis)
            Dim plasticSectionModulus_X# = width * flangeThickness * (depth - flangeThickness) + webThickness * Math.Pow((0.5 * depth - flangeThickness), 2)
            Dim plasticSectionModulus_Y# = 0.5 * (width * width) * flangeThickness + 0.25 * (webThickness * webThickness * (depth - 2 * flangeThickness))
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia
            Dim torsional_MomentOfInertia# = ((depth - 2 * flangeThickness) * Math.Pow(webThickness, 3) + 2 * width * Math.Pow(flangeThickness, 3)) / 3

            'Calculation of warping constant
            Dim warpingConstant# = ((depth - 2 * flangeThickness) * (depth - 2 * flangeThickness) * flangeThickness * Math.Pow(width, 3)) / 24
            outputs.Add(IStructCrossSectionDesignProperties_J, torsional_MomentOfInertia)
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)

            'The values of flange gage bolt and web gage bolt are edited by the user interactively
            'These default values are set to zero at present 
            Dim flangeGage# = 0, webGage# = 0
            outputs.Add(IStructFlangedBoltGage_gf, flangeGage)
            outputs.Add(IStructFlangedBoltGage_gw, webGage)

            CalculateSSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate T section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis,
        ''' Moment of inertia about X-axis and Y-axis,Plastic and Elastic section modulus about X-axis and Y-axis,
        ''' Radius of gyration about X-axis and Y-axis,Torsional moment of inertia and Warping constant,Flexural constant,
        ''' Polar radius of gyration about shear center. 
        ''' </summary>
        ''' <param name="inputProperties">Input properties can be depth, width, etc which are necessary in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateTSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))

            'In T section, flange width will be the width 
            Dim flangeWidth# = width
            outputs.Add(IStructFlangedSectionDimensions_bf, flangeWidth)
            outputs.Add(IStructFlangedSectionDimensions_d, depth)

            'Total area is calculated by summing up the individual areas of the T section (flange, web)
            Dim area# = (width * flangeThickness) + (webThickness * (depth - flangeThickness))
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Total perimeter is calculated by summing up the individual perimeters of the T section (flange, web, fillets)  
            Dim perimeter# = width + 2 * flangeThickness + (width - webThickness) + 2 * (depth - flangeThickness) + webThickness
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'T section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is symmetric about only Y axis.
            Dim isSymmetryAboutX As Boolean
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Individual entity areas for centroidY calculation
            Dim A1# = (depth - flangeThickness) * webThickness
            Dim A2# = width * flangeThickness

            'CentroidY is calculated as follows
            'Individual element centroid along Y
            'While calculating individual centroid, incremental distance must be considered from the reference axis

            Dim y1 = (depth - flangeThickness) / 2
            Dim y2 = (depth - flangeThickness) + (flangeThickness / 2)

            Dim centroid_Y# = (A1 * y1 + A2 * y2) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, depth - centroid_Y)

            'Calculation of moment of inertia about X axis
            Dim momentOfInertia_X# = (webThickness * Math.Pow((depth - flangeThickness), 3) / 12) + A1 * Math.Pow((y1 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (width * Math.Pow(flangeThickness, 3) / 12) + A2 * Math.Pow((y2 - centroid_Y), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y axis
            Dim momentOfInertia_Y# = ((flangeThickness * Math.Pow(width, 3)) + (depth - 2 * flangeThickness) * Math.Pow(webThickness, 3)) / 12
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Vertical distance from designated member edge to plastic neutral axis
            Dim verticalDistance_PlasticNeutralAxis# = GetPlasticNeutralAxisDistanceForTSection(depth, width, webThickness, flangeThickness) 'width / 2
            outputs.Add(IJUAL_yp, verticalDistance_PlasticNeutralAxis)

            'T-section is not symmetric about X axis. So, in order to calculate the value of elastic section modulus,
            ' initially it is checked for some conditions
            Dim elasticSectionModulus_X#
            If (centroid_Y > 0.5 * depth) Then
                If (centroid_Y <> 0.0) Then
                    elasticSectionModulus_X = momentOfInertia_X / centroid_Y
                Else
                    If (depth - centroid_Y <> 0.0) Then
                        elasticSectionModulus_X = momentOfInertia_X / (depth - centroid_Y)
                    End If
                End If
            End If

            'T-section is symmetric about Y axis, elastic section modulus is just the moment of inertia about particular axis upon
            'vertical distance from edge of the member to neutral axis respectively.
            Dim elasticSectionModulus_Y# = momentOfInertia_Y / (0.5 * width)
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'The following are the relations/formulae used to calculate plastic section modulus (about X and Y axis)
            'In order to find the same, we need to check by input properties for satisfying some conditions.
            Dim m1#, m2#, plasticSectionModulus_X#
            m1 = 0.5 * (depth - width * flangeThickness / webThickness + flangeThickness)
            m2 = 0.5 * (width * flangeThickness + depth * webThickness - webThickness * flangeThickness) / flangeWidth
            If ((webThickness + flangeThickness) <> 0 And width <> 0.0) Then
                plasticSectionModulus_X = 0.5 * (flangeWidth * (m2 * m2) + (flangeWidth - webThickness) * (flangeThickness - m2) * (flangeThickness - m2) + (depth - m2) * (depth - m2) * webThickness)
            End If
            Dim plasticSectionModulus_Y# = 0.25 * (flangeWidth * flangeWidth * flangeThickness + webThickness * webThickness * (depth - flangeThickness))
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia
            Dim torsional_MomentOfInertia# = ((flangeWidth * flangeThickness * flangeThickness * flangeThickness) + (webThickness * webThickness * webThickness * (depth - flangeThickness))) / 3
            outputs.Add(IStructCrossSectionDesignProperties_J, torsional_MomentOfInertia)

            'The values of warpingConstant, polarRadiusOfGyration_ShearCenter, flexural_Constant are approximated
            'The below formula is used for calculating warping constant
            Dim warpingConstant# = ((Math.Pow(flangeWidth, 3) * Math.Pow(flangeThickness, 3)) / 4 + (Math.Pow((depth - flangeThickness), 3) * Math.Pow(webThickness, 3))) / 36
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)

            '=================================================================
            'The calculation of following properties are purely approximated
            '=================================================================

            'The below formula is used in calculating polar radius of gyration
            'For the calculation of the same, we need to find out the value of eccentric shear distance
            Dim eo_y# = (depth / 2) * (flangeThickness + depth) / (1 + ((flangeThickness * Math.Pow(flangeWidth, 3)) / ((depth - flangeThickness) * Math.Pow(webThickness, 3))))
            Dim r# = Math.Abs(centroid_Y - eo_y)

            Dim polarRadiusOfGyration_ShearCenter# = Math.Sqrt(Math.Pow(r, 2) + (momentOfInertia_X + momentOfInertia_Y) / area)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)

            'The following relation is used in calculating the flexural constant
            'This uses the value of polar radius of gyration about shear center in the calculation
            Dim flexural_Constant# = Math.Abs(1 - ((r * r) / (polarRadiusOfGyration_ShearCenter * polarRadiusOfGyration_ShearCenter)))
            outputs.Add(IStructCrossSectionDesignProperties_H, flexural_Constant)

            'The values of flange gage bolt and web gage bolt are edited by the user interactively
            'These default values are set to zero at present 
            Dim flangeGage# = 0, webGage# = 0
            outputs.Add(IStructFlangedBoltGage_gf, flangeGage)
            outputs.Add(IStructFlangedBoltGage_gw, webGage)

            CalculateTSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate ST section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia and Warping constant,
        ''' Flexural constant, Polar radius of gyration about shear center 
        ''' Properties like Plastic section modulus about X-axis and Y-axis, Torsional moment of inertia and Warping constant are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties can be depth, width, etc which are necessary in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateSTSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))
            Dim kdesign# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_kdesign))

            'In ST section, flange width will be the width 
            Dim flangeWidth# = width
            outputs.Add(IStructFlangedSectionDimensions_bf, flangeWidth)
            outputs.Add(IStructFlangedSectionDimensions_d, depth)

            'Total area is calculated by summing up the individual areas of the ST section (flange, web)
            'Here the angle is constant, it will be always 15 degrees.
            'Find out the individual areas for simplicity

            Dim a1# = webThickness * (depth - kdesign)
            Dim a2# = 0.5 * 0.134 * (width - webThickness) * width
            Dim a3# = width * (kdesign - 0.134 * (width - webThickness))
            Dim area# = a1 + a2 + a3
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Total perimeter is calculated by summing up the individual perimeters of the ST section (flange, web)
            Dim perimeter# = width + 2 * (kdesign - 0.134 * (width - webThickness)) + 2 * 0.8 * (width - webThickness) + 2 * (depth - 2 * kdesign)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'ST section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is symmetric about only Y axis.
            Dim isSymmetryAboutX As Boolean
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Calculation of centroid about X
            Dim centroid_X# = width / 2

            'Calculation of centroid about Y
            'Calculate the entities centroid for simplicity 

            Dim y1# = (depth - kdesign) / 2
            Dim y2# = (depth - kdesign) + ((width + 2 * webThickness) / 3 * (width + webThickness)) * 0.134 * (width - webThickness)
            Dim y3# = (depth - kdesign) + 0.134 * (width - webThickness) + webThickness / 2

            Dim centroid_Y# = (a1 * y1 + a2 * y2 + a3 * y3) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, depth - centroid_Y)

            'Vertical distance from designated member edge to plastic neutral axis
            Dim verticalDistance_PlasticNeutralAxis# = GetPlasticNeutralAxisDistanceForTSection(depth, width, webThickness, flangeThickness) '(centroid_Y - 0.5 * depth) / 2
            outputs.Add(IJUAL_yp, verticalDistance_PlasticNeutralAxis)

            'Calculation of moment of inertia about X
            'Parallel axis theorem is used in calculating the overall moment of inertia about X

            Dim momentOfInertia_X# = (1 / 12) * webThickness * Math.Pow((depth - kdesign), 3) + a1 * Math.Pow((y1 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / (36 * (width + webThickness))) * Math.Pow((0.134 * (width - webThickness)), 3) * (Math.Pow(width, 2) + 4 * width * width + Math.Pow(webThickness, 2))
            momentOfInertia_X = momentOfInertia_X + a2 * Math.Pow((y2 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * width * Math.Pow(webThickness, 3) + a3 * Math.Pow((y3 - centroid_Y), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y

            Dim momentOfInertia_Y# = (1 / 12) * (depth - kdesign) * Math.Pow(webThickness, 3)
            momentOfInertia_Y = momentOfInertia_Y + (0.134 * (width - webThickness) / 48) * (width + webThickness) * (Math.Pow(width, 2) * Math.Pow(webThickness, 2))
            momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * webThickness * Math.Pow(width, 3)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Elastic section modulus calculations
            'The denominator in the formula checks the maximum value of (centroid_Y, (depth - centroid_Y))in case of X 
            'And (centroid_X, (width - centroid_X)) in case of Y and calculates the elastic section modulus 
            Dim elasticSectionModulus_X# = momentOfInertia_X / Max(centroid_Y, (depth - centroid_Y))
            Dim elasticSectionModulus_Y# = momentOfInertia_Y / Max(centroid_X, (width - centroid_X))
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            '=================================================================
            'The calculations of following properties are purely approximated
            '=================================================================

            'The below formula is used in calculating polar radius of gyration about shear center
            'For the calculation of the same, we need to find out the value of eccentric shear distance about y
            Dim eo_y# = (depth / 2) * (flangeThickness + depth) / (1 + ((flangeThickness * Math.Pow(width, 3)) / ((depth - flangeThickness) * Math.Pow(webThickness, 3))))
            Dim r# = Math.Abs(centroid_Y - eo_y)
            Dim polarRadiusOfGyration_ShearCenter# = Math.Sqrt(Math.Pow(r, 2) + (momentOfInertia_X + momentOfInertia_Y) / area)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)

            'The following relation is used in calculating the flexural constant
            'This uses the value of polar radius of gyration about shear center
            Dim flexural_Constant# = Math.Abs(1 - ((r * r) / (polarRadiusOfGyration_ShearCenter * polarRadiusOfGyration_ShearCenter)))
            outputs.Add(IStructCrossSectionDesignProperties_H, flexural_Constant)

            'The values of flange gage bolt and web gage bolt are edited by the user interactively
            'These default values are set to zero at present 
            Dim flangeGage# = 0, webGage# = 0
            outputs.Add(IStructFlangedBoltGage_gf, flangeGage)
            outputs.Add(IStructFlangedBoltGage_gw, webGage)

            CalculateSTSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate C section properties like area, perimeter, unit weight, Centroid about X-axis, Eccentric shear distance,
        ''' Horizontal distance from designated member to plastic neutral axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis
        ''' Torsional moment of inertia and Warping constant, Flexural constant and Polar radius of gyration about shear center
        ''' </summary>
        ''' <param name="inputProperties">Input properties like depth, width, flange and web thickness, k-design/k-detail are necessary in calculating properties</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateCSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))

            'In C section, flange width will be the width 
            Dim flangewidth# = width
            outputs.Add(IStructFlangedSectionDimensions_bf, flangewidth)
            outputs.Add(IStructFlangedSectionDimensions_d, depth)

            'Total area is calculated by summing up the individual area of the sections (flanges, web)
            Dim area# = 2 * width * flangeThickness + webThickness * (depth - 2 * flangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Total perimeter is calculated by summing up the individual perimeters of the section (flanges, web, fillets)
            Dim perimeter# = 4 * width + 4 * flangeThickness + 2 * (depth - 2 * flangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'C section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is symmetric about only X axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAboutY As Boolean

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Calculation of centroid about X
            'For our simplicity, find out the individual areas

            Dim A1# = (width * flangeThickness)
            Dim A2# = (depth - 2 * flangeThickness) * webThickness
            Dim A3# = (width * flangeThickness)

            'Calculate the individaul entity centroid about X from the reference axis
            'Considering bottom flange, web, top flange as a sequence

            Dim x1# = (width / 2)
            Dim x2# = (webThickness / 2)
            Dim x3# = (width / 2)

            'Method followed is Sum of each(Entity Area* Entity Centroid)/Total Area

            Dim centroid_X# = (A1 * x1 + A2 * x2 + A3 * x3) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)

            'Calculation of centroid about Y
            'Calculate the individaul entity centroid about Y

            Dim y1# = flangeThickness / 2
            Dim y2# = flangeThickness + (depth - 2 * flangeThickness) / 2
            Dim y3# = flangeThickness + (depth - 2 * flangeThickness) + flangeThickness / 2

            'Method followed is Sum of each(Entity Area* Entity Centroid)/Total Area
            Dim centroid_Y# = (A1 * y1 + A2 * y2 + A3 * y3) / area

            'Calculation of eccentric shear distance
            Dim C1# = width - webThickness / 2
            Dim C2# = 2 + ((webThickness / flangeThickness) * (depth - flangeThickness) / (width - webThickness / 2)) / 3
            Dim eccentricShearDistance# = Math.Abs((-1) * (C1 / C2 - (0.5 * webThickness)))
            outputs.Add(IJUAC_eo_x, eccentricShearDistance)

            'Calculation of moment of inertia about X axis
            Dim momentOfInertia_X# = ((width * Math.Pow(depth, 3)) - (width - webThickness) * Math.Pow((depth - 2 * flangeThickness), 3)) / 12
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y axis
            Dim momentOfInertia_Y# = (2 * flangeThickness * Math.Pow(width, 3) + (depth - 2 * flangeThickness) * Math.Pow(webThickness, 3)) / 3
            momentOfInertia_Y = momentOfInertia_Y - ((2 * flangeThickness * width + (depth - 2 * flangeThickness) * webThickness) * Math.Pow(centroid_X, 2))
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X is simply moment of inertia about X upon the vertical distance from designated member to plastic neutral axis
            Dim elasticSectionModulus_X# = 2 * (momentOfInertia_X / depth)

            'Calculation of elastic section modulus about Y
            'Checking for necessary conditions

            Dim elasticSectionModulus_Y#
            If (centroid_X > (0.5 * width)) Then
                elasticSectionModulus_Y# = (momentOfInertia_Y / centroid_X)
            Else
                elasticSectionModulus_Y = momentOfInertia_Y / Max(centroid_X, (width - centroid_X))
            End If
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'In calculating plastic section modulus about X, standard relation is used
            Dim plasticSectionModulus_X# = (width - webThickness) * flangeThickness * (depth - flangeThickness) + 0.25 * depth * depth * webThickness

            'The following are the relations/formula used to calculate plastic section modulus (about Y axis)
            'In order to find the same, we need to check by input properties for satisfying some conditions.

            Dim m1# = 0.5 * (width - 0.5 * depth * webThickness / flangeThickness + webThickness)
            Dim m2# = (width * flangeThickness + 0.5 * depth * webThickness - flangeThickness * webThickness) / depth
            Dim plasticSectionModulus_Y#

            plasticSectionModulus_Y = flangeThickness * (width - m2) * (width - m2) + 0.5 * (depth - 2.0 * flangeThickness) * (m2 - webThickness) * (m2 - webThickness) + 0.5 * depth * m2 * m2

            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia and warping constants
            Dim torsionalMomentOfInertia# = ((depth - 2 * flangeThickness) * Math.Pow(webThickness, 3) + (2 * width * Math.Pow(flangeThickness, 3))) / 3
            Dim warpingConstant# = (depth - flangeThickness) * (depth - flangeThickness) * (width - (0.5 * webThickness)) * (width - (0.5 * webThickness)) * (width - (0.5 * webThickness)) * flangeThickness / 12.0 * (2.0 * (depth - flangeThickness) * webThickness + 3.0 * (width - (0.5 * webThickness)) * flangeThickness) / ((depth - flangeThickness) * webThickness + 6.0 * (width - (0.5 * webThickness)) * flangeThickness)
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)

            '=================================================================
            'The calculations of following properties are purely approximated
            '=================================================================

            'The below formula is used in calculating polar radius of gyration about shear center
            'This uses the value of eccentric shear distance in its calculation
            Dim t# = Math.Abs(centroid_Y - eccentricShearDistance)
            Dim polarRadiusOfGyration_ShearCenter# = Math.Sqrt((t * t) + (momentOfInertia_X + momentOfInertia_Y) / area)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)

            'The following relation is used in calculating the flexural constant
            'This uses the value of polar radius of gyration about shear center
            Dim flexuralConstant# = Math.Abs(1 - (t * t) / (polarRadiusOfGyration_ShearCenter * polarRadiusOfGyration_ShearCenter))
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)

            'The values of flange gage bolt and web gage bolt are edited by the user interactively
            'These default values are set to zero at present 
            Dim flangeGage# = 0, webGage# = 0
            outputs.Add(IStructFlangedBoltGage_gf, flangeGage)
            outputs.Add(IStructFlangedBoltGage_gw, webGage)

            CalculateCSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate L section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Horizontal and vertical distances from designated member to plastic neutral axis, Plastic and Elastic section modulus about X-axis and Y-axis
        ''' Radius of gyration about X-axis and Y-axis, Torsional moment of inertia,Flexural constant and Polar radius of gyration about shear center 
        ''' Properties like Warping Constant is left for editing according to user calculation whose default value is set to zero
        ''' </summary>
        ''' <param name="inputProperties">Input properties like depth, width, flange and web thickness, k-design/k-detail are necessary in calculating properties</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateLSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))

            'In L section, flange width will be the width 
            Dim flangeWidth# = width
            outputs.Add(IStructFlangedSectionDimensions_bf, flangeWidth)
            outputs.Add(IStructFlangedSectionDimensions_d, depth)

            'Total area is calculated by summing up the individual areas of the section (flange, web)
            Dim area# = flangeThickness * width + webThickness * (depth - flangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Total perimeter is calculated by summing up the individual perimeters of the section (flange, web)
            Dim perimeter# = webThickness + depth + width + flangeThickness + (width - webThickness) + (depth - flangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'L section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is neither symmetric about X nor Y axis 
            Dim isSymmetryAboutX As Boolean
            Dim isSymmetryAboutY As Boolean

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Calculation of centroid about X
            'Calculate the entity areas for simplicity
            Dim A1# = (width * flangeThickness)
            Dim A2# = (depth - flangeThickness) * webThickness

            'Calculate entities centroid about X
            Dim X1# = width / 2
            Dim X2# = webThickness / 2

            'Overall centroid about X
            Dim centroid_X# = ((A1 * X1) + (A2 * X2)) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)

            'Calculation of centroid about Y
            'Calculate entities centroid about Y

            Dim y1 = flangeThickness / 2
            Dim y2 = flangeThickness + (depth - flangeThickness) / 2

            'Overall centroid about Y
            Dim centroid_Y# = ((A1 * y1) + (A2 * y2)) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of moment of inertia about X 

            Dim momentOfInertia_X# = width * Math.Pow(flangeThickness, 3) / 12 + A1 * Math.Pow(y1 - centroid_Y, 2) 'Flange MI
            momentOfInertia_X = momentOfInertia_X + webThickness * Math.Pow((depth - flangeThickness), 3) / 12 + A2 * Math.Pow((y2 - centroid_Y), 2) 'Flange MI + Web MI

            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y 
            Dim momentOfInertia_Y# = flangeThickness * Math.Pow(width, 3) / 12 + A1 * Math.Pow((X1 - centroid_X), 2) 'Flange MI
            momentOfInertia_Y = momentOfInertia_Y + (depth - flangeThickness) * Math.Pow(webThickness, 3) / 12 + A2 * Math.Pow((X2 - centroid_X), 2) 'Flange MI + Web MI

            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of radius of gyration about XY
            'This involves the calculation of moment of inertia about XY
            'Standard formula is used in calculating the same
            Dim momentOfInertia_XY# = depth * webThickness * (centroid_X - webThickness / 2) * (centroid_Y - depth / 2) + (width - webThickness) * flangeThickness * (centroid_X - (width + webThickness) / 2) * (centroid_Y - flangeThickness / 2)
            Dim principle_y# = 0.5 * (momentOfInertia_X + momentOfInertia_Y) - Math.Sqrt(0.25 * ((momentOfInertia_X - momentOfInertia_Y) * (momentOfInertia_X - momentOfInertia_Y) + (momentOfInertia_XY * momentOfInertia_XY)))
            Dim radiusOfGyration_XY# = Math.Sqrt(principle_y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)

            'Calculation of elastic section modulus about X
            'Checking for necessary conditions
            Dim elasticSectionModulus_X#
            If (centroid_Y < 0.5 * depth) Then
                'If (centroid_Y <> 0.0) Then
                '    elasticSectionModulus_X = momentOfInertia_X / centroid_Y
                'Else
                If ((depth - centroid_Y) <> 0.0) Then
                    elasticSectionModulus_X = momentOfInertia_X / (depth - centroid_Y)
                End If
                'End If
            End If
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)

            'Calculation of elastic section modulus about Y
            'Checking for necessary conditions
            Dim elasticSectionModulus_Y#
            If centroid_X > 0.5 * width Then
                elasticSectionModulus_Y = momentOfInertia_Y / centroid_X
            Else
                elasticSectionModulus_Y = momentOfInertia_Y / (width - centroid_X)
            End If
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of horizontal distance from designated member edge to plastic neutral axis
            'It can be either the maximum value of CentroidY or depth-CentroidY  
            Dim horizontalDistance_PlasticNeutralAxis# = GetPlasticNeutralAxisDistanceForTSection(width, depth, flangeThickness, webThickness) 'Max(centroid_Y, (depth - centroid_Y))
            outputs.Add(IJUAL_xp, horizontalDistance_PlasticNeutralAxis)

            'Calculation of vertical distance from designated member edge to plastic neutral axis
            'It can be either the maximum value of CentroidX or width-CentroidX  
            Dim verticalDistance_PlasticNeutralAxis# = GetPlasticNeutralAxisDistanceForTSection(depth, width, webThickness, flangeThickness) 'Max(centroid_X, (width - centroid_X))
            outputs.Add(IJUAL_yp, verticalDistance_PlasticNeutralAxis)

            'Calculation of torsional moment of inertia
            'Checking for all necessary conditions
            Dim j9#, ctr#, t2#, xj#, Z#
            Dim t1# = 0.0

            Dim A11# = Min(depth, webThickness)
            Dim B1# = Max(depth, webThickness)
            If (A11 <> 0.0) Then
                If (B1 / A11 > 29.0) Then
                    j9 = Math.Pow(A11, 3.0) * B1 / 3.0
                Else
                    For ctr = 1 To 5 Step 2
                    Next
                    t2 = 0.5 * Math.PI * ctr * B1 / A11
                    t1 = t1 + Math.Pow(ctr, -5.0) * Math.Tanh(t2)
                    j9 = Math.Pow(A11, 3.0) * B1 / 3.0 * (1.0 - (192.0 * (A11 / B1) * t1 / Math.Pow(Math.PI, 5.0)))
                End If
            End If

            xj = j9

            A11 = Min(flangeThickness, width - webThickness)
            B1 = Max(flangeThickness, width - webThickness)

            If (A11 <> 0.0) Then

                If (B1 / A11 > 29.0) Then
                    j9 = Math.Pow(A11, 3.0) * B1 / 3.0
                Else

                    For ctr = 1 To 5 Step 2
                    Next
                    t2 = 0.5 * Math.PI * ctr * B1 / A11
                    t1 = t1 + Math.Pow(ctr, -5.0) * Math.Tanh(t2)
                    j9 = Math.Pow(A11, 3.0) * B1 / 3.0 * (1.0 - (192.0 * (A11 / B1) * t1 / Math.Pow(Math.PI, 5.0)))
                End If
            End If

            xj = xj + j9
            Z = j9

            A11 = Min(width, flangeThickness)
            B1 = Max(width, flangeThickness)

            If (A11 <> 0.0) Then

                If (B1 / A11 > 29.0) Then
                    j9 = Math.Pow(A11, 3.0) * B1 / 3.0
                Else
                    t1 = 0.0
                    For ctr = 1 To 5 Step 2
                        t2 = 0.5 * Math.PI * ctr * B1 / A11
                        t1 = t1 + Math.Pow(ctr, -5.0) * Math.Tanh(t2)
                        j9 = Math.Pow(A11, 3.0) * B1 / 3.0 * (1.0 - (192.0 * (A11 / B1) * t1 / Math.Pow(Math.PI, 5.0)))
                    Next
                End If
            End If

            Dim torsionalMomentOfInertia# = Max(xj, Z + j9)
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            'Calculation of plastic section modulus about X
            'Checking for necessary conditions with the help of inputs
            Dim plasticSectionModulus_X#
            If ((webThickness + flangeThickness) <> 0 And width <> 0.0) Then
                Dim z1# = 0.5 * (depth - width * flangeThickness / webThickness + flangeThickness)
                Dim z2# = 0.5 * (width * flangeThickness + depth * webThickness - webThickness * flangeThickness) / width
                plasticSectionModulus_X = 0.5 * (width * z2 * z2 + (width - webThickness) * (flangeThickness - z2) * (flangeThickness - z2) + (depth - z2) * (depth - z2) * webThickness)
            End If
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)

            'Calculation of plastic section modulus about Y
            'Checking for necessary conditions with the help of inputs
            Dim x11# = 0.5 * (width - depth * webThickness / flangeThickness + webThickness)
            Dim x22# = 0.5 * (width * flangeThickness + depth * webThickness - flangeThickness * webThickness) / depth
            Dim plasticSectionModulus_Y#
            If (x11 >= webThickness And x11 < width) Then
                plasticSectionModulus_Y = 0.5 * flangeThickness * ((width - x11) * (width - x11) + (x11 * x11)) + (depth - flangeThickness) * webThickness * (x11 - 0.5 * webThickness)
            ElseIf (x22 < webThickness And x22 > 0.0) Then
                plasticSectionModulus_Y = 0.5 * (flangeThickness * (width - x22) * (width - x22) + (depth - flangeThickness) * (webThickness - x22) * (webThickness - x22) + depth * x22 * x22)

            End If
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            '=================================================================
            'The calculations of following properties are purely approximated
            '=================================================================

            'Calculation of polar radius of gyration about shear center
            Dim I1# = (depth - 0.5 * flangeThickness) * Math.Pow(webThickness, 3.0)
            Dim I2# = flangeThickness * Math.Pow((width - 0.5 * webThickness), 3.0)
            Dim eo_y# = 0.5 * (flangeThickness + (depth - 0.5 * flangeThickness)) * I1 / (I1 + I2)
            Dim v# = centroid_Y - eo_y
            Dim polarRadiusOfGyration_ShearCenter# = Math.Sqrt((v * v) + (momentOfInertia_X + momentOfInertia_Y) / area)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)

            'Calculation of flexural constant
            'This uses the value of polar moment of inertia about shear center
            Dim flexuralConstant# = Math.Abs(1 - (v * v) / (polarRadiusOfGyration_ShearCenter * polarRadiusOfGyration_ShearCenter))
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)

            'The following properties are left for edited by the user interactively
            'The default values for these properties are set as zero

            Dim warpingConstant# = 0
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)

            Dim longSideGage# = 0, longSideGage1# = 0, longSideGage2# = 0, shortSideGage# = 0, shortSideGage1# = 0, shortSideGage2# = 0
            outputs.Add(IStructAngleBoltGage_lsg, longSideGage)
            outputs.Add(IStructAngleBoltGage_lsg1, longSideGage1)
            outputs.Add(IStructAngleBoltGage_lsg2, longSideGage2)
            outputs.Add(IStructAngleBoltGage_ssg, shortSideGage)
            outputs.Add(IStructAngleBoltGage_ssg1, shortSideGage1)
            outputs.Add(IStructAngleBoltGage_ssg2, shortSideGage2)

            CalculateLSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate 2L section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Horizontal and vertical distances from designated member to plastic neutral axis, Plastic and Elastic section modulus about X-axis and Y-axis
        ''' Radius of gyration about X-axis and Y-axis, Torsional moment of inertia, Warping constant, Flexural constant and Polar radius of gyration about shear center 
        ''' </summary>
        ''' <param name="inputProperties">Input properties like flange width and thickness, depth, web thickness, back to back spacing, kdesign/kdetail are necessary in calculating properties</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function Calculate2LSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim flangeWidth# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_bf))
            Dim spacing_BackToBack# = CDbl(inputProperties.Item(IJUA2L_bb))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))

            'In 2L section, overall width is equal to twice of flange width plus back to back spacing
            Dim width# = (2 * flangeWidth) + spacing_BackToBack
            outputs.Add(IStructCrossSectionDimensions_Width, width)
            outputs.Add(IStructFlangedSectionDimensions_d, depth)

            'Area is calculated by considering the areas of individual segments of the whole cross section
            Dim area# = 2 * flangeWidth * flangeThickness + 2 * (depth - flangeThickness) * webThickness
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Perimeter is calculated by considering the perimeters of individual segments of the whole cross section
            'L perimeter is always equal to rectangle of same dimensions
            Dim perimeter# = 4 * (flangeWidth + depth) '2 * (webThickness + flangeThickness + flangeWidth + (depth - flangeThickness))
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            '2L section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is symmetric about Y axis only
            Dim isSymmetryAboutX As Boolean
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Calculation of centroid
            'Centroid X is basically calculated for its need in calculation of moment of inertia about X
            'It is not added in the output dictionary
            Dim centroid_X# = flangeWidth + (0.5 * spacing_BackToBack)

            'Calculation of centroid about Y
            'Calculate entities centroid about Y
            Dim A1# = (flangeWidth * flangeThickness)
            Dim A2# = (depth - flangeThickness) * webThickness
            Dim entityY1 = flangeThickness / 2
            Dim entityY2 = flangeThickness + (depth - flangeThickness) / 2

            'Overall centroid about Y
            Dim centroid_Y# = 2 * ((A1 * entityY1) + (A2 * entityY2)) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of moment of inertia about X 
            'Moment of inertia about X is calculated by considering the whole cross section as individual L sections
            Dim momentOfInertia_X# = 2.0 * (webThickness * Math.Pow(depth, 3) / 12.0 + (flangeWidth - webThickness) * Math.Pow(flangeThickness, 3.0) / 12.0 + depth * webThickness * Math.Pow((0.5 * depth - centroid_Y), 2))
            momentOfInertia_X = momentOfInertia_X + 2 * (flangeThickness * (flangeWidth - webThickness) * Math.Pow((centroid_Y - 0.5 * flangeThickness), 2))
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y
            'Moment of inertia about Y is not added in the output dictionary. It is simply used for the calculation of radius of gyration about Y
            Dim momentOfInertia_Y# = 2.0 * (flangeThickness * Math.Pow(flangeWidth, 3.0) / 12.0 + flangeThickness * flangeWidth * Math.Pow((flangeWidth / 2.0 + spacing_BackToBack / 2.0), 2.0) + (depth - flangeThickness) * Math.Pow(webThickness, 3.0) / 12.0)
            momentOfInertia_Y = momentOfInertia_Y + 2 * ((depth - flangeThickness) * webThickness * Math.Pow((webThickness / 2.0 + spacing_BackToBack / 2.0), 2.0))

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X
            'Before proceeding into the actual calculation, it is checked for some conditions with the help of inputs and calculated properties
            Dim elasticSectionModulus_X#
            If centroid_Y < 0.5 * depth Then
                If depth - centroid_Y <> 0.0 Then
                    elasticSectionModulus_X = momentOfInertia_X / (depth - centroid_Y)
                End If
            End If
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)

            'Calculation of plastic section modulus about X
            'Before proceeding into the actual calculation, it is checked for some conditions with the help of input properties
            Dim plasticSectionModulus_X#
            If (webThickness + flangeThickness <> 0 And flangeWidth <> 0.0) Then
                Dim y1# = 0.5 * (depth - flangeWidth * flangeThickness / webThickness + flangeThickness)
                Dim y2# = 0.5 * (flangeWidth * flangeThickness + depth * webThickness - webThickness * flangeThickness) / flangeWidth

                plasticSectionModulus_X = (flangeWidth * y2 * y2 + (flangeWidth - webThickness) * (flangeThickness - y2) * (flangeThickness - y2) + (depth - y2) * (depth - y2) * webThickness)

            End If
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)

            '=================================================================
            'The calculations of following properties are purely approximated
            '=================================================================

            'Calculation of polar radius of gyration about shear center
            'This involves the calculation of eccentric shear distance about Y
            Dim I1# = (depth - 0.5 * flangeThickness) * Math.Pow(webThickness, 3.0)
            Dim I2# = flangeThickness * Math.Pow((flangeWidth - 0.5 * webThickness), 3.0)
            Dim eo_y# = 0.5 * (flangeThickness + (depth - 0.5 * flangeThickness)) * I1 / (I1 + I2)
            Dim v# = centroid_Y - eo_y
            Dim polarRadiusOfGyration_ShearCenter# = Math.Sqrt((v * v) + (momentOfInertia_X + momentOfInertia_Y) / area)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)

            'Calculation of flexural constant
            'This uses the value of polar moment of inertia about shear center
            Dim flexuralConstant# = Math.Abs(1 - (v * v) / (polarRadiusOfGyration_ShearCenter * polarRadiusOfGyration_ShearCenter))
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)

            'The following properties are left for edited by the user interactively
            'The default values for these properties are set as zero
            Dim longSideGage# = 0, longSideGage1# = 0, longSideGage2# = 0, shortSideGage# = 0, shortSideGage1# = 0, shortSideGage2# = 0
            outputs.Add(IStructAngleBoltGage_lsg, longSideGage)
            outputs.Add(IStructAngleBoltGage_lsg1, longSideGage1)
            outputs.Add(IStructAngleBoltGage_lsg2, longSideGage2)
            outputs.Add(IStructAngleBoltGage_ssg, shortSideGage)
            outputs.Add(IStructAngleBoltGage_ssg1, shortSideGage1)
            outputs.Add(IStructAngleBoltGage_ssg2, shortSideGage2)

            Calculate2LSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate PIPE section properties like area, perimeter, unit weight, Moment of inertia about X-axis and Y-axis,
        ''' Elastic and Plastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia 
        ''' </summary>
        ''' <param name="inputProperties">Input properties width, depth, design/nominal thickness are helpful in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculatePIPESectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim thicknessNominal# = CDbl(inputProperties.Item(IJUAHSS_tnom))
            Dim thicknessDesign# = CDbl(inputProperties.Item(IJUAHSS_tdes))

            'PIPESection is always hollow
            Dim isHollow As Boolean = True

            'According to the geometry of the section, it is symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'In PIPE section width is equal to depth
            Dim width# = depth
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Calculation of depth upon thickness ratio
            Dim diameter_Thickness# = depth / thicknessNominal
            outputs.Add(IJUAHSSC_D_t, diameter_Thickness)

            'Inner diameter is the difference of depth and twice of thickness
            Dim innerDiameter# = depth - 2 * thicknessDesign

            'Calculation of area and perimeter
            'Standard formula is used in calculation of area and perimeter
            Dim area# = Math.PI * (depth * depth - innerDiameter * innerDiameter) / 4
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            Dim perimeter# = Math.PI * depth
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Since the section is symmetric about both X and Y axis, the values of moment of inertia about X and Y axis are equal
            'Standard formula is used in calculation of moment of inertia
            Dim momentOfInertia_X# = Math.PI * (depth * depth - innerDiameter * innerDiameter) * (depth * depth + innerDiameter * innerDiameter) / 64
            Dim momentOfInertia_Y# = momentOfInertia_X
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Since the section is symmetric about both X and Y axis, the values of radius of gyration about X and Y axis are equal
            'Standard formula is used in calculation of radius of gyration 
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Since the section is symmetric about both X and Y axis, the values of elastic section modulus about X and Y axis are equal
            'Standard formula is used in calculation of elastic section modulus
            Dim elasticSectionModulus_X# = momentOfInertia_X / (depth / 2)
            Dim elasticSectionModulus_Y# = elasticSectionModulus_X
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Since the section is symmetric about both X and Y axis, the values of plastic section modulus about X and Y axis are equal
            'Standard formula is used in calculation of plastic section modulus
            Dim plasticSectionModulus_X# = (depth * depth * depth - innerDiameter * innerDiameter * innerDiameter) / 6
            Dim plasticSectionModulus_Y# = plasticSectionModulus_X
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Torsional moment of inertia is twice of that of moment of inertia about X or Y axis
            Dim torsionalMomentOfInertia# = 2 * momentOfInertia_X
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            CalculatePIPESectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate HSSR section properties like area, perimeter, unit weight, Moment of inertia about X-axis and Y-axis,
        ''' Elastic and Plastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia 
        ''' </summary>
        ''' <param name="inputProperties">Input properties width, depth, design/nominal thickness are helpful in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateHSSRSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim thicknessDesign# = CDbl(inputProperties.Item(IJUAHSS_tdes))

            'HSSRSection is always hollow
            Dim isHollow As Boolean = True

            'According to the geometry of the section, it is symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Inner depth is the difference of depth and twice of thickness
            Dim innerDepth# = depth - 2 * thicknessDesign

            'Inner width is the difference of width and twice of thickness
            Dim innerWidth# = width - 2 * thicknessDesign

            'Calculation of area and perimeter
            'Standard formula is used in calculation of area and perimeter
            Dim area# = (depth * width - innerDepth * innerWidth)
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            Dim perimeter# = 2 * width + 2 * depth
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'width to thickness ratio and depth to thickness ratio
            'should be inner width and depth
            Dim width_thickness# = (width / thicknessDesign) - 2
            Dim depth_thickness# = (depth / thicknessDesign) - 2
            outputs.Add(IJUAHSSR_b_t, width_thickness)
            outputs.Add(IJUAHSSR_h_t, depth_thickness)

            'Calculation of moment of inertia about X and Y axis
            Dim momentOfInertia_X# = ((width * Math.Pow(depth, 3) / 12) - (innerWidth * Math.Pow(innerDepth, 3) / 12))
            Dim momentOfInertia_Y# = ((depth * Math.Pow(width, 3) / 12) - (innerDepth * Math.Pow(innerWidth, 3) / 12))
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Calculation of radius of gyration about X and Y axis
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X and Y axis
            Dim elasticSectionModulus_X# = momentOfInertia_X / (depth / 2)
            Dim elasticSectionModulus_Y# = momentOfInertia_Y / (width / 2)
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of plastic section modulus about X and Y axis
            Dim plasticSectionModulus_X# = ((width * depth * depth) - (innerDepth * innerDepth * innerDepth * innerWidth) / depth) / 6 'GetSectionModulusForHSSSection(depth, width, thicknessDesign) 'momentOfInertia_X / (6 * depth)
            Dim plasticSectionModulus_Y# = ((width * width * depth) - (innerDepth * innerWidth * innerWidth * innerWidth) / width) / 6 'GetSectionModulusForHSSSection(depth, width, thicknessDesign)
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia
            Dim h#, k#, v# 'local variables used for simplicity of calculation
            h = 2 * ((width - thicknessDesign) + (depth - thicknessDesign)) - 0.05 * depth * (4 - Math.PI)
            v = (width - thicknessDesign) * (depth - thicknessDesign) - 0.05 * 0.05 * depth * depth * (4 - Math.PI)
            k = 2 * v * thicknessDesign / h
            Dim torsionalMomentOfInertia# = 1 / 3 * Math.Pow(thicknessDesign, 3) * h + (2 * k * v)
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            CalculateHSSRSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate CS section properties like area, perimeter, unit weight,Moment of inertia about X-axis and Y-axis
        ''' Elastic and Plastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis
        ''' Properties like Polar radius of gyration about shear center, Flexural constant are left for edited by the user interactively, whose default values are set to zero. 
        ''' </summary>
        ''' <param name="inputProperties">Input properties like depth and width are used in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateCSSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following is the required input which is necessary in calculation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))

            'CS section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'In CS section width is always equal to depth
            Dim width# = depth
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Calculation of area and perimeter
            'Standard formula is used in calculation of the same(area and perimeter)
            Dim area# = Math.PI * depth * depth / 4
            Dim perimeter# = Math.PI * depth
            outputs.Add(IStructCrossSectionDimensions_Area, area)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Calculation of moment of inertia about X-axis
            Dim momentOfInertia_X# = (Math.PI * Math.Pow(depth, 4)) / 64
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of plastic section modulus X-axis
            Dim plasticSectionModulus_X# = Math.Pow(depth, 3) / 6
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)

            'Calculation of radius of gyration about X-axis
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)

            'Calculation of radius of gyration about Y-axis
            'This value will be equal to radius of gyration about X-axis because of symmetry
            Dim radiusOfGyration_Y# = radiusOfGyration_X
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X-axis
            Dim elasticSectionModulus_X# = (Math.PI * Math.Pow(depth, 3)) / 32
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)

            'The following properties are left for editing by the user whose default values are set to zero
            Dim polarRdiusOfGyration_ShearCenter# = 0, flexuralConstant# = 0
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRdiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)

            CalculateCSSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate RS section properties like area, perimeter, unit weight, Moment of inertia about X-axis and Y-axis
        ''' Elastic and Plastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis
        ''' Properties like Polar radius of gyration about shear center, Flexural constant are left for edited by the user interactively, whose default values are set to zero. 
        ''' </summary>
        ''' <param name="inputProperties">Input properties like depth and width  are used in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateRSSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))

            'RS section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Calculation of area and perimeter
            'Standard formula is used in calculation of the area and perimeter
            Dim area# = depth * width
            Dim perimeter# = 2 * (depth + width)
            outputs.Add(IStructCrossSectionDimensions_Area, area)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Calculation of moment of inertia about X-axis
            Dim momentOfInertia_X# = (width * Math.Pow(depth, 3)) / 12
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y-axis
            Dim momentOfInertia_Y# = (depth * Math.Pow(width, 3)) / 12
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Calculation of plastic section modulus about X-axis
            Dim plasticSectionModulus_X# = (width * Math.Pow(depth, 2)) / 4
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)

            'Calculation of plastic section modulus about Y-axis
            Dim plasticSectionModulus_Y# = (depth * Math.Pow(width, 2)) / 4
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of radius of gyration about X-axis
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)

            'Calculation of radius of gyration about Y-axis
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X-axis
            Dim elasticSectionModulus_X# = (width * Math.Pow(depth, 2)) / 6
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)

            'Calculation of elastic section modulus about Y-axis
            Dim elasticSectionModulus_Y# = (depth * Math.Pow(width, 2)) / 6
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'The following properties are left for edit by the user whose default values are set to zero
            Dim polarRdiusOfGyration_ShearCenter# = 0, flexuralConstant# = 0
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRdiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)

            CalculateRSSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUI section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia, Warping constant,
        ''' Horizontal and vertical distances from designated member edge to plastic neutral axis.
        ''' Properties like Polar radius of gyration about shear center, Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like depth, top flange' s width and thickness and web thickness are necessary in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUISectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))

            'BUI Section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Overall area is calculated by summing up individual areas of the section (web and flanges)
            Dim area# = 2 * topFlangeWidth * topFlangeThickness + webThickness * (depth - 2 * topFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Overall perimeter is calculated by summing up individual perimeters of the section (web and flanges)
            Dim perimeter# = 4 * topFlangeWidth + 4 * topFlangeThickness + 2 * (depth - 2 * topFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Width of the section is set to its top flange width 
            Dim width# = topFlangeWidth
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Centroid is calculated based upon section symmetry.
            'Since the section is symmetric about X and Y axis, calculation is simple
            Dim centroid_X# = width / 2
            Dim centroid_Y# = depth / 2
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of moment of inertia about X axis
            'Parallel axis theorem is used for calculation of overall moment of inertia
            Dim momentOfInertia_X# = width * Math.Pow(topFlangeThickness, 3) / 12 + width * topFlangeThickness * Math.Pow((topFlangeThickness / 2 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + webThickness * Math.Pow((depth - 2 * topFlangeThickness), 3) / 12 + (webThickness * (depth - 2 * topFlangeThickness) * Math.Pow((topFlangeThickness + 0.5 * (depth - 2 * topFlangeThickness)) - centroid_Y, 2))
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * Math.Pow(topFlangeThickness, 3) * width + (topFlangeThickness * width) * Math.Pow((topFlangeThickness + (depth - 2 * topFlangeThickness) + (topFlangeThickness / 2) - centroid_Y), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Similarly carry out the same for moment of inertia about Y
            Dim momentOfInertia_Y# = 2 * ((1 / 12) * topFlangeThickness * Math.Pow(width, 3) + (topFlangeThickness * width) * Math.Pow((width / 2 - centroid_X), 2))
            momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * Math.Pow(webThickness, 3) * (depth - 2 * topFlangeThickness) + (webThickness * (depth - 2 * topFlangeThickness)) * Math.Pow((width / 2 - centroid_X), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X and Y
            'The denominator in the formula checks the maximum value of (centroid_Y, (depth - centroid_Y))in case of X 
            'And (centroid_X, (flangeWidth - centroid_X)) in case of Y and calculates the elastic section modulus 
            Dim elasticSectionModulus_X# = momentOfInertia_X / Max(centroid_Y, (depth - centroid_Y))
            Dim elasticSectionModulus_Y# = momentOfInertia_Y / Max(centroid_X, (width - centroid_X))
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of horizontal and vertical distances from designated member edge to plastic neutral axis
            Dim horizontalDistance_PlasticNeutralAxis# = (width * topFlangeThickness + (depth - topFlangeThickness) * webThickness + topFlangeThickness * (webThickness - width)) / (2 * webThickness)
            Dim verticalDistance_PlasticNeutralAxis# = (topFlangeThickness * width + (depth - 2 * topFlangeThickness) * width + topFlangeThickness * width) / (2 * depth)
            outputs.Add(IJUAL_xp, horizontalDistance_PlasticNeutralAxis)
            outputs.Add(IJUAL_yp, verticalDistance_PlasticNeutralAxis)

            'Calculation of plastic section modulus about X and Y
            Dim plasticSectionModulus_X# = area / 2 * (((webThickness * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) / 2) + (topFlangeThickness * width * (depth - horizontalDistance_PlasticNeutralAxis - topFlangeThickness / 2))) / (webThickness * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) + _
                  topFlangeThickness * width) + (webThickness * (horizontalDistance_PlasticNeutralAxis - topFlangeThickness) * (horizontalDistance_PlasticNeutralAxis - topFlangeThickness) / 2 + topFlangeThickness * width * (horizontalDistance_PlasticNeutralAxis - topFlangeThickness / 2)) / (webThickness * (horizontalDistance_PlasticNeutralAxis - topFlangeThickness) + width * topFlangeThickness))

            Dim plasticSectionModulus_Y# = area / 2 * ((topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - width) / 2) * (verticalDistance_PlasticNeutralAxis - (width - width) / 2) / 2 + (depth - topFlangeThickness - topFlangeThickness) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) / 2 + _
                topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - width) / 2) * (verticalDistance_PlasticNeutralAxis - (width - width) / 2) / 2) / (topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - width) / 2) + (depth - topFlangeThickness - topFlangeThickness) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) + topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - width) / 2)) + _
                (topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - width) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - width) / 2) / 2 + (depth - topFlangeThickness - topFlangeThickness) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) / 2 + _
                topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - width) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - width) / 2) / 2) / (topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - width) / 2) + (depth - topFlangeThickness - topFlangeThickness) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) + topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - width) / 2)))

            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia and warping constant
            Dim torsionalMomentOfInertia# = (width * topFlangeThickness * topFlangeThickness * topFlangeThickness + width * topFlangeThickness * topFlangeThickness * topFlangeThickness + webThickness * webThickness * webThickness * (depth - topFlangeThickness - topFlangeThickness)) / 3
            Dim warpingConstant# = (topFlangeThickness * (depth - 2 * topFlangeThickness) * (depth - 2 * topFlangeThickness) * (width * width * width * width * width * width / (Math.Pow(width, 3) + Math.Pow(width, 3))) / 12)
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)

            'The following properties are left for edited by the user whose default values are set to zero/One
            Dim warpingStaticalMoment# = 0, polarRadiusOfGyration_ShearCenter# = 0, flexuralConstant# = 0, radiusOfGyration_XY# = 0
            Dim flangeGage# = 0, webGage# = 0, lengthExtension# = 0, depthExtension# = 0, topFlangeWidthExtension# = 0
            Dim isModifiable As Long = 0, systemProperties As Long = 1

            outputs.Add(IStructFlangedBoltGage_gf, flangeGage)
            outputs.Add(IStructFlangedBoltGage_gw, webGage)
            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalMoment)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)
            outputs.Add(IUABuiltUpWeb_DepthExt, depthExtension)
            outputs.Add(IUABuiltUpTopFlange_TopFlangeWidthExt, topFlangeWidthExtension)

            CalculateBUISectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUI Tap web section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia, Warping constant,
        ''' Horizontal and vertical distances from designated member edge to plastic neutral axis.
        ''' Properties like Polar radius of gyration about shear center, Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like Top and bottom flange' s width and thickness, web thickness, length start and end are used in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUITapWebSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs for calculation of properties
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim depthStart# = CDbl(inputProperties.Item(IUABuiltUpITaperWeb_DepthStart))
            Dim depthEnd# = CDbl(inputProperties.Item(IUABuiltUpITaperWeb_DepthEnd))

            'The value of depth will be the maximum value of depth start or depth end 
            Dim depth# = Max(depthStart, depthEnd)
            outputs.Add(IStructCrossSectionDimensions_Depth, depth)

            'The value of width will be the maximum value of top flange width or bottom flange width 
            Dim width# = Max(topFlangeWidth, bottomFlangeWidth)
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'BUITapWeb is not a hollow section
            Dim isHollow As Boolean

            'Checking for symmetry about X
            Dim isSymmetryAboutX As Boolean
            If topFlangeWidth = bottomFlangeWidth And topFlangeThickness = bottomFlangeThickness Then
                isSymmetryAboutX = True
            End If

            'BIUTapWeb is always symmetric about Y
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Calculation of area
            'Maximum values have been taken into consideration
            Dim area# = (topFlangeThickness * topFlangeWidth) + (bottomFlangeThickness * bottomFlangeWidth) + (depth - topFlangeThickness - bottomFlangeThickness) * webThickness
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Calculation of perimeter
            Dim perimeter# = 2 * (topFlangeWidth + bottomFlangeWidth + topFlangeThickness + bottomFlangeThickness) + 2 * (depth - topFlangeThickness - bottomFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Calculation of centroid about X
            Dim centroid_X# = width / 2
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)

            'Calculation of centroid about Y
            'Calculate the each entity area of the cross section

            Dim A1# = (bottomFlangeThickness * bottomFlangeWidth)
            Dim A2# = webThickness * (depth - topFlangeThickness - bottomFlangeThickness)
            Dim A3# = (topFlangeWidth * topFlangeThickness)

            'Calculate the each entity centroid about Y
            Dim y1# = bottomFlangeThickness / 2
            Dim y2# = bottomFlangeThickness + (depth - topFlangeThickness - bottomFlangeThickness) / 2
            Dim y3# = topFlangeThickness + (depth - topFlangeThickness - bottomFlangeThickness) + topFlangeThickness / 2

            'Overall centroid about Y is the summation of product of (entity area and entity centroid) divided by total area
            Dim centroid_Y# = ((A1 * y1) + (A2 * y2) + (A3 * y3)) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of moment of inertia about X axis
            'Parallel axis theorem is used for the calculation
            Dim momentOfInertia_X# = (1 / 12) * bottomFlangeWidth * Math.Pow(bottomFlangeThickness, 3) + A1 * Math.Pow((y1 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * webThickness * Math.Pow((depth - topFlangeThickness - bottomFlangeThickness), 3) + A2 * Math.Pow((y2 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * topFlangeWidth * Math.Pow(topFlangeThickness, 3) + A3 * Math.Pow((y3 - centroid_Y), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y axis
            'Parallel axis theorem is used for the calculation
            Dim momentOfInertia_Y# = (1 / 12) * bottomFlangeThickness * Math.Pow(bottomFlangeWidth, 3) + A1 * Math.Pow((width / 2 - centroid_X), 2)
            momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * (depth - topFlangeThickness - bottomFlangeThickness) * Math.Pow(webThickness, 3) + A2 * Math.Pow((width / 2 - centroid_X), 2)
            momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * topFlangeThickness * Math.Pow(topFlangeWidth, 3) + A3 * Math.Pow((width / 2 - centroid_X), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Calculation of radius of gyration about X
            Dim radiusOfGyration_X = Math.Sqrt(momentOfInertia_X / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)

            'Calculation of radius of gyration about Y
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X
            'The denominator in the formula checks the maximum value of (centroid_Y, (depth - centroid_Y))in case of X 
            'And (centroid_X, (flangeWidth - centroid_X)) in case of Y and calculates the elastic section modulus

            Dim elasticSectionModulus_X# = momentOfInertia_X / Max(centroid_Y, (depth - centroid_Y))
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)

            'Calculation of elastic section modulus about Y
            Dim elasticSectionModulus_Y# = momentOfInertia_Y / Max(centroid_X, (width - centroid_X))
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of horizontal and vertical distances from designated member edge to plastic neutral axis
            'These are required in the calculations of plastic section modulus about X and Y
            'These are not added to the output dictionary
            Dim horizontalDistance_PlasticNeutralAxis# = (topFlangeWidth * topFlangeThickness + (depth - topFlangeThickness) * webThickness + bottomFlangeThickness * (webThickness - bottomFlangeWidth)) / (2 * webThickness)
            Dim verticalDistance_PlasticNeutralAxis# = (topFlangeThickness * width + (depth - topFlangeThickness - bottomFlangeThickness) * width + bottomFlangeThickness * width) / (2 * depth)

            'Calculation of plastic section modulus about X
            Dim plasticSectionModulus_X# = area / 2 * (((webThickness * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) / 2) + (topFlangeThickness * topFlangeWidth * (depth - horizontalDistance_PlasticNeutralAxis - topFlangeThickness / 2))) / (webThickness * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) + _
                  topFlangeThickness * topFlangeWidth) + (webThickness * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) / 2 + bottomFlangeThickness * bottomFlangeWidth * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness / 2)) / (webThickness * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) + bottomFlangeWidth * bottomFlangeThickness))
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)

            'Calculation of plastic section modulus about Y
            Dim plasticSectionModulus_Y# = area / 2 * ((topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) / 2 + (depth - topFlangeThickness - bottomFlangeThickness) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) / 2 + _
                bottomFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) / 2) / (topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) + (depth - topFlangeThickness - bottomFlangeThickness) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) + bottomFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2)) + _
                (topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) / 2 + (depth - topFlangeThickness - bottomFlangeThickness) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) / 2 + _
                bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) / 2) / (topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) + (depth - topFlangeThickness - bottomFlangeThickness) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) + bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2)))
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia
            Dim torsionalMomentOfInertia# = (topFlangeWidth * Math.Pow(topFlangeThickness, 3) + bottomFlangeWidth * Math.Pow(bottomFlangeThickness, 3) + Math.Pow(webThickness, 3) * (depth - topFlangeThickness - bottomFlangeThickness)) / 3
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            'Calculation of warping constant
            Dim warpingConstant# = (topFlangeThickness * (depth - 2 * topFlangeThickness) * (depth - 2 * topFlangeThickness) * (topFlangeWidth * topFlangeWidth * topFlangeWidth * bottomFlangeWidth * bottomFlangeWidth * bottomFlangeWidth / (Math.Pow(topFlangeWidth, 3) + Math.Pow(bottomFlangeWidth, 3))) / 12)
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)

            'The following properties are left edited by the user, whose default values are set as zero/one as of now

            Dim polarRadiusOfGyration_ShearCenter = 0, radiusOfGyration_XY = 0, warpingStaticalMoment = 0
            Dim flexuralConstant# = 0, topFlangeWidthExtension# = 0, bottomFlangeWidthExtension# = 0
            Dim lengthExtension# = 0, depthExtension# = 0, isModifiable As Long = 0, systemProperties As Long = 1

            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalMoment)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)
            outputs.Add(IUABuiltUpWeb_DepthExt, depthExtension)
            outputs.Add(IUABuiltUpTopFlange_TopFlangeWidthExt, topFlangeWidthExtension)
            outputs.Add(IUABuiltUpBottomFlange_BottomFlangeWidthExt, bottomFlangeWidthExtension)

            CalculateBUITapWebSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUIUE section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia, Warping constant,
        ''' Horizontal and vertical distances from designated member edge to plastic neutral axis.
        ''' Polar radius of gyration about shear center, Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like depth, web thickness, top and bottom flange' s width and thickness are necessary in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUIUESectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))

            'BUIUE Section is not considered as hollow section
            Dim isHollow As Boolean

            'Verifying for symmetry about X axis
            Dim isSymmetryAboutX As Boolean

            If (topFlangeThickness = bottomFlangeThickness) And (topFlangeWidth = bottomFlangeWidth) Then
                isSymmetryAboutX = True
            End If

            'BUIUE section is always symmetric about Y axis
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Overall area is calculated  by summing up individual areas of the section (top flange, bottom flange, web)
            Dim area# = topFlangeWidth * topFlangeThickness + bottomFlangeWidth * bottomFlangeThickness + webThickness * (depth - topFlangeThickness - bottomFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Overall perimeter is calculated  by summing up individual perimeters of the section (top flange, bottom flange, web)
            Dim perimeter# = 2 * (topFlangeWidth + bottomFlangeWidth) + 2 * (topFlangeThickness + bottomFlangeThickness) + 2 * (depth - topFlangeThickness - bottomFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Value of width is the maximum of top flange width or bottom flange width
            Dim width# = Max(topFlangeWidth, bottomFlangeWidth)
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Calculation of centroid about X
            'Since width is the maximum of Top flange width or Bottom flange width

            Dim centroid_X#
            centroid_X = width / 2
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)

            'Calculation of centroid about Y
            'Calculate the each entity area of the cross section

            Dim A1# = (bottomFlangeThickness * bottomFlangeWidth)
            Dim A2# = webThickness * (depth - topFlangeThickness - bottomFlangeThickness)
            Dim A3# = (topFlangeWidth * topFlangeThickness)

            'Calculate the each entity centroid about Y
            Dim y1# = bottomFlangeThickness / 2
            Dim y2# = bottomFlangeThickness + (depth - topFlangeThickness - bottomFlangeThickness) / 2
            Dim y3# = topFlangeThickness + (depth - topFlangeThickness - bottomFlangeThickness) + topFlangeThickness / 2

            'Overall centroid about Y is the summation of product of (entity area and entity centroid) divided by total area
            Dim centroid_Y# = ((A1 * y1) + (A2 * y2) + (A3 * y3)) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of moment of inertia about X axis
            'Parallel axis theorem is used for the calculation
            Dim momentOfInertia_X# = (1 / 12) * bottomFlangeWidth * Math.Pow(bottomFlangeThickness, 3) + A1 * Math.Pow((y1 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * webThickness * Math.Pow((depth - topFlangeThickness - bottomFlangeThickness), 3) + A2 * Math.Pow((y2 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * topFlangeWidth * Math.Pow(topFlangeThickness, 3) + A3 * Math.Pow((y3 - centroid_Y), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y axis
            'Parallel axis theorem is used for the calculation
            Dim momentOfInertia_Y# = (1 / 12) * bottomFlangeThickness * Math.Pow(bottomFlangeWidth, 3) + A1 * Math.Pow((width / 2 - centroid_X), 2)
            momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * (depth - topFlangeThickness - bottomFlangeThickness) * Math.Pow(webThickness, 3) + A2 * Math.Pow((width / 2 - centroid_X), 2)
            momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * topFlangeThickness * Math.Pow(topFlangeWidth, 3) + A3 * Math.Pow((width / 2 - centroid_X), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X and Y
            'The denominator in the formula checks the maximum value of (centroid_Y, (depth - centroid_Y))in case of X 
            'And (centroid_X, (flangeWidth - centroid_X)) in case of Y and calculates the elastic section modulus 
            Dim elasticSectionModulus_X# = momentOfInertia_X / Max(centroid_Y, (depth - centroid_Y))
            Dim elasticSectionModulus_Y# = momentOfInertia_Y / Max(centroid_X, (width - centroid_X))
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of horizontal and vertical distances from designated member edge to plastic neutral axis
            Dim horizontalDistance_PlasticNeutralAxis# = (topFlangeWidth * topFlangeThickness + (depth - topFlangeThickness) * webThickness + bottomFlangeThickness * (webThickness - bottomFlangeWidth)) / (2 * webThickness)
            Dim verticalDistance_PlasticNeutralAxis# = (topFlangeThickness * width + (depth - topFlangeThickness - bottomFlangeThickness) * width + bottomFlangeThickness * width) / (2 * depth)
            outputs.Add(IJUAL_xp, horizontalDistance_PlasticNeutralAxis)
            outputs.Add(IJUAL_yp, verticalDistance_PlasticNeutralAxis)

            'Calculation of plastic section modulus about X and Y
            Dim plasticSectionModulus_X# = area / 2 * (((webThickness * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) / 2) + (topFlangeThickness * topFlangeWidth * (depth - horizontalDistance_PlasticNeutralAxis - topFlangeThickness / 2))) / (webThickness * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) + _
                  topFlangeThickness * topFlangeWidth) + (webThickness * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) / 2 + bottomFlangeThickness * bottomFlangeWidth * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness / 2)) / (webThickness * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) + bottomFlangeWidth * bottomFlangeThickness))

            Dim plasticSectionModulus_Y# = area / 2 * ((topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) / 2 + (depth - topFlangeThickness - bottomFlangeThickness) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) / 2 + _
                bottomFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) / 2) / (topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) + (depth - topFlangeThickness - bottomFlangeThickness) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) + bottomFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2)) + _
                (topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) / 2 + (depth - topFlangeThickness - bottomFlangeThickness) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) / 2 + _
                bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) / 2) / (topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) + (depth - topFlangeThickness - bottomFlangeThickness) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) + bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2)))

            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia and warping constant
            Dim torsionalMomentOfInertia# = (topFlangeWidth * topFlangeThickness * topFlangeThickness * topFlangeThickness + bottomFlangeWidth * bottomFlangeThickness * bottomFlangeThickness * bottomFlangeThickness + webThickness * webThickness * webThickness * (depth - topFlangeThickness - bottomFlangeThickness)) / 3
            Dim warpingConstant# = (topFlangeThickness * (depth - 2 * topFlangeThickness) * (depth - 2 * topFlangeThickness) * (topFlangeWidth * topFlangeWidth * topFlangeWidth * bottomFlangeWidth * bottomFlangeWidth * bottomFlangeWidth / (Math.Pow(topFlangeWidth, 3) + Math.Pow(bottomFlangeWidth, 3))) / 12)
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)

            'The following properties are left for edited by the user whose default values are set to zero/One
            Dim warpingStaticalConstant# = 0, polarRadiusOfGyration_ShearCenter# = 0, flexuralConstant# = 0, topFlangeWidthExtension# = 0, bottomFlangeWidthExtension# = 0
            Dim radiusOfGyration_XY# = 0, flangeGage# = 0, webGage# = 0, lengthExtension# = 0, depthExtension# = 0
            Dim isModifiable As Long = 0, systemProperties As Long = 1

            outputs.Add(IStructFlangedBoltGage_gf, flangeGage)
            outputs.Add(IStructFlangedBoltGage_gw, webGage)
            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalConstant)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)
            outputs.Add(IUABuiltUpWeb_DepthExt, depthExtension)
            outputs.Add(IUABuiltUpTopFlange_TopFlangeWidthExt, topFlangeWidthExtension)
            outputs.Add(IUABuiltUpBottomFlange_BottomFlangeWidthExt, bottomFlangeWidthExtension)

            CalculateBUIUESectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUIHaunch section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia, Warping constant,
        ''' Horizontal and vertical distances from designated member edge to plastic neutral axis.
        ''' Properties like Polar radius of gyration about shear center, Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like Top and bottom flange' s width and thickness, web thickness, length start and end, depth start and end, transition gradient start and end are helpful in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUIHaunchSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim depthStart# = CDbl(inputProperties.Item(IUABuiltUpIHaunch_DepthStart))
            Dim depthEnd# = CDbl(inputProperties.Item(IUABuiltUpIHaunch_DepthEnd))

            'The value of depth can be the maximum value of depth start or depth end 
            Dim depth# = Max(depthEnd, depthStart)
            outputs.Add(IStructCrossSectionDimensions_Depth, depth)

            'The value of width can be the maximum value of top flange width or bottom flange width
            Dim width# = Max(topFlangeWidth, bottomFlangeWidth)
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'BUIHaunch is not a hollow section
            Dim isHollow As Boolean

            'BUIHaunch is symmetric about both X and Y axis according to its geometry
            Dim isSymmetryAboutX As Boolean

            If ((topFlangeWidth = bottomFlangeWidth) And (topFlangeThickness = bottomFlangeThickness)) Then
                isSymmetryAboutX = True
            End If

            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Overall area is calculated  by summing up individual areas of the section (top flange, bottom flange, web)
            Dim area# = topFlangeWidth * topFlangeThickness + bottomFlangeWidth * bottomFlangeThickness + webThickness * (depth - topFlangeThickness - bottomFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Overall perimeter is calculated  by summing up individual perimeters of the section (top flange, bottom flange, web)
            Dim perimeter# = 2 * (topFlangeWidth + bottomFlangeWidth) + 2 * (topFlangeThickness + bottomFlangeThickness) + 2 * (depth - topFlangeThickness - bottomFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Calculation of centroid about X
            'Since width is the maximum of Top flange width or Bottom flange width

            Dim centroid_X# = width / 2
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)

            'Calculation of centroid about Y
            'Calculate the each entity area of the cross section

            Dim A1# = (bottomFlangeThickness * bottomFlangeWidth)
            Dim A2# = webThickness * (depth - topFlangeThickness - bottomFlangeThickness)
            Dim A3# = (topFlangeWidth * topFlangeThickness)

            'Calculate the each entity centroid about Y
            Dim y1# = bottomFlangeThickness / 2
            Dim y2# = bottomFlangeThickness + (depth - topFlangeThickness - bottomFlangeThickness) / 2
            Dim y3# = topFlangeThickness + (depth - topFlangeThickness - bottomFlangeThickness) + topFlangeThickness / 2

            'Overall centroid about Y is the summation of product of (entity area and entity centroid) divided by total area
            Dim centroid_Y# = ((A1 * y1) + (A2 * y2) + (A3 * y3)) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of moment of inertia about X axis
            'Parallel axis theorem is used for the calculation
            Dim momentOfInertia_X# = (1 / 12) * bottomFlangeWidth * Math.Pow(bottomFlangeThickness, 3) + A1 * Math.Pow((y1 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * webThickness * Math.Pow((depth - topFlangeThickness - bottomFlangeThickness), 3) + A2 * Math.Pow((y2 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * topFlangeWidth * Math.Pow(topFlangeThickness, 3) + A3 * Math.Pow((y3 - centroid_Y), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y axis
            'Parallel axis theorem is used for the calculation
            Dim momentOfInertia_Y# = (1 / 12) * bottomFlangeThickness * Math.Pow(bottomFlangeWidth, 3) + A1 * Math.Pow((width / 2 - centroid_X), 2)
            momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * (depth - topFlangeThickness - bottomFlangeThickness) * Math.Pow(webThickness, 3) + A2 * Math.Pow((width / 2 - centroid_X), 2)
            momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * topFlangeThickness * Math.Pow(topFlangeWidth, 3) + A3 * Math.Pow((width / 2 - centroid_X), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X and Y
            'The denominator in the formula checks the maximum value of (centroid_Y, (depth - centroid_Y))in case of X 
            'And (centroid_X, (flangeWidth - centroid_X)) in case of Y and calculates the elastic section modulus 
            Dim elasticSectionModulus_X# = momentOfInertia_X / Max(centroid_Y, (depth - centroid_Y))
            Dim elasticSectionModulus_Y# = momentOfInertia_Y / Max(centroid_X, (width - centroid_X))
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of horizontal and vertical distances from designated member edge to plastic neutral axis
            Dim horizontalDistance_PlasticNeutralAxis# = (topFlangeWidth * topFlangeThickness + (depth - topFlangeThickness) * webThickness + bottomFlangeThickness * (webThickness - bottomFlangeWidth)) / (2 * webThickness)
            Dim verticalDistance_PlasticNeutralAxis# = (topFlangeThickness * width + (depth - topFlangeThickness - bottomFlangeThickness) * width + bottomFlangeThickness * width) / (2 * depth)
            outputs.Add(IJUAL_xp, horizontalDistance_PlasticNeutralAxis)
            outputs.Add(IJUAL_yp, verticalDistance_PlasticNeutralAxis)

            'Calculation of plastic section modulus about X and Y
            Dim plasticSectionModulus_X# = area / 2 * (((webThickness * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) / 2) + (topFlangeThickness * topFlangeWidth * (depth - horizontalDistance_PlasticNeutralAxis - topFlangeThickness / 2))) / (webThickness * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) + _
                  topFlangeThickness * topFlangeWidth) + (webThickness * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) / 2 + bottomFlangeThickness * bottomFlangeWidth * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness / 2)) / (webThickness * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) + bottomFlangeWidth * bottomFlangeThickness))

            Dim plasticSectionModulus_Y# = area / 2 * ((topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) / 2 + (depth - topFlangeThickness - bottomFlangeThickness) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) / 2 + _
                bottomFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) / 2) / (topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) + (depth - topFlangeThickness - bottomFlangeThickness) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) + bottomFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2)) + _
                (topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) / 2 + (depth - topFlangeThickness - bottomFlangeThickness) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) / 2 + _
                bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2) / 2) / (topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth) / 2) + (depth - topFlangeThickness - bottomFlangeThickness) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) + bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth) / 2)))

            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia and warping constant
            Dim torsionalMomentOfInertia# = (topFlangeWidth * topFlangeThickness * topFlangeThickness * topFlangeThickness + bottomFlangeWidth * bottomFlangeThickness * bottomFlangeThickness * bottomFlangeThickness + webThickness * webThickness * webThickness * (depth - topFlangeThickness - bottomFlangeThickness)) / 3
            Dim warpingConstant# = (topFlangeThickness * (depth - 2 * topFlangeThickness) * (depth - 2 * topFlangeThickness) * (topFlangeWidth * topFlangeWidth * topFlangeWidth * bottomFlangeWidth * bottomFlangeWidth * bottomFlangeWidth / (Math.Pow(topFlangeWidth, 3) + Math.Pow(bottomFlangeWidth, 3))) / 12)
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)

            'The following properties are left edited by the user, whose default values are set as zero as of now
            Dim isModifiable As Long = 0, systemProperties As Long = 1, flexuralConstant# = 0, polarRadiusOfGyration_ShearCenter# = 0, radiusOfGyration_XY# = 0
            Dim topFlangeWidthExtension# = 0, bottomFlangeWidthExtension# = 0, lengthExtension# = 0, depthExtension# = 0, warpingStaticalMoment# = 0

            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalMoment)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)
            outputs.Add(IUABuiltUpWeb_DepthExt, depthExtension)
            outputs.Add(IUABuiltUpTopFlange_TopFlangeWidthExt, topFlangeWidthExtension)
            outputs.Add(IUABuiltUpBottomFlange_BottomFlangeWidthExt, bottomFlangeWidthExtension)

            CalculateBUIHaunchSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUL section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia, Warping constant,
        ''' Horizontal and vertical distances from designated member edge to plastic neutral axis.
        ''' Properties like Polar radius of gyration about shear center, Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like depth, web offset, web thickness, bottom flange' s width and thickness are necessary in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBULSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim offSetWeb# = CDbl(inputProperties.Item(IUABuiltUpL_OffsetWeb))

            'BUL Section is not considered as hollow section
            Dim isHollow As Boolean

            'According to the symmetry, it is neither symmetric about X nor Y
            Dim isSymmetryAboutX As Boolean
            Dim isSymmetryAboutY As Boolean

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'width is set to bottom flange width
            Dim width# = bottomFlangeWidth
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Overall area is calculated by summing up individual areas of the section (bottom flange, web)
            Dim area# = width * bottomFlangeThickness + webThickness * (depth - bottomFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Overall perimeter is calculated by summing up individual perimeters of the section (bottom flange, web)
            Dim perimeter# = width + 2 * bottomFlangeThickness + 2 * (depth - bottomFlangeThickness) + width
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Centroid_X calculation
            'For simplicity, calculate the individual entities area

            Dim a1# = bottomFlangeThickness * width
            Dim a2# = (depth - bottomFlangeThickness) * webThickness

            'Calculate the entities centroid about X
            Dim x1# = width / 2
            Dim x2# = offSetWeb + (webThickness / 2)

            'Overall centroid about X
            Dim centroid_X# = ((a1 * x1) + (a2 * x2)) / (a1 + a2)
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)

            'Centroid_Y calculation
            'Calculate the entities centroid about Y

            Dim y1# = bottomFlangeThickness / 2
            Dim y2# = bottomFlangeThickness + (depth - bottomFlangeThickness) / 2

            'Overall centroid about Y
            Dim centroid_Y# = ((a1 * y1) + (a2 * y2)) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of horizontal distance from designated member edge to plastic neutral axis
            'Before calculation, it is checked for a condition using input properties
            Dim horizontalDistance_PlasticNeutralAxis#
            If width * bottomFlangeThickness < (depth - bottomFlangeThickness) * webThickness Then
                horizontalDistance_PlasticNeutralAxis = (webThickness * depth + bottomFlangeThickness * webThickness - width * bottomFlangeThickness) / (2 * webThickness)
            Else
                horizontalDistance_PlasticNeutralAxis = (bottomFlangeThickness * width + webThickness * (depth - bottomFlangeThickness)) / (2 * width)
            End If
            outputs.Add(IJUAL_xp, horizontalDistance_PlasticNeutralAxis)

            'Calculation of vertical distance from designated member edge to plastic neutral axis
            'Before calculation, it is checked for a condition using input properties
            Dim verticalDistance_PlasticNeutralAxis#
            If (offSetWeb + webThickness) * bottomFlangeThickness + (depth - bottomFlangeThickness) * webThickness > (width - offSetWeb - webThickness) * bottomFlangeThickness Then
                verticalDistance_PlasticNeutralAxis = ((depth - bottomFlangeThickness) * offSetWeb + (depth - bottomFlangeThickness) * (offSetWeb + webThickness) + bottomFlangeThickness * width) / (2 * depth * bottomFlangeThickness)
            Else
                verticalDistance_PlasticNeutralAxis = (width * bottomFlangeThickness - (depth - bottomFlangeThickness) * webThickness) / (2 * bottomFlangeThickness)
            End If
            outputs.Add(IJUAL_yp, verticalDistance_PlasticNeutralAxis)

            'Calculation of moment of inertia about X
            Dim momentOfInertia_X# = (webThickness * Math.Pow(depth, 3) + (width - webThickness) * Math.Pow(bottomFlangeThickness, 3)) / 12
            momentOfInertia_X = momentOfInertia_X + depth * webThickness * Math.Pow(((0.5 * depth) - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (bottomFlangeThickness * (width - webThickness) * Math.Pow((centroid_Y - (0.5 * bottomFlangeThickness)), 2))
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y
            Dim momentOfInertia_Y# = (bottomFlangeThickness * Math.Pow(width, 3) / 12) + (width * bottomFlangeThickness * Math.Pow(((0.5 * width) - centroid_X), 2))
            momentOfInertia_Y = momentOfInertia_Y + (depth - bottomFlangeThickness) * Math.Pow(webThickness, 3.0) / 12.0
            momentOfInertia_Y = momentOfInertia_Y + (depth - bottomFlangeThickness) * webThickness * Math.Pow((centroid_X - 0.5 * webThickness), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X and Y
            'The denominator in the formula checks the maximum value of (centroid_Y, (depth - centroid_Y))in case of X 
            'And (centroid_X, (width - centroid_X)) in case of Y and calculates the elastic section modulus 

            Dim elasticSectionModulus_X# = momentOfInertia_X / Max(centroid_Y, (depth - centroid_Y))
            Dim elasticSectionModulus_Y# = momentOfInertia_Y / Max(centroid_X, (width - centroid_X))
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of plastic section modulus about X
            'Before calculation, it is checked for a condition using input properties
            Dim plasticSectionModulus_X#
            If width * bottomFlangeThickness < (depth - bottomFlangeThickness) * webThickness Then
                plasticSectionModulus_X = area / 2 * ((webThickness * (depth - horizontalDistance_PlasticNeutralAxis) * (depth - horizontalDistance_PlasticNeutralAxis) / 2) / (webThickness * (depth - horizontalDistance_PlasticNeutralAxis)) + (webThickness * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) / 2 + width * bottomFlangeThickness * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness / 2)) / (webThickness * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) + width * bottomFlangeThickness))
            Else
                plasticSectionModulus_X = area / 2 * ((width * (bottomFlangeThickness - horizontalDistance_PlasticNeutralAxis) * (bottomFlangeThickness - horizontalDistance_PlasticNeutralAxis) / 2 + webThickness * (depth - bottomFlangeThickness) * (bottomFlangeThickness - 2 * horizontalDistance_PlasticNeutralAxis + depth) / 2) / (width * (bottomFlangeThickness - horizontalDistance_PlasticNeutralAxis) + webThickness * (depth - bottomFlangeThickness)) + (width * horizontalDistance_PlasticNeutralAxis * horizontalDistance_PlasticNeutralAxis / 2) / (width * horizontalDistance_PlasticNeutralAxis))
            End If
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)

            'Calculation of plastic section modulus about Y
            'Before calculation, it is checked for a condition using input properties
            Dim plasticSectionModulus_Y#
            If (offSetWeb + webThickness) * bottomFlangeThickness + (depth - bottomFlangeThickness) * webThickness > (width - offSetWeb - webThickness) * bottomFlangeThickness Then
                plasticSectionModulus_Y = area / 2 * ((bottomFlangeThickness * verticalDistance_PlasticNeutralAxis * verticalDistance_PlasticNeutralAxis / 2 + (depth - bottomFlangeThickness) * (verticalDistance_PlasticNeutralAxis - offSetWeb) * (verticalDistance_PlasticNeutralAxis - offSetWeb) / 2) / (bottomFlangeThickness * verticalDistance_PlasticNeutralAxis + (depth - bottomFlangeThickness) * (verticalDistance_PlasticNeutralAxis - offSetWeb)) + ((depth - bottomFlangeThickness) * (webThickness - verticalDistance_PlasticNeutralAxis + offSetWeb) * (webThickness - verticalDistance_PlasticNeutralAxis + offSetWeb) / 2 + _
                    bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis) * (width - verticalDistance_PlasticNeutralAxis) / 2) / ((depth - bottomFlangeThickness) * (webThickness - verticalDistance_PlasticNeutralAxis + offSetWeb) + bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis)))
            Else
                plasticSectionModulus_Y = area / 2 * (((depth - bottomFlangeThickness) * webThickness * (verticalDistance_PlasticNeutralAxis - offSetWeb - webThickness / 2) + bottomFlangeThickness * verticalDistance_PlasticNeutralAxis * verticalDistance_PlasticNeutralAxis / 2) / ((depth - bottomFlangeThickness) * webThickness + bottomFlangeThickness * verticalDistance_PlasticNeutralAxis) + (bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis) * (width - verticalDistance_PlasticNeutralAxis) / 2) / (bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis)))
            End If
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia
            Dim torsionalMomentOfInertia# = (width * bottomFlangeThickness * bottomFlangeThickness * bottomFlangeThickness + webThickness * webThickness * webThickness * (depth - bottomFlangeThickness)) / 3
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            'Calculation of warping constant
            Dim warpingConstant# = (width * width * width * bottomFlangeThickness * bottomFlangeThickness * bottomFlangeThickness + (depth - bottomFlangeThickness) * (depth - bottomFlangeThickness) * (depth - bottomFlangeThickness) * webThickness * webThickness * webThickness) / 36
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)

            'The following properties are left for edited by the user interactively , whose default values are set to zero/One as of now
            Dim polarRadiusOfGyration_ShearCenter# = 0, radiusOfGyration_XY# = 0, warpingStaticalMoment# = 0, flexuralConstant# = 0, lengthExtension# = 0, depthExtension# = 0
            Dim longSideGage# = 0, longSideGage1# = 0, longSideGage2# = 0, shortSideGage# = 0, shortSideGage1# = 0, shortSideGage2# = 0, bottomFlangeWidthExtension# = 0
            Dim isModifiable As Long = 0, systemProperties As Long = 1

            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalMoment)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IStructAngleBoltGage_lsg, longSideGage)
            outputs.Add(IStructAngleBoltGage_lsg1, longSideGage1)
            outputs.Add(IStructAngleBoltGage_lsg2, longSideGage2)
            outputs.Add(IStructAngleBoltGage_ssg, shortSideGage)
            outputs.Add(IStructAngleBoltGage_ssg1, shortSideGage1)
            outputs.Add(IStructAngleBoltGage_ssg2, shortSideGage2)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)
            outputs.Add(IUABuiltUpWeb_DepthExt, depthExtension)
            outputs.Add(IUABuiltUpBottomFlange_BottomFlangeWidthExt, bottomFlangeWidthExtension)

            CalculateBULSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUC section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia,
        ''' Horizontal and vertical distances from designated member edge to plastic neutral axis.
        ''' Properties like Warping constant, Polar radius of gyration about shear center, Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like depth, web offset - top and bottom , web thickness, top and bottom flange' s width and thickness are necessary in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUCSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim offSetWebTop# = CDbl(inputProperties.Item(IUABuiltUpC_OffsetWebTop))
            Dim offSetWebBottom# = CDbl(inputProperties.Item(IUABuiltUpC_OffsetWebBot))

            'BUC Section is not considered as hollow section
            Dim isHollow As Boolean

            'Verifying for section symmetry about X axis
            Dim isSymmetryAboutX As Boolean
            If (offSetWebBottom = offSetWebTop) And (topFlangeWidth = bottomFlangeWidth) And (topFlangeThickness = bottomFlangeThickness) Then
                isSymmetryAboutX = True
            End If

            'BUC Section is always unsymmetrical about Y axis
            Dim isSymmetryAboutY As Boolean

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Overall area is calculated by summing up individual areas of the section (top flange, bottom flange, web)
            Dim area# = topFlangeWidth * topFlangeThickness + bottomFlangeWidth * bottomFlangeThickness + webThickness * (depth - topFlangeThickness - bottomFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Overall perimeter is calculated by summing up individual perimeters of the section (top flange, bottom flange, web)
            Dim perimeter# = topFlangeWidth + 2 * topFlangeThickness + bottomFlangeWidth + 2 * bottomFlangeThickness + 2 * (depth - topFlangeThickness - bottomFlangeThickness) + (topFlangeWidth - webThickness) + (bottomFlangeWidth - webThickness)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Value of width is the maximum of top flange width or bottom flange width
            Dim width# = Max(topFlangeWidth, bottomFlangeWidth)
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Centroid_X calculation
            'Calculate the individual areas for simplicity

            Dim a1# = bottomFlangeWidth * bottomFlangeThickness
            Dim a2# = (depth - bottomFlangeThickness - topFlangeThickness) * webThickness
            Dim a3# = topFlangeThickness * topFlangeWidth

            'Calculate the individual entity centroid
            'Check for possible conditions

            Dim x1#, x2#, x3#

            If (offSetWebBottom > offSetWebTop) Then
                x1 = bottomFlangeWidth / 2
                x2 = offSetWebBottom + webThickness / 2
                x3 = (offSetWebBottom - offSetWebTop) + topFlangeWidth / 2
            Else
                If (offSetWebBottom < offSetWebTop) Then
                    x1 = (offSetWebTop - offSetWebBottom) + bottomFlangeWidth / 2
                    x2 = offSetWebTop + webThickness / 2
                    x3 = topFlangeWidth / 2
                Else
                    'Considering offset for bottom web is equal to offset for top web from the reference axis
                    x1 = bottomFlangeWidth / 2
                    x2 = offSetWebBottom + webThickness / 2
                    x3 = topFlangeWidth / 2
                End If
            End If
            'Overall centroid about X
            Dim centroid_X# = (a1 * x1 + a2 * x2 + a3 * x3) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)

            'Centroid_Y calculation
            'Calculate entity centroid about Y

            Dim y1# = bottomFlangeThickness / 2
            Dim y2# = bottomFlangeThickness + (depth - bottomFlangeThickness - topFlangeThickness) / 2
            Dim y3# = bottomFlangeThickness + (depth - bottomFlangeThickness - topFlangeThickness) + topFlangeThickness / 2

            'Overall centroid about Y
            Dim centroid_Y# = (a1 * y1 + a2 * y2 + a3 * y3) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of horizontal distance from designated member edge to plastic neutral axis
            Dim horizontalDistance_PlasticNeutralAxis# = (topFlangeWidth * topFlangeThickness + (depth - topFlangeThickness) * webThickness + bottomFlangeThickness * webThickness - bottomFlangeWidth * bottomFlangeThickness) / (2 * webThickness)
            outputs.Add(IJUAL_xp, horizontalDistance_PlasticNeutralAxis)

            'Calculation of vertical distance from designated member edge to plastic neutral axis
            Dim verticalDistance_PlasticNeutralAxis# = ((depth - topFlangeThickness - bottomFlangeThickness) * (2 * Max(offSetWebTop, offSetWebBottom) + webThickness) + 2 * bottomFlangeThickness * width + topFlangeThickness * width * 2 - topFlangeWidth * topFlangeThickness - bottomFlangeWidth * bottomFlangeThickness) / (2 * depth)
            outputs.Add(IJUAL_yp, verticalDistance_PlasticNeutralAxis)

            'Calculation of moment of inertia about X axis
            'Parallel axis theorem is used for calculating the overall moment of inertia about X
            Dim momentOfInertia_X# = (1 / 12) * bottomFlangeWidth * Math.Pow(bottomFlangeThickness, 3) + a1 * Math.Pow((y1 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * webThickness * Math.Pow((depth - bottomFlangeThickness - topFlangeThickness), 3) + a2 * Math.Pow((y2 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * topFlangeWidth * Math.Pow(topFlangeThickness, 3) + a3 * Math.Pow((y3 - centroid_Y), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y axis
            'Parallel axis theorem is used for calculating the overall moment of inertia about Y
            Dim momentOfInertia_Y = (1 / 12) * bottomFlangeThickness * Math.Pow(bottomFlangeWidth, 3) + a1 * Math.Pow((x1 - centroid_X), 2)
            momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * (depth - bottomFlangeThickness - topFlangeThickness) * Math.Pow(webThickness, 3) + a2 * Math.Pow((x2 - centroid_X), 2)
            momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * topFlangeThickness * Math.Pow(topFlangeWidth, 3) + a3 * Math.Pow((x3 - centroid_X), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X and Y
            'The denominator in the formula checks the maximum value of (centroid_Y, (depth - centroid_Y))in case of X 
            'And (centroid_X, (width - centroid_X)) in case of Y and calculates the elastic section modulus 
            Dim elasticSectionModulus_X# = momentOfInertia_X / Max(centroid_Y, (depth - centroid_Y))
            Dim elasticSectionModulus_Y# = momentOfInertia_Y / Max(centroid_X, (width - centroid_X))
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of plastic section modulus about X
            Dim plasticSectionModulus_X# = area / 2 * (((webThickness * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) / 2) + (topFlangeThickness * topFlangeWidth * (depth - horizontalDistance_PlasticNeutralAxis - topFlangeThickness / 2))) / (webThickness * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) + topFlangeThickness * topFlangeWidth) + _
                (webThickness * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) / 2 + bottomFlangeThickness * bottomFlangeWidth * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness / 2)) / (webThickness * (horizontalDistance_PlasticNeutralAxis - bottomFlangeThickness) + bottomFlangeWidth * bottomFlangeThickness))
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)

            'Calculation of plastic section modulus about Y
            Dim plasticSectionModulus_Y# = area / 2 * (((depth - topFlangeThickness - bottomFlangeThickness) * (verticalDistance_PlasticNeutralAxis - Max(offSetWebTop, offSetWebBottom)) * (verticalDistance_PlasticNeutralAxis - Max(offSetWebTop, offSetWebBottom)) / 2 + bottomFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth)) * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth)) / 2 + topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth)) * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth)) / 2) / (topFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - topFlangeWidth)) + _
                (depth - topFlangeThickness - bottomFlangeThickness) * (verticalDistance_PlasticNeutralAxis - Max(offSetWebTop, offSetWebBottom)) + bottomFlangeThickness * (verticalDistance_PlasticNeutralAxis - (width - bottomFlangeWidth))) + ((depth - topFlangeThickness - bottomFlangeThickness) * (Max(offSetWebTop, offSetWebBottom) + webThickness - verticalDistance_PlasticNeutralAxis) * (Max(offSetWebTop, offSetWebBottom) + webThickness - verticalDistance_PlasticNeutralAxis) / 2 + _
                bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis) * (width - verticalDistance_PlasticNeutralAxis) / 2 + topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis) * (width - verticalDistance_PlasticNeutralAxis) / 2) / ((depth - topFlangeThickness - bottomFlangeThickness) * (Max(offSetWebTop, offSetWebBottom) + webThickness - verticalDistance_PlasticNeutralAxis) + topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis) + bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis)))
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia 
            Dim torsionalMomentOfInertia# = (topFlangeWidth * topFlangeThickness * topFlangeThickness * topFlangeThickness + bottomFlangeWidth * bottomFlangeThickness * bottomFlangeThickness * bottomFlangeThickness + webThickness * webThickness * webThickness * (depth - topFlangeThickness - bottomFlangeThickness)) / 3
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            'The following properties are left for edited by the user interactively , whose default values are set to zero/One as of now
            Dim radiusOfGyration_XY# = 0, warpingStaticalConstant# = 0, polarRadiusOfGyration_ShearCenter# = 0, warpingConstant# = 0, lengthExtension# = 0, depthExtension# = 0
            Dim flexuralConstant# = 0, flangeGage# = 0, webGage# = 0, bottomFlangeWidthExtension# = 0, topFlangeWidthExtension# = 0
            Dim isModifiable As Long = 0, systemProperties As Long = 1

            outputs.Add(IStructFlangedBoltGage_gf, flangeGage)
            outputs.Add(IStructFlangedBoltGage_gw, webGage)
            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalConstant)
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)
            outputs.Add(IUABuiltUpWeb_DepthExt, depthExtension)
            outputs.Add(IUABuiltUpBottomFlange_BottomFlangeWidthExt, bottomFlangeWidthExtension)
            outputs.Add(IUABuiltUpTopFlange_TopFlangeWidthExt, topFlangeWidthExtension)

            CalculateBUCSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUBoxFM section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia,
        ''' Horizontal and vertical distances from designated member edge to plastic neutral axis.
        ''' Properties like Warping constant, Polar radius of gyration about shear center, Radius of gyration about XY
        ''' Flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like depth, right web offset - top and bottom , left web offset - top and bottom, web thickness top and bottom flange' s width and thickness are necessary in calculating properties</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUBoxFMSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim offsetLeftWebTop# = CDbl(inputProperties.Item(IUABUBoxFlangeMajor_OffsetLeftWebTop))
            Dim offsetRightWebTop# = CDbl(inputProperties.Item(IUABUBoxFlangeMajor_OffsetRightWebTop))
            Dim offsetLeftWebBottom# = CDbl(inputProperties.Item(IUABUBoxFlangeMajor_OffsetLeftWebBot))
            Dim offsetRightWebBottom# = CDbl(inputProperties.Item(IUABUBoxFlangeMajor_OffsetRightWebBot))

            'offset top flange is the sum of right and left web's top offset 
            Dim offsetTopFlange# = offsetLeftWebTop + offsetRightWebTop

            'offset bottom flange is the sum of right and left web's bottom offset 
            Dim offsetBottomFlange# = offsetLeftWebBottom + offsetRightWebBottom

            'BUBoxFM is always hollow section
            Dim isHollow As Boolean = True

            'verifying the cross-section for its symmetry about X
            Dim isSymmetryAboutX As Boolean
            If ((topFlangeWidth = bottomFlangeWidth) And (topFlangeThickness = bottomFlangeThickness) And (offsetLeftWebTop = offsetLeftWebBottom) And (offsetRightWebTop = offsetRightWebBottom)) Then
                isSymmetryAboutX = True
            End If

            'verifying the cross-section for its symmetry about Y
            Dim isSymmetryAboutY As Boolean
            If ((topFlangeWidth = bottomFlangeWidth) And (offsetLeftWebTop = offsetRightWebTop) And (offsetLeftWebBottom = offsetRightWebBottom)) Then
                isSymmetryAboutY = True
            End If

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Overall area is calculated by summing up the individual areas of section ( top flange, bottom flange, webs)
            Dim area# = (topFlangeWidth * topFlangeThickness) + (bottomFlangeWidth * bottomFlangeThickness) + 2 * (depth - topFlangeThickness - bottomFlangeThickness) * webThickness
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Overall perimeter is calculated by summing up the individual lengths of the section ( top flange, bottom flange, webs)
            Dim perimeter# = topFlangeWidth + 2 * topFlangeThickness + bottomFlangeWidth + 2 * bottomFlangeThickness
            perimeter = perimeter + offsetTopFlange + offsetBottomFlange + 2 * depth + 2 * (depth - topFlangeThickness - bottomFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Width is set to maximum value of top flange width or bottom flange width
            Dim width# = Max(topFlangeWidth, bottomFlangeWidth)
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Calculation of centroid about X
            'Calculate the entities area for simplicity

            Dim A1# = (bottomFlangeThickness * bottomFlangeWidth)
            Dim A2# = webThickness * (depth - topFlangeThickness - bottomFlangeThickness)
            Dim A3# = webThickness * (depth - topFlangeThickness - bottomFlangeThickness)
            Dim A4# = (topFlangeWidth * topFlangeThickness)

            'Calculate the individual entity centroid
            'Check for possible conditions

            Dim x1#, x2#, x3#, x4#

            If (offsetLeftWebBottom > offsetLeftWebTop) Then
                x1 = bottomFlangeWidth / 2
                x2 = offsetLeftWebBottom + webThickness / 2
                x3 = bottomFlangeWidth - offsetRightWebBottom - webThickness / 2
                x4 = (offsetLeftWebBottom - offsetLeftWebTop) + topFlangeWidth / 2
            Else
                If (offsetLeftWebBottom < offsetLeftWebTop) Then
                    x1 = (offsetLeftWebTop - offsetLeftWebBottom) + bottomFlangeWidth / 2
                    x2 = offsetLeftWebTop + webThickness / 2
                    x3 = (offsetLeftWebTop - offsetLeftWebBottom) + bottomFlangeWidth - offsetRightWebBottom - webThickness / 2
                    x4 = topFlangeWidth / 2
                Else
                    'Considering offset for bottom left web is equal to offset for top left web  from the reference axis
                    x1 = bottomFlangeWidth / 2
                    x2 = offsetLeftWebBottom + webThickness / 2
                    x3 = bottomFlangeWidth - offsetRightWebBottom - webThickness / 2
                    x4 = topFlangeWidth / 2
                End If
            End If

            Dim centroid_X# = ((A1 * x1) + (A2 * x2) + (A3 * x3) + (A4 * x4)) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)

            'Calculation of centroid about Y
            'Calculate the entities centroid about Y

            Dim y1# = bottomFlangeThickness / 2
            Dim y2# = bottomFlangeThickness + (depth - topFlangeThickness - bottomFlangeThickness) / 2
            Dim y3# = bottomFlangeThickness + (depth - topFlangeThickness - bottomFlangeThickness) / 2
            Dim y4# = topFlangeThickness + (depth - topFlangeThickness - bottomFlangeThickness) + topFlangeThickness / 2

            'Overall centroid about Y is the summation of product of (entity area and entity centroid) divided by total area
            Dim centroid_Y# = ((A1 * y1) + (A2 * y2) + (A3 * y3) + (A4 * y4)) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'horizontal distance from designated member edge to plastic neutral axis calculation
            Dim horizontalDistance_PlasticNeutralAxis# = (topFlangeWidth * topFlangeThickness - bottomFlangeWidth * bottomFlangeThickness + 2 * webThickness * depth) / (4 * webThickness)

            'vertical distance from designated member edge to plastic neutral axis is simply half of the width
            Dim verticalDistance_PlasticNeutralAxis# = width / 2
            outputs.Add(IJUAL_xp, horizontalDistance_PlasticNeutralAxis)
            outputs.Add(IJUAL_yp, verticalDistance_PlasticNeutralAxis)

            'Moment of inertia about X axis
            'Parallel axis theorem is used in calculating the overall moment of inertia about X
            Dim momentOfInertia_X# = (1 / 12) * bottomFlangeWidth * Math.Pow(bottomFlangeThickness, 3) + A1 * Math.Pow((y1 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * webThickness * Math.Pow((depth - topFlangeThickness - bottomFlangeThickness), 3) + A2 * Math.Pow((y2 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * webThickness * Math.Pow((depth - topFlangeThickness - bottomFlangeThickness), 3) + A3 * Math.Pow((y3 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * topFlangeWidth * Math.Pow(topFlangeThickness, 3) + A4 * Math.Pow((y4 - centroid_Y), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y axis
            Dim momentOfInertia_Y#
            If (offsetLeftWebBottom > offsetLeftWebTop) Then
                momentOfInertia_Y# = (1 / 12) * bottomFlangeThickness * Math.Pow(bottomFlangeWidth, 3) + A1 * Math.Pow((x1 - centroid_X), 2)
                momentOfInertia_Y = momentOfInertia_Y + offsetLeftWebBottom + (1 / 12) * (depth - topFlangeThickness - bottomFlangeThickness) * Math.Pow(webThickness, 3) + A2 * Math.Pow((x2 - centroid_X), 2)
                momentOfInertia_Y = momentOfInertia_Y + offsetLeftWebBottom + (1 / 12) * (depth - topFlangeThickness - bottomFlangeThickness) * Math.Pow(webThickness, 3) + A2 * Math.Pow((x3 - centroid_X), 2)
                momentOfInertia_Y = momentOfInertia_Y + (offsetLeftWebBottom - offsetLeftWebTop) + (1 / 12) * topFlangeThickness * Math.Pow(topFlangeWidth, 3) + A4 * Math.Pow((x4 - centroid_X), 2)
            Else
                If (offsetLeftWebBottom < offsetLeftWebTop) Then
                    momentOfInertia_Y# = (offsetLeftWebTop - offsetLeftWebBottom) + (1 / 12) * bottomFlangeThickness * Math.Pow(bottomFlangeWidth, 3) + A1 * Math.Pow((x1 - centroid_X), 2)
                    momentOfInertia_Y = momentOfInertia_Y + offsetLeftWebTop + (1 / 12) * (depth - topFlangeThickness - bottomFlangeThickness) * Math.Pow(webThickness, 3) + A2 * Math.Pow((x2 - centroid_X), 2)
                    momentOfInertia_Y = momentOfInertia_Y + offsetLeftWebTop + (1 / 12) * (depth - topFlangeThickness - bottomFlangeThickness) * Math.Pow(webThickness, 3) + A2 * Math.Pow((x3 - centroid_X), 2)
                    momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * topFlangeThickness * Math.Pow(topFlangeWidth, 3) + A4 * Math.Pow((x4 - centroid_X), 2)
                Else
                    momentOfInertia_Y# = (1 / 12) * bottomFlangeThickness * Math.Pow(bottomFlangeWidth, 3) + A1 * Math.Pow((x1 - centroid_X), 2)
                    momentOfInertia_Y = momentOfInertia_Y + offsetRightWebBottom + (1 / 12) * (depth - topFlangeThickness - bottomFlangeThickness) * Math.Pow(webThickness, 3) + A2 * Math.Pow((x2 - centroid_X), 2)
                    momentOfInertia_Y = momentOfInertia_Y + offsetRightWebBottom + (1 / 12) * (depth - topFlangeThickness - bottomFlangeThickness) * Math.Pow(webThickness, 3) + A2 * Math.Pow((x3 - centroid_X), 2)
                    momentOfInertia_Y = momentOfInertia_Y + (1 / 12) * topFlangeThickness * Math.Pow(topFlangeWidth, 3) + A4 * Math.Pow((x4 - centroid_X), 2)
                End If
            End If
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X and Y
            'The denominator in the formula checks the maximum value of (centroid_Y, (depth - centroid_Y))in case of X 
            'And (centroid_X, (width - centroid_X)) in case of Y and calculates the elastic section modulus 
            Dim elasticSectionModulus_X# = momentOfInertia_X / Max(centroid_Y, (depth - centroid_Y))
            Dim elasticSectionModulus_Y# = momentOfInertia_Y / Max(centroid_X, (width - centroid_X))
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of plastic section modulus about X axis
            Dim plasticSectionModulus_X# = area / 2 * ((topFlangeWidth * topFlangeThickness * (depth - horizontalDistance_PlasticNeutralAxis - offsetTopFlange - topFlangeThickness / 2) + 2 * webThickness * (depth - horizontalDistance_PlasticNeutralAxis) * (depth - horizontalDistance_PlasticNeutralAxis) / 2) / (topFlangeWidth * topFlangeThickness + _
                2 * webThickness * (depth - horizontalDistance_PlasticNeutralAxis)) + (bottomFlangeWidth * bottomFlangeThickness * (horizontalDistance_PlasticNeutralAxis - offsetBottomFlange - offsetTopFlange / 2) + 2 * webThickness * horizontalDistance_PlasticNeutralAxis * horizontalDistance_PlasticNeutralAxis / 2) / (bottomFlangeWidth * bottomFlangeThickness + 2 * webThickness * horizontalDistance_PlasticNeutralAxis))
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)

            'Calculation of plastic section modulus about Y axis
            Dim plasticSectionModulus_Y# = area / 2 * ((depth * webThickness * (verticalDistance_PlasticNeutralAxis - webThickness / 2) + topFlangeThickness * (verticalDistance_PlasticNeutralAxis - webThickness) * (verticalDistance_PlasticNeutralAxis - webThickness) / 2 + bottomFlangeThickness * (verticalDistance_PlasticNeutralAxis - webThickness) * (verticalDistance_PlasticNeutralAxis - webThickness) / 2) / (depth * webThickness + _
                topFlangeThickness * (verticalDistance_PlasticNeutralAxis - webThickness) + bottomFlangeThickness * (verticalDistance_PlasticNeutralAxis - webThickness)) + (depth * webThickness * (width - verticalDistance_PlasticNeutralAxis - webThickness / 2) + topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - webThickness) * (width - verticalDistance_PlasticNeutralAxis - webThickness) / 2 + bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - webThickness) * (width - verticalDistance_PlasticNeutralAxis - webThickness) / 2) / (depth * webThickness + topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - webThickness) + bottomFlangeThickness * (width - verticalDistance_PlasticNeutralAxis - webThickness)))
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia 
            Dim torsionalMomentOfInertia# = (topFlangeWidth * topFlangeThickness * topFlangeThickness * topFlangeThickness + bottomFlangeWidth * bottomFlangeThickness * bottomFlangeThickness * bottomFlangeThickness + 2 * webThickness * webThickness * webThickness * (depth - topFlangeThickness - bottomFlangeThickness)) / 3
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            'The following properties are left for edited by the user interactively , whose default values are set to zero/One as of now
            Dim polarRadiusOfGyration_ShearCenter# = 0, warpingConstant# = 0, radiusOfGyration_XY# = 0, warpingStaticalMoment# = 0, flexuralConstant# = 0, lengthExtension# = 0, depthExtension# = 0
            Dim isModifiable As Long = 0, systemProperties As Long = 1

            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalMoment)
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)
            outputs.Add(IUABuiltUpWeb_DepthExt, depthExtension)

            CalculateBUBoxFMSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUTube section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia
        ''' Properties like Warping constant,Horizontal and vertical distances from designated member edge to plastic neutral axis.
        ''' Polar radius of gyration about shear center, Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like tube diameter and tube thickness  are used in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUTubeSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim tubeDiameter# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeDiameter))
            Dim tubeThickness# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeThickness))

            'BUTube section is considered as hollow
            Dim isHollow As Boolean = True

            'According to the geometry of the section, it is symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Depth and width are equal and are set to tube diameter 
            Dim depth# = tubeDiameter
            Dim width# = tubeDiameter
            outputs.Add(IStructCrossSectionDimensions_Depth, depth)
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Standard formulae are used for the calculations of area and perimeter
            Dim area# = Math.PI * (tubeDiameter * tubeDiameter - Math.Pow((tubeDiameter - (2 * tubeThickness)), 2)) / 4
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            Dim perimeter# = Math.PI * depth
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Since the section is symmetric about both X and Y axis, centroidX and centroidY are equal
            Dim centroid_X# = width / 2
            Dim centroid_Y# = centroid_X
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Since the section is symmetric about both X and Y axis, the values of moment of inertia about X and Y axis are equal
            'Standard formula is used in calculation of moment of inertia
            Dim momentOfInertia_X# = Math.PI * (Math.Pow(tubeDiameter, 4) - Math.Pow((tubeDiameter - (2 * tubeThickness)), 4)) / 64
            Dim momentOfInertia_Y# = momentOfInertia_X
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Since the section is symmetric about both X and Y axis, the values of plastic section modulus about X and Y axis are equal
            'Standard formula is used in calculation of plastic section modulus
            Dim plasticSectionModulus_X# = 1 / 6 * (Math.Pow(tubeDiameter, 3) - Math.Pow((tubeDiameter - (2 * tubeThickness)), 3))
            Dim plasticSectionModulus_Y# = plasticSectionModulus_X
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Since the section is symmetric about both X and Y axis, the values of radius of gyration about X and Y axis are equal
            'Standard formula is used in calculation of radius of gyration 
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Since the section is symmetric about both X and Y axis, the values of elastic section modulus about X and Y axis are equal
            'Standard formula is used in calculation of elastic section modulus
            Dim elasticSectionModulus_X# = Math.PI * (Math.Pow(depth, 4) - Math.Pow((depth - 2 * tubeThickness), 4)) / (32 * depth)
            Dim elasticSectionModulus_Y# = elasticSectionModulus_X
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Torsional moment of inertia is twice of that of moment of inertia about X or Y axis
            Dim torsionalMomentOfInertia# = 2 * momentOfInertia_X
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            'The following properties are left for edited by the user whose default values are set to zero/one
            Dim polarRadiusOfGyration_ShearCenter = 0, radiusOfGyration_XY# = 0, warpingStaticalConstant = 0, warpingConstant# = 0, flexuralConstant = 0, lengthExtension# = 0
            Dim isModifiable As Long = 0, systemProperties As Long = 1

            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)
            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalConstant)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)

            CalculateBUTubeSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUCone section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia.
        ''' Properties like Warping constant,Horizontal and vertical distances from designated member edge to plastic neutral axis.
        ''' Polar radius of gyration about shear center, Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like diameter (start and end),cone thickness  are used in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUConeSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim coneThickness# = CDbl(inputProperties.Item(IUABuiltUpCone_ConeThickness))
            Dim diameterStart# = CDbl(inputProperties.Item(IUABuiltUpCone_DiameterStart))
            Dim diameterEnd# = CDbl(inputProperties.Item(IUABuiltUpCone_DiameterEnd))

            'BUCone section is considered as hollow
            Dim isHollow As Boolean = True

            'According to the geometry of the section, it is symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Depth is set to the maximum value of diameter start/ end values 
            Dim depth# = Max(diameterStart, diameterEnd)
            outputs.Add(IStructCrossSectionDimensions_Depth, depth)

            'Width is set to the depth
            Dim width# = depth
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'All the below properties are calculated by considering the bidder side of the cross section

            'Calculation af area
            Dim area# = (Math.PI / 4) * (Math.Pow(depth, 2) - Math.Pow((depth - (2 * coneThickness)), 2))
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Calculation of perimeter
            Dim perimeter# = Math.PI * depth
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Calculation of centroid about X and Y
            Dim centroid_X# = width / 2
            Dim centroid_Y# = depth / 2
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of moment of inertia about X and Y
            Dim momentOfInertia_X# = Math.PI * (Math.Pow(depth, 4) - Math.Pow((2 * coneThickness), 4)) / 64
            Dim momentOfInertia_Y# = momentOfInertia_X
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Calculation of radius of gyration about X and Y
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = radiusOfGyration_X
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X and Y
            Dim elasticSectionModulus_X# = momentOfInertia_X / (32 * depth)
            Dim elasticSectionModulus_Y# = elasticSectionModulus_X
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of plastic section modulus about X and Y
            Dim plasticSectionModulus_X# = 1 / 6 * (Math.Pow(depth, 3) - Math.Pow((depth - (2 * coneThickness)), 3))
            Dim plasticSectionModulus_Y# = plasticSectionModulus_X
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia
            Dim torsionalMomentOfInertia# = 2 * momentOfInertia_X
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            'The following properties are left for edited by the user whose default values are set to zero/one
            Dim flexuralConstant# = 0, lengthExtension# = 0, polarRadiusOfGyration_ShearCenter = 0, warpingStaticalConstant = 0
            Dim isModifiable As Long = 0, systemProperties As Long = 1, radiusOfGyration_XY# = 0, warpingConstant# = 0

            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)
            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalConstant)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)

            CalculateBUConeSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUCan section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia
        ''' Properties like Warping constant,Horizontal and vertical distances from designated member edge to plastic neutral axis,
        ''' Polar radius of gyration about shear center, Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like tube diameter, tube thickness, diameter (start and end), length of cone(start and end),cone1 and cone2 thickness  are used in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUCanSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim tubeDiameter# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeDiameter))
            Dim tubeThickness# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeThickness))
            Dim diameterStart# = CDbl(inputProperties.Item(IUABuiltUpCan_DiameterStart))
            Dim diameterEnd# = CDbl(inputProperties.Item(IUABuiltUpCan_DiameterEnd))

            'BUCan section is considered as hollow
            Dim isHollow As Boolean = True

            'According to the geometry of the section, it is symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Depth is set to the maximum value of diameter start/ end values and tube diameter
            Dim depth# = Max(diameterStart, diameterEnd)
            depth = Max(depth, tubeDiameter)
            outputs.Add(IStructCrossSectionDimensions_Depth, depth)

            'width is also set to the maximum value of diameter start/ end values and tube diameter
            Dim width# = Max(diameterStart, diameterEnd)
            width = Max(width, tubeDiameter)
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Calculation of area
            Dim area# = Math.PI * (depth * depth - Math.Pow((depth - (2 * tubeThickness)), 2)) / 4
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Calculation of perimeter
            Dim perimeter# = Math.PI * depth
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Calculation of centroid about X and Y
            Dim centroid_X# = width / 2
            Dim centroid_Y# = depth / 2
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of moment of inertia about X and Y
            Dim momentOfInertia_X# = Math.PI * (Math.Pow(depth, 4) - Math.Pow((2 * tubeThickness), 4)) / 64
            Dim momentOfInertia_Y# = momentOfInertia_X
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Calculation of radius of gyration about X and Y
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X and Y
            Dim elasticSectionModulus_X# = momentOfInertia_X / (32 * depth)
            Dim elasticSectionModulus_Y# = elasticSectionModulus_X
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of plastic section modulus about X and Y
            Dim plasticSectionModulus_X# = 1 / 6 * (Math.Pow(depth, 3) - Math.Pow((depth - (2 * tubeThickness)), 3))
            Dim plasticSectionModulus_Y# = plasticSectionModulus_X
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia
            Dim torsionalMomentOfInertia# = 2 * momentOfInertia_X
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            'The following properties are left for edited by the user whose default values are set to zero/one
            Dim polarRadiusOfGyration_ShearCenter = 0, radiusOfGyration_XY# = 0, warpingStaticalConstant = 0, warpingConstant# = 0
            Dim flexuralConstant# = 0, lengthExtension# = 0, isModifiable As Long = 0, systemProperties As Long = 1

            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)
            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalConstant)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)

            CalculateBUCanSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUEndCan section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia
        ''' Properties like Warping constant,Horizontal and vertical distances from designated member edge to plastic neutral axis,
        ''' Polar radius of gyration about shear center, Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like tube diameter, tube thickness, cone diameter, cone length are used in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUEndCanSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim tubeDiameter# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeDiameter))
            Dim tubeThickness# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeThickness))
            Dim coneDiameter# = CDbl(inputProperties.Item(IUABUILTUPENDCAN_DIAMETERCONE))

            'BUEndCan section is considered as hollow
            Dim isHollow As Boolean = True

            'According to the geometry of the section, it is symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'Depth is set to the maximum value of tube diameter or cone diameter
            Dim depth# = Max(tubeDiameter, coneDiameter)
            outputs.Add(IStructCrossSectionDimensions_Depth, depth)

            'Width is set to depth
            Dim width# = depth
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Calculation of area
            Dim area# = Math.PI * (depth * depth - Math.Pow((depth - (2 * tubeThickness)), 2)) / 4
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Calculation of perimeter
            Dim perimeter# = Math.PI * depth
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Calculation of centroid about X and Y
            Dim centroid_X# = width / 2
            Dim centroid_Y# = depth / 2
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of moment of inertia about X and Y
            Dim momentOfInertia_X# = Math.PI * (Math.Pow(depth, 4) - Math.Pow((2 * tubeThickness), 4)) / 64
            Dim momentOfInertia_Y# = momentOfInertia_X
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Calculation of radius of gyration about X and Y
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X and Y
            Dim elasticSectionModulus_X# = momentOfInertia_X / (32 * depth)
            Dim elasticSectionModulus_Y# = momentOfInertia_X
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of plastic section modulus about X and Y
            Dim plasticSectionModulus_X# = 1 / 6 * (Math.Pow(depth, 3) - Math.Pow((depth - (2 * tubeThickness)), 3))
            Dim plasticSectionModulus_Y# = plasticSectionModulus_X
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Calculation of torsional moment of inertia
            Dim torsionalMomentOfInertia# = 2 * momentOfInertia_X
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            Dim isTransitionAtStart As Boolean

            'The following properties are left for edited by the user whose default values are set to zero/one
            Dim lengthExtension# = 0, depthExtension# = 0, polarRadiusOfGyration_ShearCenter = 0, radiusOfGyration_XY# = 0, warpingConstant# = 0
            Dim warpingStaticalConstant = 0, flexuralConstant# = 0, isModifiable As Long = 0, systemProperties As Long = 1

            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)
            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalConstant)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpEndCan_IsTransitionAtStart, CDbl(isTransitionAtStart))
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)

            CalculateBUEndCanSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUFlat section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis
        ''' Properties like Torsional moment of inertia, Warping constant,
        ''' Horizontal and vertical distances from designated member edge to plastic neutral axis, Polar radius of gyration about shear center
        ''' Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties depth,web thickness are used in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUFlatSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))

            'BUFlat section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is symmetric about X and Y axis.
            Dim isSymmetryAboutX As Boolean = True
            Dim isSymmetryABoutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryABoutY))

            'In this section, width is set to web thickness
            Dim width# = webThickness
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Calculation of area
            Dim area# = (depth * width)
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Calculation of perimeter
            Dim perimeter# = 2 * (depth + width)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Calculation of centroid ablut X and Y axis
            Dim centroid_X# = 0.5 * width
            Dim centroid_Y# = 0.5 * depth
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculation of moment of inertia about X axis
            Dim momentOfInertia_X# = (width * Math.Pow(depth, 3)) / 12
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y axis
            Dim momentOfInertia_Y# = (depth * Math.Pow(width, 3)) / 12
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            'Calculation of radius of gyration about X axis
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)

            'Calculation of radius of gyration about X axis
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Calculation of elastic section modulus about X axis
            Dim elasticSectionModulus_X# = width * depth * depth / 6
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)

            'Calculation of elastic section modulus about Y axis
            Dim elasticSectionModulus_Y# = depth * width * width / 6
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Calculation of plastic section modulus about X axis
            Dim plasticSectionModulus_X# = width * depth * depth / 4
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)

            'Calculation of plastic section modulus about Y axis
            Dim plasticSectionModulus_Y# = depth * width * width / 4
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'The following properties are left for edited by the user whose default values are set to zero/one
            Dim lengthExtension# = 0, depthExtension# = 0
            Dim isModifiable As Long = 0, systemProperties As Long = 1, warpingStaticalConstant = 0
            Dim torsionalMomentOfInertia# = 0, warpingConstant# = 0, flexuralConstant# = 0, polarRadiusOfGyration_ShearCenter = 0, radiusOfGyration_XY# = 0

            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)
            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalConstant)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarRadiusOfGyration_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)
            outputs.Add(IUABuiltUpWeb_DepthExt, depthExtension)

            CalculateBUFlatSectionProperties = outputs

        End Function

        ''' <summary>
        ''' Function to calculate BUTee section properties like area, perimeter, unit weight, Centroid about X-axis and Y-axis, Moment of inertia about X-axis and Y-axis,
        ''' Plastic and Elastic section modulus about X-axis and Y-axis, Radius of gyration about X-axis and Y-axis, Torsional moment of inertia, Warping constant,
        ''' Horizontal and vertical distances from designated member edge to plastic neutral axis.
        ''' Properties like Polar radius of gyration about shear center, Radius of gyration about XY, flexural constant, warping statical moment etc are left for editing according to user calculations.
        ''' </summary>
        ''' <param name="inputProperties">Input properties like top flange width, depth, top flange thickness, web thickness are used in calculating properties.</param>
        ''' <returns>This returns the calculated cross section properties name and values.</returns>
        Private Function CalculateBUTeeSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As Dictionary(Of String, Double)

            Dim outputs As New Dictionary(Of String, Double)

            'Following are the required inputs which are necessary in calculation of properties
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))

            'BUTee section is not considered as hollow
            Dim isHollow As Boolean

            'According to the geometry of the section, it is symmetric about only Y axis.
            Dim isSymmetryAboutX As Boolean
            Dim isSymmetryAboutY As Boolean = True

            outputs.Add(IStructCrossSectionDesignProperties_IsHollow, CDbl(isHollow))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutX, CDbl(isSymmetryAboutX))
            outputs.Add(IStructCrossSectionDesignProperties_IsSymmetricAboutY, CDbl(isSymmetryAboutY))

            'width is set to top flange width in BUTee section
            Dim width# = topFlangeWidth
            outputs.Add(IStructCrossSectionDimensions_Width, width)

            'Total area is calculated by summing up the individual areas of the BUTee section (flange, web)
            Dim area# = width * topFlangeThickness + webThickness * (depth - topFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Area, area)

            'Total perimeter is calculated by summing up the individual perimeters of the BUTee section (flange, web)
            Dim perimeter# = 2 * (width + topFlangeThickness) + 2 * (depth - topFlangeThickness)
            outputs.Add(IStructCrossSectionDimensions_Perimeter, perimeter)

            'The unit weight is calculated by multiplying overall area of the section and material density
            'It is calculated per unit length. Units-->(kg/m)
            'Interactively the units can be set to metric or any other as required

            Dim unitWeight# = area * 7850
            outputs.Add(IStructCrossSectionUnitWeight_UnitWeight, unitWeight)

            'Calculation of centroid about X
            Dim centroid_X# = width / 2
            outputs.Add(IStructCrossSectionDesignProperties_CentroidX, centroid_X)

            'Calculation of centroid about Y
            'Calculate the individual entities area for simplicity
            Dim a1# = webThickness * (depth - topFlangeThickness)
            Dim a2# = width * topFlangeThickness

            'Calculate the entity centroid about Y
            Dim y1# = (depth - topFlangeThickness) / 2
            Dim y2# = (depth - topFlangeThickness) + topFlangeThickness / 2

            'Overall centroid about Y 
            Dim centroid_Y# = (a1 * y1 + a2 * y2) / area
            outputs.Add(IStructCrossSectionDesignProperties_CentroidY, centroid_Y)

            'Calculations of horizontal and vertical distances from designated member edge to plastic neutral axis
            'Horizontal distance calculation is checked for certain conditions before proceeding 
            Dim horizontalDistance_PlasticNeutralAxis#
            If width * topFlangeThickness < (depth - webThickness) * webThickness Then
                horizontalDistance_PlasticNeutralAxis = (width * topFlangeThickness + (depth - topFlangeThickness) * webThickness) / (2 * webThickness)
            Else
                horizontalDistance_PlasticNeutralAxis = (2 * width * depth - width * topFlangeThickness - webThickness * (depth - topFlangeThickness)) / (2 * width)
            End If
            outputs.Add(IJUAL_xp, horizontalDistance_PlasticNeutralAxis)

            'Since BUTee is symmetric about Y axis, vertical distance is simply half of width
            Dim verticalDistance_PlasticNeutralAxis# = width / 2
            outputs.Add(IJUAL_yp, verticalDistance_PlasticNeutralAxis)

            'Calculation of moment of inertia about X
            Dim momentOfInertia_X# = (1 / 12) * webThickness * Math.Pow((depth - topFlangeThickness), 3) + a1 * Math.Pow((y1 - centroid_Y), 2)
            momentOfInertia_X = momentOfInertia_X + (1 / 12) * width * Math.Pow(topFlangeThickness, 3) + a2 * Math.Pow((y2 - centroid_Y), 2)
            outputs.Add(IStructCrossSectionDesignProperties_Ixx, momentOfInertia_X)

            'Calculation of moment of inertia about Y
            Dim momentOfInertia_Y# = (1 / 12) * (depth - topFlangeThickness) * Math.Pow(webThickness, 3) + (1 / 12) * topFlangeThickness * Math.Pow(width, 3)
            outputs.Add(IStructCrossSectionDesignProperties_Iyy, momentOfInertia_Y)

            ''Radius of gyration is the square root of (moment of inertia(about particular axis) to the area)
            Dim radiusOfGyration_X# = Math.Sqrt(momentOfInertia_X / area)
            Dim radiusOfGyration_Y# = Math.Sqrt(momentOfInertia_Y / area)
            outputs.Add(IStructCrossSectionDesignProperties_Rxx, radiusOfGyration_X)
            outputs.Add(IStructCrossSectionDesignProperties_Ryy, radiusOfGyration_Y)

            'Elastic section modulus calculations
            'The denominator in the formula checks the maximum value of (centroid_Y, (depth - centroid_Y))in case of X 
            'And (centroid_X, (flangeWidth - centroid_X)) in case of Y and calculates the elastic section modulus 
            Dim elasticSectionModulus_X# = momentOfInertia_X / Max(centroid_Y, (depth - centroid_Y))
            Dim elasticSectionModulus_Y# = momentOfInertia_Y / Max(centroid_X, (width - centroid_X))
            outputs.Add(IStructCrossSectionDesignProperties_Sxx, elasticSectionModulus_X)
            outputs.Add(IStructCrossSectionDesignProperties_Syy, elasticSectionModulus_Y)

            'Plastic section modulus calculations
            'Plastic section modulus along X axis is checked for certain conditions before proceeding 
            Dim plasticSectionModulus_X#
            If width * topFlangeThickness < (depth - webThickness) * webThickness Then
                plasticSectionModulus_X = area / 2 * ((webThickness * Math.Pow((depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis), 2) / 2 + topFlangeThickness * width * (depth - horizontalDistance_PlasticNeutralAxis - topFlangeThickness / 2)) / (webThickness * (depth - topFlangeThickness - horizontalDistance_PlasticNeutralAxis) + width * topFlangeThickness) + webThickness * Math.Pow(horizontalDistance_PlasticNeutralAxis, 2) / 2)
            Else
                plasticSectionModulus_X = area / 2 * ((width * (depth - horizontalDistance_PlasticNeutralAxis) * (depth - horizontalDistance_PlasticNeutralAxis) / 2) / (width * (depth - horizontalDistance_PlasticNeutralAxis)) + (width * (horizontalDistance_PlasticNeutralAxis - depth + topFlangeThickness) * (horizontalDistance_PlasticNeutralAxis - depth + topFlangeThickness) / 2 + _
                    webThickness * (depth - topFlangeThickness) * (2 * horizontalDistance_PlasticNeutralAxis - depth + topFlangeThickness) / 2) / (width * (horizontalDistance_PlasticNeutralAxis - depth + topFlangeThickness) + webThickness * (depth - topFlangeThickness)))
            End If
            outputs.Add(IStructCrossSectionDesignProperties_Zxx, plasticSectionModulus_X)

            'Calculation os plastic section modulus along Y axis
            Dim plasticSectionModulus_Y# = area / 2 * ((topFlangeThickness * verticalDistance_PlasticNeutralAxis * verticalDistance_PlasticNeutralAxis / 2 + (depth - topFlangeThickness) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) / 2) / (topFlangeThickness * verticalDistance_PlasticNeutralAxis + (depth - topFlangeThickness) * (verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2)) + (topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis) * (width - verticalDistance_PlasticNeutralAxis) / 2 + _
                (depth - topFlangeThickness) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2) / 2) / (topFlangeThickness * (width - verticalDistance_PlasticNeutralAxis) + (depth - topFlangeThickness) * (width - verticalDistance_PlasticNeutralAxis - (width - webThickness) / 2)))
            outputs.Add(IStructCrossSectionDesignProperties_Zyy, plasticSectionModulus_Y)

            'Torsional moment of inertia and warping constant calculations
            'Standard relations are used in calculating the same
            Dim torsionalMomentOfInertia# = (width * topFlangeThickness * topFlangeThickness * topFlangeThickness + webThickness * webThickness * webThickness * (depth - topFlangeThickness)) / 3
            outputs.Add(IStructCrossSectionDesignProperties_J, torsionalMomentOfInertia)

            Dim warpingConstant# = (width * width * width * topFlangeThickness * topFlangeThickness * topFlangeThickness / 4 + (depth - topFlangeThickness) * (depth - topFlangeThickness) * (depth - topFlangeThickness) * webThickness * webThickness * webThickness) / 36
            outputs.Add(IStructCrossSectionDesignProperties_Cw, warpingConstant)

            'Polar radius of gyration about shear center,Radius of gyration about XY, flexural constant, warping statical moment 
            'Properties like isModifiable,systemProperties,length extension, depth extension, top flange width extension are left for edited by the user interactively
            'The default values of these properties are set to zero/one 
            Dim radiusOfGyration_XY# = 0, warpingStaticalConstant# = 0, polarMomentOfInertia_ShearCenter# = 0, flexuralConstant# = 0
            Dim isModifiable As Long = 0, systemProperties As Long = 1
            Dim lengthExtension# = 0, depthExtension# = 0, topFlangeWidthExtension# = 0, flangeGage# = 0, webGage# = 0

            outputs.Add(IStructFlangedBoltGage_gf, flangeGage)
            outputs.Add(IStructFlangedBoltGage_gw, webGage)
            outputs.Add(IStructCrossSectionDesignProperties_Rxy, radiusOfGyration_XY)
            outputs.Add(IStructCrossSectionDesignProperties_Sw, warpingStaticalConstant)
            outputs.Add(IStructCrossSectionDesignProperties_ro, polarMomentOfInertia_ShearCenter)
            outputs.Add(IStructCrossSectionDesignProperties_H, flexuralConstant)
            outputs.Add(IUABuiltUpCompute_IsModifiable, isModifiable)
            outputs.Add(IUABuiltUpCompute_SectionProperties, systemProperties)
            outputs.Add(IUABuiltUpLengthExt_LengthExt, lengthExtension)
            outputs.Add(IUABuiltUpWeb_DepthExt, depthExtension)
            outputs.Add(IUABuiltUpTopFlange_TopFlangeWidthExt, topFlangeWidthExtension)

            CalculateBUTeeSectionProperties = outputs

        End Function

        Private Function Max(ByVal x#, ByVal y#) As Double

            'Calculation of maximum value
            If x >= y Then
                Max = x
            Else
                Max = y
            End If

        End Function

        Private Function Min(ByVal x#, ByVal y#) As Double

            'Calculation of minimum value
            If x <= y Then
                Min = x
            Else
                Min = y
            End If

        End Function

#Region "Validate k-design"

        ''' <summary>
        ''' Determines whether [is valid K design] [the specified kdesign].
        ''' </summary>
        ''' <param name="kdesign">The kdesign.</param>
        ''' <param name="flangeThickness">The flange thickness.</param>
        ''' <param name="depth">The depth.</param>
        ''' <returns>
        ''' <c>true</c> if [is valid K design] [the specified kdesign]; otherwise, <c>false</c>.
        ''' </returns>
        Private Function IsValidKDesign(ByVal kdesign#, ByVal flangeThickness#, ByVal depth#) As Boolean

            If ((kdesign >= flangeThickness) And (kdesign < depth / 2) And (flangeThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

        ''' <summary>
        ''' Determines whether [is valid K design for single flange] [the specified kdesign].
        ''' </summary>
        ''' <param name="kdesign">The kdesign.</param>
        ''' <param name="flangeThickness">The flange thickness.</param>
        ''' <param name="depth">The depth.</param>
        ''' <returns>
        ''' <c>true</c> if [is valid K design for single flange] [the specified kdesign]; otherwise, <c>false</c>.
        ''' </returns>
        Private Function IsValidKDesignForSingleFlange(ByVal kdesign#, ByVal flangeThickness#, ByVal depth#) As Boolean

            If ((kdesign >= flangeThickness) And (kdesign < depth) And (flangeThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate web thickness for AISC CSTypes"

        Private Function IsValidCSWebThickness(ByVal webThickness#, ByVal Width#) As Boolean

            If ((Width > webThickness) And (webThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate web thickness for AISC 2LType"

        Private Function IsValid2LWebThickness(ByVal webThickness#, ByVal flangeWidth#) As Boolean

            If ((flangeWidth > webThickness) And (webThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate depth for AISC CSTypes"

        Private Function IsValidDepth(ByVal crossSectionType$, ByVal depth#, ByVal flangeThickness#, ByVal kdesign#) As Boolean

            If crossSectionType = "W" Or crossSectionType = "M" Or crossSectionType = "S" Or crossSectionType = "HP" Or crossSectionType = "C" Or crossSectionType = "MC" Then
                If ((depth > 2 * flangeThickness) And (depth > 2 * kdesign)) Then
                    Return True
                Else
                    Return False
                End If
            Else
                If ((depth > flangeThickness) And (depth > kdesign)) Then
                    Return True
                Else
                    Return False
                End If
            End If

        End Function

#End Region

#Region "Validate PIPE"

        Private Function IsValidPIPE(ByVal depth#, ByVal thicknessDesign#) As Boolean

            If ((depth > 2 * thicknessDesign) And (thicknessDesign > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate CS"

        Private Function IsValidCS(ByVal depth#) As Boolean

            If (depth > 0) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate RS"

        Private Function IsValidRS(ByVal depth#, ByVal width#) As Boolean

            If ((depth > 0) And (width > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate HSSR"

        Private Function IsValidHSSR(ByVal depth#, ByVal width#, ByVal thicknessDesign#) As Boolean

            If ((depth > 2 * thicknessDesign) And (width > 2 * thicknessDesign) And (thicknessDesign > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate BUTee"

        Private Function IsValidBUTee(ByVal depth#, ByVal topFlangeWidth#, ByVal topFlangeThickness#, ByVal webThickness#) As Boolean

            If ((topFlangeThickness < depth / 2) And (webThickness < topFlangeWidth) And (webThickness > 0) And (topFlangeThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate EndCan"

        Private Function IsValidEndCan(ByVal tubeDiameter#, ByVal tubeThickness#, ByVal diameterCone#) As Boolean

            If ((tubeDiameter > 2 * tubeThickness) And (tubeDiameter > diameterCone) And (tubeThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate BUTube"

        Private Function IsValidBUTube(ByVal tubeDiameter#, ByVal tubeThickness#) As Boolean

            If ((tubeDiameter > 2 * tubeThickness) And (tubeThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate BUCan"

        Private Function IsValidBUCan(ByVal diameterStart#, ByVal diameterEnd#, ByVal tubeDiameter#, ByVal tubeThickness#, ByVal cone1Thickness#, ByVal cone2Thickness#) As Boolean

            If ((tubeDiameter > diameterStart) And (tubeDiameter > diameterEnd) And (tubeDiameter > 2 * tubeThickness) And (diameterStart > 2 * cone1Thickness) And (diameterEnd > 2 * cone2Thickness) And (cone1Thickness > 0) And (cone2Thickness > 0) And (tubeThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate BUCone"

        Private Function IsValidBUCone(ByVal diameterStart#, ByVal diameterEnd#, ByVal coneThickness#) As Boolean

            If ((diameterStart > 2 * coneThickness) And (diameterEnd > 2 * coneThickness) And (coneThickness > 0) And (diameterStart <> diameterEnd)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate BUFlat"

        Private Function IsValidBUFlat(ByVal depth#, ByVal webThickness#) As Boolean

            If ((depth > 0) And (webThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate BUBox and BUC"

        Private Function IsValidBUBoxAndBUC(ByVal depth#, ByVal topFlangeThickness#, ByVal bottomFlangeThickness#, ByVal topFlangeWidth#, ByVal bottomFlangeWidth#, ByVal webThickness#) As Boolean

            If (depth > (topFlangeThickness + bottomFlangeThickness) And (webThickness < topFlangeWidth) And (webThickness < bottomFlangeWidth) And (webThickness > 0) And (topFlangeThickness > 0) And (bottomFlangeThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate BUL"

        Private Function IsValidBUL(ByVal depth#, ByVal bottomFlangeThickness#, ByVal bottomFlangeWidth#, ByVal webThickness#) As Boolean

            If ((depth > bottomFlangeThickness) And (webThickness < bottomFlangeWidth) And (webThickness > 0) And (bottomFlangeThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate BUIHaunch"

        Private Function IsValidBUIHaunch(ByVal depth#, ByVal topFlangeThickness#, ByVal bottomFlangeThickness#, ByVal topFlangeWidth#, ByVal bottomFlangeWidth#, ByVal webThickness#, ByVal lengthStart#, ByVal lengthEnd#, ByVal depthStart#, ByVal depthEnd#, ByVal depthHaunch#) As Boolean

            If (depth > (topFlangeThickness + bottomFlangeThickness) And (webThickness < topFlangeWidth) And (webThickness < bottomFlangeWidth) And (webThickness > 0) And (topFlangeThickness > 0) And (bottomFlangeThickness > 0) And (depthHaunch <> depthStart) And (depthHaunch <> depthEnd) And (depthStart > 0) And (depthEnd > 0)) Then
                Return True
            Else
                Return False
            End If
        End Function

#End Region

#Region "Validate BUIUE"

        Private Function IsValidBUIUE(ByVal depth#, ByVal topFlangeThickness#, ByVal bottomFlangeThickness#, ByVal topFlangeWidth#, ByVal bottomFlangeWidth#, ByVal webThickness#) As Boolean

            If (depth > (topFlangeThickness + bottomFlangeThickness) And (webThickness < topFlangeWidth) And (webThickness < bottomFlangeWidth) And (webThickness > 0) And (topFlangeThickness > 0) And (bottomFlangeThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

#Region "Validate BUITapWeb"

        Private Function IsValidBUITapWeb(ByVal depth#, ByVal topFlangeThickness#, ByVal bottomFlangeThickness#, ByVal topFlangeWidth#, ByVal bottomFlangeWidth#, ByVal webThickness#) As Boolean

            If (depth > (topFlangeThickness + bottomFlangeThickness) And (webThickness < topFlangeWidth) And (webThickness < bottomFlangeWidth) And (webThickness > 0) And (topFlangeThickness > 0) And (bottomFlangeThickness > 0)) Then
                Return True
            Else
                Return False
            End If
        End Function

#End Region

#Region "Validate BUI"

        Private Function IsValidBUI(ByVal depth#, ByVal topFlangeThickness#, ByVal topFlangeWidth#, ByVal webThickness#) As Boolean

            If ((depth > topFlangeThickness) And (webThickness < topFlangeWidth) And (webThickness > 0) And (topFlangeThickness > 0)) Then
                Return True
            Else
                Return False
            End If

        End Function

#End Region

        ''' <summary>
        ''' Gets the plastic neutral axis distance for T section.
        ''' </summary>
        ''' <param name="depth">The depth.</param>
        ''' <param name="width">The width.</param>
        ''' <param name="webThickness">The web thickness.</param>
        ''' <param name="flangeThickness">The flange thickness.</param>
        ''' <returns></returns>
        Private Function GetPlasticNeutralAxisDistanceForTSection(ByVal depth As Double, ByVal width As Double, ByVal webThickness As Double, ByVal flangeThickness As Double) As Double
            Dim flangeArea# = flangeThickness * width
            Dim webArea# = (depth - flangeThickness) * webThickness
            'if the web area is more than the flange area then the axis is passing through the web and vice versa
            If flangeArea < webArea Then
                GetPlasticNeutralAxisDistanceForTSection = ((depth + flangeThickness) / 2) - (flangeArea / (2 * webThickness))
            Else
                GetPlasticNeutralAxisDistanceForTSection = (flangeArea + webArea) / (2 * width)
            End If

        End Function

        ''' <summary>
        ''' Validates the I section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateISectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))
            Dim kdesign# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_kdesign))

            'Validating web thickness
            If Not IsValidCSWebThickness(webThickness, width) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_WEBTHICKNESS_NOT_LESS_WIDTH, _
                "Web thickness cannot be less than or equal to zero. Also, web thickness should be less than the width.")
            End If

            'Validating for depth
            If Not IsValidDepth(CSType, depth, flangeThickness, kdesign) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_DEPTH_TWICE_FLANGE_KDESIGN, _
                "Depth should be greater than twice of flange thickness and twice of kdesign.")
            End If

            'Validating kdesign
            If Not IsValidKDesign(kdesign, flangeThickness, depth) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_KDESIGN_GREATER_FLANGE_LESS_DEPTH, _
               "The kdesign value should be greater than or equal to the flange thickness and should be less than half of the depth.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the S section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateSSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))
            Dim kdesign# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_kdesign))

            'Validating web thickness
            If Not IsValidCSWebThickness(webThickness, width) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_WEBTHICKNESS_NOT_LESS_WIDTH, _
                "Web thickness cannot be less than or equal to zero. Also, web thickness should be less than the width.")
            End If

            'Validating depth
            If Not IsValidDepth(CSType, depth, flangeThickness, kdesign) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_DEPTH_TWICE_FLANGE_KDESIGN, _
                "Depth should be greater than twice the flange thickness and twice the kdesign.")
            End If

            'Validating kdesign
            If Not IsValidKDesign(kdesign, flangeThickness, depth) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_KDESIGN_GREATER_FLANGE_LESS_DEPTH, _
                "The kdesign value should be greater than or equal to the flange thickness and should be less than half of the depth.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the T section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateTSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))
            Dim kdesign# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_kdesign))

            'Validating web thickness
            If Not IsValidCSWebThickness(webThickness, width) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_WEBTHICKNESS_NOT_LESS_WIDTH, _
               "Web thickness cannot be less than or equal to zero. Also, web thickness should be less than the width.")
            End If

            'Validating depth
            If Not IsValidDepth(CSType, depth, flangeThickness, kdesign) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_DEPTH_GREATER_FLANGE_KDESIGN, _
                "Depth should be greater than flange thickness and kdesign.")
            End If

            'Validating kdesign
            If Not IsValidKDesignForSingleFlange(kdesign, flangeThickness, depth) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_KDESIGN_GREATER_FLANGE_LESS_DEPTH, _
                "The kdesign value should be greater than or equal to the flange thickness and should be less than half of the depth.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the ST section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateSTSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))
            Dim kdesign# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_kdesign))

            'Validating web thickness
            If Not IsValidCSWebThickness(webThickness, width) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_WEBTHICKNESS_NOT_LESS_WIDTH, _
                "Web thickness cannot be less than or equal to zero. Also, web thickness should be less than the width.")
            End If

            'Validating depth
            If Not IsValidDepth(CSType, depth, flangeThickness, kdesign) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_DEPTH_GREATER_FLANGE_KDESIGN, _
                "Depth should be greater than flange thickness and kdesign.")
            End If

            'Validating kdesign
            If Not IsValidKDesignForSingleFlange(kdesign, flangeThickness, depth) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_KDESIGN_GREATER_FLANGE_LESS_DEPTH, _
                "The kdesign value should be greater than or equal to the flange thickness and should be less than half of the depth.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the C section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateCSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))
            Dim kdesign# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_kdesign))

            'Validating web thickness
            If Not IsValidCSWebThickness(webThickness, width) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_WEBTHICKNESS_NOT_LESS_WIDTH, _
                "Web thickness cannot be less than or equal to zero. Also, web thickness should be less than the width.")
            End If

            'Validating depth
            If Not IsValidDepth(CSType, depth, flangeThickness, kdesign) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_DEPTH_TWICE_FLANGE_KDESIGN, _
                "Depth should be greater than twice the flange thickness and twice the kdesign.")
            End If

            'Validating kdesign
            If Not IsValidKDesignForSingleFlange(kdesign, flangeThickness, depth) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_KDESIGN_GREATER_FLANGE_LESS_DEPTH, _
                "The kdesign value should be greater than or equal to the flange thickness and should be less than half of the depth.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the L section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateLSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))
            Dim kdesign# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_kdesign))

            'Validating web thickness
            If Not IsValidCSWebThickness(webThickness, width) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_WEBTHICKNESS_NOT_LESS_WIDTH, _
                "Web thickness cannot be less than or equal to zero. Also, web thickness should be less than the width.")
            End If

            'Validating depth
            If Not IsValidDepth(CSType, depth, flangeThickness, kdesign) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_DEPTH_GREATER_FLANGE_KDESIGN, _
                "Depth should be greater than flange thickness and kdesign.")
            End If

            'Validating kdesign
            If Not IsValidKDesignForSingleFlange(kdesign, flangeThickness, depth) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_KDESIGN_GREATER_FLANGE_LESS_DEPTH, _
                "The kdesign value should be greater than or equal to the flange thickness and should be less than half of the depth.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the 2L section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function Validate2LSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim flangeWidth# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_bf))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim flangeThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tf))
            Dim webThickness# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_tw))
            Dim kdesign# = CDbl(inputProperties.Item(IStructFlangedSectionDimensions_kdesign))

            'Validating web thickness
            If Not IsValid2LWebThickness(webThickness, flangeWidth) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_WEBTHICKNESS_LESS_FLANGEWIDTH, _
                "Web thickness should always be less than the flange width.")
            End If

            'Validating depth
            If Not IsValidDepth(CSType, depth, flangeThickness, kdesign) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_DEPTH_GREATER_FLANGE_KDESIGN, _
                "Depth should be greater than flange thickness and kdesign.")
            End If

            'Validating kdesign. Commented as kdesign values are not given for 2L section in the spreadsheet. Check should be added back when that is fixed.
            'If Not IsValidKDesignForSingleFlange(kdesign, flangeThickness, depth) Then
            '    Return GetString(SectionLibraryCalculatorResourceIDs.MSG_KDESIGN_GREATER_FLANGE_LESS_DEPTH, _
            '    "The kdesign value should be greater than or equal to the flange thickness and should be less than half of the depth.")
            'End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the HSSR section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateHSSRSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim thicknessDesign# = CDbl(inputProperties.Item(IJUAHSS_tdes))

            'Validating HSSR section for depth, width, thickness design
            If Not IsValidHSSR(depth, width, thicknessDesign) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_DEPTH_WIDTH_TWICE_THICKNESS_ZERO, _
                "Depth and width should be greater than twice the thickness design, and thickness design must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the PIPE section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidatePIPESectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim thicknessDesign# = CDbl(inputProperties.Item(IJUAHSS_tdes))

            'Validating PIPE section for depth and thickness design
            If Not IsValidPIPE(depth, thicknessDesign) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_DEPTH_TWICE_THICKNESS_ZERO, _
                "Depth should be greater than twice the thickness design, and thickness design must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the CS section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateCSSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))

            If Not IsValidCS(depth) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_VALIDATE_CSSECTION, _
                "Depth should always be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the RS section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateRSSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim width# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Width))

            If Not IsValidRS(depth, width) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_VALIDATE_RSSECTION, _
                "Depth and width must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUI section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUISectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))

            'Validating input properties like depth, top flange thickness and width, web thickness
            If Not IsValidBUI(depth, topFlangeThickness, topFlangeWidth, webThickness) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUI, _
                "Input properties should satisfy the following conditions : 1.Depth should be greater than top flange thickness. 2.Top flange width should be greater than web thickness. 3. Web thickness must be greater than zero 4.Top flange thickness must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUIUE section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUIUESectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))

            'Validating input properties like depth, topFlangeThickness, bottomFlangeThickness, topFlangeWidth, bottomFlangeWidth, webThickness
            If Not IsValidBUIUE(depth, topFlangeThickness, bottomFlangeThickness, topFlangeWidth, bottomFlangeWidth, webThickness) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUIT_BUIUE_BUIHAUNCH_BUC_BUBOXFM, _
                "Input properties should satisfy the following conditions: 1. Depth should be greater than the sum of the top flange and bottom flange thickness; 2. Top flange width and bottom flange width should be greater than the web thickness; 3. Web thickness must be greater than zero; 4. Top flange and bottom flange thickness must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUITapWeb section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUITapWebSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim depthStart# = CDbl(inputProperties.Item(IUABuiltUpITaperWeb_DepthStart))
            Dim depthEnd# = CDbl(inputProperties.Item(IUABuiltUpITaperWeb_DepthEnd))

            'The value of depth will be the maximum value of depth start or depth end 
            Dim depth# = Max(depthStart, depthEnd)

            'The value of width will be the maximum value of top flange width or bottom flange width 
            Dim width# = Max(topFlangeWidth, bottomFlangeWidth)

            'Validating input properties like depth, topFlangeThickness, bottomFlangeThickness, topFlangeWidth, bottomFlangeWidth, webThickness
            If Not IsValidBUITapWeb(depth, topFlangeThickness, bottomFlangeThickness, topFlangeWidth, bottomFlangeWidth, webThickness) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUIT_BUIUE_BUIHAUNCH_BUC_BUBOXFM, _
                "Input properties should satisfy the following conditions: 1. Depth should be greater than the sum of the top flange and bottom flange thickness; 2. Top flange width and bottom flange width should be greater than the web thickness; 3. Web thickness must be greater than zero; 4. Top flange and bottom flange thickness must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUIHaunch section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUIHaunchSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim depthStart# = CDbl(inputProperties.Item(IUABuiltUpIHaunch_DepthStart))
            Dim depthEnd# = CDbl(inputProperties.Item(IUABuiltUpIHaunch_DepthEnd))
            Dim lengthStart# = CDbl(inputProperties.Item(IUABuiltUpIHaunch_LengthStart))
            Dim lengthEnd# = CDbl(inputProperties.Item(IUABuiltUpIHaunch_LengthEnd))
            Dim depthHaunch# = CDbl(inputProperties.Item(IUABuiltUpIHaunch_DepthHaunch))

            'The value of depth can be the maximum value of depth start or depth end 
            Dim depth# = Max(depthEnd, depthStart)

            'The value of width can be the maximum value of top flange width or bottom flange width
            Dim width# = Max(topFlangeWidth, bottomFlangeWidth)

            'Validating input properties like depth, topFlangeThickness, bottomFlangeThickness, topFlangeWidth, bottomFlangeWidth, webThickness, lengthStart, lengthEnd
            If Not IsValidBUIHaunch(depth, topFlangeThickness, bottomFlangeThickness, topFlangeWidth, bottomFlangeWidth, webThickness, lengthStart, lengthEnd, depthStart, depthEnd, depthHaunch) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUIT_BUIUE_BUIHAUNCH_BUC_BUBOXFM, _
                "Input properties should satisfy the following conditions: 1. Depth should be greater than the sum of the top flange and bottom flange thickness; 2. Top flange width and bottom flange width should be greater than the web thickness; 3. Web thickness must be greater than zero; 4. Top flange and bottom flange thickness must be greater than zero. 5.Depth Haunch must be greater than zero and it should not be equal to Depth Start and Depth End. 6. Depth Start and Depth End must be greater than zero  ")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUL section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBULSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim offSetWeb# = CDbl(inputProperties.Item(IUABuiltUpL_OffsetWeb))

            'Validating input properties like depth, bottomFlangeThickness, bottomFlangeWidth, webThickness
            If Not IsValidBUL(depth, bottomFlangeThickness, bottomFlangeWidth, webThickness) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUIT_BUIUE_BUIHAUNCH_BUC_BUBOXFM, _
                "Input properties should satisfy the following conditions: 1. Depth should be greater than the sum of the top flange and bottom flange thickness; 2. Top flange width and bottom flange width should be greater than the web thickness; 3. Web thickness must be greater than zero; 4. Top flange and bottom flange thickness must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUC section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUCSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim offSetWebTop# = CDbl(inputProperties.Item(IUABuiltUpC_OffsetWebTop))
            Dim offSetWebBottom# = CDbl(inputProperties.Item(IUABuiltUpC_OffsetWebBot))

            'Validating input properties like depth, topFlangeThickness, bottomFlangeThickness, topFlangeWidth, bottomFlangeWidth, webThickness
            If Not IsValidBUBoxAndBUC(depth, topFlangeThickness, bottomFlangeThickness, topFlangeWidth, bottomFlangeWidth, webThickness) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUIT_BUIUE_BUIHAUNCH_BUC_BUBOXFM, _
                "Input properties should satisfy the following conditions: 1. Depth should be greater than the sum of the top flange and bottom flange thickness; 2. Top flange width and bottom flange width should be greater than the web thickness; 3. Web thickness must be greater than zero; 4. Top flange and bottom flange thickness must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUBoxFM section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUBoxFMSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim bottomFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeWidth))
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim bottomFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpBottomFlange_BottomFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim offsetLeftWebTop# = CDbl(inputProperties.Item(IUABUBoxFlangeMajor_OffsetLeftWebTop))
            Dim offsetRightWebTop# = CDbl(inputProperties.Item(IUABUBoxFlangeMajor_OffsetRightWebTop))
            Dim offsetLeftWebBottom# = CDbl(inputProperties.Item(IUABUBoxFlangeMajor_OffsetLeftWebBot))
            Dim offsetRightWebBottom# = CDbl(inputProperties.Item(IUABUBoxFlangeMajor_OffsetRightWebBot))

            'Validating input properties like depth, topFlangeThickness, bottomFlangeThickness, topFlangeWidth, bottomFlangeWidth, webThickness
            If Not IsValidBUBoxAndBUC(depth, topFlangeThickness, bottomFlangeThickness, topFlangeWidth, bottomFlangeWidth, 0.5 * webThickness) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUIT_BUIUE_BUIHAUNCH_BUC_BUBOXFM, _
                "Input properties should satisfy the following conditions: 1. Depth should be greater than the sum of the top flange and bottom flange thickness; 2. Top flange width and bottom flange width should be greater than the web thickness; 3. Web thickness must be greater than zero; 4. Top flange and bottom flange thickness must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUTube section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUTubeSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim tubeDiameter# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeDiameter))
            Dim tubeThickness# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeThickness))

            'Validating input properties like tubeDiameter, tubeThickness
            If Not IsValidBUTube(tubeDiameter, tubeThickness) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUTUBE, _
                "Input properties should satisfy the following conditions: 1. Tube diameter should be greater than twice the tube thickness; 2. Tube thickness must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUCone section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUConeSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim coneThickness# = CDbl(inputProperties.Item(IUABuiltUpCone_ConeThickness))
            Dim diameterStart# = CDbl(inputProperties.Item(IUABuiltUpCone_DiameterStart))
            Dim diameterEnd# = CDbl(inputProperties.Item(IUABuiltUpCone_DiameterEnd))

            'Validating input properties like diameterStart, diameterEnd, coneThickness
            If Not IsValidBUCone(diameterStart, diameterEnd, coneThickness) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUCONE, _
                "Input properties should satisfy the following conditions: 1. Diameter start and diameter end should be greater than twice the cone thickness; 2. Diameter start should not be equal to diameter end; 3. Cone thickness must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUCan section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUCanSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim tubeDiameter# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeDiameter))
            Dim tubeThickness# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeThickness))
            Dim diameterStart# = CDbl(inputProperties.Item(IUABuiltUpCan_DiameterStart))
            Dim diameterEnd# = CDbl(inputProperties.Item(IUABuiltUpCan_DiameterEnd))
            Dim cone1Thickness# = CDbl(inputProperties.Item(IUABuiltUpCone1_Cone1Thickness))
            Dim cone2Thickness# = CDbl(inputProperties.Item(IUABuiltUpCone2_Cone2Thickness))

            'Validating input properties like tubeDiameter, tubeThickness
            If Not IsValidBUCan(diameterStart, diameterEnd, tubeDiameter, tubeThickness, cone1Thickness, cone2Thickness) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUCAN, _
                "Input properties should satisfy the following conditions: 1. Tube diameter should be greater than twice the tube thickness; 2. Diameter start should be greater than twice the cone thickness; 3. Diameter end should be greater than twice the cone thickness;  4. Tube diameter should be greater than the diameter start; 5. Tube diameter should be greater than the diameter end; 6. Tube thickness, cone1Thickness, and cone2Thickness must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUEndCan section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUEndCanSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim tubeDiameter# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeDiameter))
            Dim tubeThickness# = CDbl(inputProperties.Item(IUABuiltUpTube_TubeThickness))
            Dim coneDiameter# = CDbl(inputProperties.Item(IUABUILTUPENDCAN_DIAMETERCONE))

            If Not IsValidEndCan(tubeDiameter, tubeThickness, coneDiameter) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUENDCAN, _
                "Input properties should satisfy the following conditions: 1. Tube diameter should be greater than cone diameter and should be greater than twice the tube thickness; 2. Tube thickness should be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUFlat section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUFlatSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))

            If Not IsValidBUFlat(depth, webThickness) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_VALIDATE_BUFLAT, _
                "Both depth and web thickness must be greater than zero.")
            End If

            Return String.Empty

        End Function

        ''' <summary>
        ''' Validates the BUTee section properties.
        ''' </summary>
        ''' <param name="inputProperties">The input properties.</param>
        ''' <returns>Error message</returns>
        Private Function ValidateBUTeeSectionProperties(ByVal inputProperties As Dictionary(Of String, Object)) As String

            'Following are the required inputs which are necessary in validation of properties
            Dim topFlangeWidth# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeWidth))
            Dim topFlangeThickness# = CDbl(inputProperties.Item(IUABuiltUpTopFlange_TopFlangeThickness))
            Dim webThickness# = CDbl(inputProperties.Item(IUABuiltUpWeb_WebThickness))
            Dim depth# = CDbl(inputProperties.Item(IStructCrossSectionDimensions_Depth))

            'Validating input properties like depth, topFlangeWidth, topFlangeThickness, webThickness
            If Not IsValidBUTee(depth, topFlangeWidth, topFlangeThickness, webThickness) Then
                Return GetString(SectionLibraryCalculatorResourceIDs.MSG_INPUT_VALIDATION_BUENDCAN, _
                "Input properties should satisfy the following conditions: 1. Web thickness should be greater than zero and should always be less than the flange width; 2. Flange thickness should always be less than half of the depth; 3. Top flange thickness must be greater than zero.")
            End If

            Return String.Empty

        End Function

#End Region

    End Class
End Namespace
