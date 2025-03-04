//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   TypeL.cs
//   Marine_Assy,Ingr.SP3D.Content.Support.Rules.TypeL
//   Author       :Vijay
//   Creation Date:02.Aug.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  02.Aug.2013     Vijay     CR-CP-224470 Convert HS_Marine_Assy to C# .Net
//  04.Dec.2013    Rajeswari  DI-CP-241804 Modified the code as part of hardening
//  11.Dec.2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//  25.Apr.2016      PVK      TR-CP-292881	Resolve the issues found in Marine Assemblies
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System.Linq;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class TypeL : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string VERTICALSECTION = "VERTICALSECTION";
        private const string HORIZONTALSECTION = "HORIZONTALSECTION";
        private const string LEFTPAD = "LEFTPAD";
        private const string HORIZONTALSECTIONPAD = "HORIZONTALSECTIONPAD";
        private const string BRACEPAD = "BRACEPAD";
        private const string BRACESECTION = "BRACESECTION";
        private string sectionSize = string.Empty, supportingType = string.Empty;
        private int sectionType = 0, overhangOption = 0, connection = 0;
        private double uBoltOffset = 0, structAngle = 0, braceHorizontalOffset = 0, braceVerticalOffset = 0, cornerOverhang = 0;
        private double PI = 0;
        private bool includeLeftPad, includeBracePad, includeBrace, frameSnipToFlge, frameSnipToWeb, isVerticalRoute, isVerticalStruct, slopedSteelY, slopedRoute, slopedSteel, isCornerSupport, isCornerOverhangExists;
        static int routeCount = 0;
        int[] uBoltCount = new int[routeCount];
        bool[] isUboltExist = new bool[routeCount];
        string[] uBoltPart1 = new string[routeCount];
        string[] uBolt = new string[routeCount];
        double[] routeAngle;
        MarineAssemblyServices.PADProperties padProperties;

        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    PI = Math.Round(Math.Atan(1) * 4, 3);
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    //Get the attributes from assembly
                    sectionType = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnSection", "SectionType")).PropValue;
                    overhangOption = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnOHOption", "OverhangOpt")).PropValue;
                    includeLeftPad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnLeftPad", "IncludeLeftPad")).PropValue;
                    bool uBoltOption = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnPAttachOpt", "PipeAttachOption")).PropValue;
                    bool uboltOffsetFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnPAttachOff", "OffsetFrmRule")).PropValue;
                    uBoltOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnPAttachOff", "Offset")).PropValue;
                    connection = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnFrmConn", "Connection")).PropValue;
                    includeBracePad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnBracePad", "IncludeBrPad")).PropValue;
                    includeBrace = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnBrace", "IncludeBrace")).PropValue;
                    isCornerSupport = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnCornerSupp", "IsCornerSupport")).PropValue;
                    frameSnipToFlge = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FrameSniptoFlange")).PropValue;
                    frameSnipToWeb = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "FrameSniptoWeb")).PropValue;

                    //Check the Corner Overhang Property in the catalog
                    if (support.SupportsInterface("IJOAhsMrnCornerOH"))
                        isCornerOverhangExists = true;
                    else
                        isCornerOverhangExists = false;

                    if (isCornerOverhangExists == false)
                    {
                        if (overhangOption == 3)       //User defined
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARINGMESSAGE", "Parts" + ": " + "WARNING: " + "Only 'By Catalog Rule' overhang option is applicable for Corner Support. Resetting the value. Bulkload the new HS_Marine_Assy.xls to use 'Corner Overhang' property when 'User defined' option is set.", "", "TypeL.cs", 96);
                            overhangOption = 1;
                            support.SetPropertyValue(1, "IJOAhsMrnOHOption", "OverhangOpt");
                        }
                    }
                    else
                    {
                        try
                        { cornerOverhang = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnCornerOH", "CornerOverhang")).PropValue; }
                        catch
                        { cornerOverhang = 0; }
                    }

                    routeCount = SupportHelper.SupportedObjects.Count;
                    if (routeCount > 10)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Number of Pipes should be less than or equal to 10", "", "TypeL.cs", 118);
                        return null;
                    }
                    //Warning message for flatbar if wrapping or welded connection selected
                    if (sectionType == 4)
                    {
                        if (connection == 2 || connection == 3)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Wrapping and Welded connections are not applicable for Flatbar Section", "", "TypeL.cs", 126);
                        }
                    }

                    //Get the Section Size
                    sectionSize = MarineAssemblyServices.GetSectionSize(this);
                    PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnSection", "SectionType");
                    //Set the Snip Angle
                    if (frameSnipToFlge == true || frameSnipToWeb == true)
                    {
                        if (sectionSizeCodelist.PropValue != 1)     //Snipped Steel is only applicable for L Section
                        {
                            frameSnipToFlge = false;
                            frameSnipToWeb = false;
                        }
                    }

                    //Get the UBolt
                    string[] uBoltPart = new string[routeCount];
                    int[] pipeAttachment = new int[routeCount];

                    Array.Resize(ref uBoltPart, routeCount);
                    uBoltPart = MarineAssemblyServices.GetUboltPart(this);

                    string tempUBoltType = string.Empty;
                    Array.Resize(ref uBoltCount, routeCount);
                    Array.Resize(ref isUboltExist, routeCount);

                    Array.Resize(ref routeAngle, routeCount);
                    Array.Resize(ref uBoltPart1, routeCount);
                    Array.Resize(ref pipeAttachment, routeCount);
                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        tempUBoltType = "U" + routeIndex;
                        pipeAttachment[routeIndex - 1] = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnPAttach", tempUBoltType)).PropValue;
                        if (routeIndex == 1)
                            uBoltCount[routeIndex - 1] = 1 + routeIndex;     //Initialize to 1
                        else
                            uBoltCount[routeIndex - 1] = uBoltCount[routeIndex - 1 - 1];
                        if (uBoltOption == false)      //User defined
                        {
                            if (pipeAttachment[routeIndex - 1] == 3)       //None
                                isUboltExist[routeIndex - 1] = false;
                            else
                            {
                                uBoltCount[routeIndex - 1] = uBoltCount[routeIndex - 1] + 1;
                                isUboltExist[routeIndex - 1] = true;
                            }
                        }
                        else        //Default Ubolt is True
                        {
                            uBoltCount[routeIndex - 1] = uBoltCount[routeIndex - 1] + 1;
                            isUboltExist[routeIndex - 1] = true;
                        }
                    }

                    string[] temp = new string[1];
                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        if (isUboltExist[routeIndex - 1] == true)
                        {
                            temp = uBoltPart[routeIndex - 1].Split('-');
                            uBoltPart1[routeIndex - 1] = temp[0] + "-";
                        }
                    }
                    
                    routeAngle = MarineAssemblyServices.GetRoutAngleAndIsRouteVertical(this, routeCount, out isVerticalRoute);
                    if (routeCount > 1)
                    {
                        for (int routeIndex = 1; routeIndex <= routeCount - 1; routeIndex++)
                        {
                            if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle[routeIndex - 1]), Math.Round(routeAngle[(routeIndex - 1) + 1])) == false)
                            {
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe slopes are not equal. Please check", "", "TypeL.cs", 200);
                                return null;
                            }
                        }
                    }

                    //Check to see what they are connecting to
                    //Checking the angles of Stucture X and Stucture Y axes with Global Z axis
                    supportingType = MarineAssemblyServices.GetSupportingTypes(this);

                    if (supportingType == "Slab")
                    {
                        structAngle = Math.Round(RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, OrientationAlong.Global_Z), 2);
                        if (Symbols.HgrCompareDoubleService.cmpdbl(structAngle, Math.Round((PI / 2), 2)) == true)
                        {
                            structAngle = Math.Round(RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Y, OrientationAlong.Global_Z), 2);
                            slopedSteelY = true;
                        }
                    }
                    else
                        structAngle = Math.Round(RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, OrientationAlong.Global_Z), 2);

                    //when structure is vertical
                    if (structAngle < PI / 4)
                        isVerticalStruct = true;
                    else if (structAngle > 3 * PI / 4)
                    {
                        isVerticalStruct = true;
                        structAngle = PI - structAngle;
                    }
                    else
                    {
                        isVerticalStruct = false;
                        structAngle = PI / 2 - structAngle;
                    }

                    if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle[0]),0)== false)
                        slopedRoute = true;
                    else if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(structAngle, 3), 0) == false)
                        slopedSteel = true;

                    if (isCornerSupport == true)
                    {
                        slopedRoute = false;
                        slopedSteel = false;
                    }
                    //Get Pad part and Dimensions

                    string sectionCode, steelStd;
                    padProperties = MarineAssemblyServices.GetPadPartAndDimensions(this, sectionSize, out sectionCode, out steelStd);

                    //get Brace Offset
                    MarineAssemblyServices.GetBraceOffset(this, sectionSize, ref braceHorizontalOffset, ref braceVerticalOffset);
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                    //Get the UboltOffset
                    if (uboltOffsetFromRule == true)
                    {
                        PartClass hsMarineServiceDim = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_PAttachOffset");
                        IEnumerable<BusinessObject> hsMarineServiceDimPart = hsMarineServiceDim.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        hsMarineServiceDimPart = hsMarineServiceDimPart.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvPAttachOff", "SectionSize")).PropValue == sectionCode));
                        if (hsMarineServiceDimPart.Count() > 0)
                            uBoltOffset = (double)((PropertyValueDouble)hsMarineServiceDimPart.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvPAttachOff", "PipeAttachOffset")).PropValue;
                    }

                    //Create the list of parts
                    string steelStandard = MarineAssemblyServices.GetSteelStandard(this, steelStd);

                    parts.Add(new PartInfo(HORIZONTALSECTION, sectionSize + " " + steelStandard));
                    parts.Add(new PartInfo(VERTICALSECTION, sectionSize + " " + steelStandard));

                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        if (isUboltExist[routeIndex - 1] == true)
                        {
                            Array.Resize(ref uBolt, parts.Count + 2);
                            uBolt[parts.Count + 1] = "sUbolt_" + routeIndex + 1;
                            parts.Add(new PartInfo(uBolt[uBoltCount[routeIndex - 1]], uBoltPart[routeIndex - 1]));
                        }
                    }

                    if (includeLeftPad == true)
                    {
                        parts.Add(new PartInfo(LEFTPAD, padProperties.padPart));

                        if (isCornerSupport == true)
                            parts.Add(new PartInfo(HORIZONTALSECTIONPAD, padProperties.padPart));
                    }

                    if (includeBrace == true)
                    {
                        parts.Add(new PartInfo(BRACESECTION, sectionSize + " " + steelStandard));

                        if (includeBracePad == true)
                            parts.Add(new PartInfo(BRACEPAD, padProperties.padPart));
                    }
                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 4;
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                double boundingBoxWidth = 0, boundingBoxHeight = 0, leftPipeDiameter = 0, rightPipeDiameter = 0, overhangLeft = 0, overhangRight = 0, cutbackangle = 0;
                MarineAssemblyServices.GetBoundingBoxDimensionsAndPipeDiameter(this, routeCount, ref boundingBoxWidth, ref boundingBoxHeight, ref leftPipeDiameter, ref rightPipeDiameter);

                //Auto Dimensioning of Supports
                int textIndex = 0;
                for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                {
                    if (isUboltExist[routeIndex - 1] == true)
                    {
                        textIndex = textIndex + 1;
                        MarineAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[uBolt[uBoltCount[routeIndex - 1]]], "Route");
                    }
                }

                textIndex = textIndex + 1;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[HORIZONTALSECTION], "BeginCap");

                textIndex = textIndex + 1;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[HORIZONTALSECTION], "EndCap");

                PropertyValueCodelist braceCodelistcp2 = (PropertyValueCodelist)componentDictionary[VERTICALSECTION].GetPropertyValue("IJOAhsSteelCP", "CP2");
                braceCodelistcp2.PropValue = 1;
                componentDictionary[VERTICALSECTION].SetPropertyValue(braceCodelistcp2.PropValue, "IJOAhsSteelCP", "CP2");
                PropertyValueCodelist braceCodelistcp4 = (PropertyValueCodelist)componentDictionary[VERTICALSECTION].GetPropertyValue("IJOAhsSteelCP", "CP4");
                braceCodelistcp4.PropValue = 3;
                componentDictionary[VERTICALSECTION].SetPropertyValue(braceCodelistcp4.PropValue, "IJOAhsSteelCP", "CP4"); ;

                string strVerticalSectionPort = string.Empty;
                if (structAngle <= 0)
                    strVerticalSectionPort = "EndCap";
                else
                    strVerticalSectionPort = "EndFace";

                textIndex = textIndex + 1;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[VERTICALSECTION], strVerticalSectionPort);

                //Get Section Structure dimensions
                BusinessObject sectionPart = componentDictionary[HORIZONTALSECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                //Get SteelWidth and SteelDepth 
                double sectionWidth = 0.0, sectionDepth = 0.0, sectionThickness = 0.0;
                if (sectionType == 4)
                {
                    sectionWidth = crossSection.Width;
                    sectionDepth = crossSection.Depth;
                }
                else
                {
                    sectionWidth = crossSection.Width;
                    sectionDepth = crossSection.Depth;
                    sectionThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                }

                //Check Width Offset and Height Offset values
                double widthOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnWrappingOff", "WidthOffset")).PropValue;
                double heightOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnWrappingOff", "HeightOffset")).PropValue;
                if (widthOffset > sectionDepth)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING" + "Frame offsets should not be more than frame depth, resetting the values", "", "TypeL.cs", 385);
                    widthOffset = 0;
                    heightOffset = 0;
                    support.SetPropertyValue(0.0, "IJOAhsMrnWrappingOff", "WidthOffset");
                    support.SetPropertyValue(0.0, "IJOAhsMrnWrappingOff", "HeightOffset");
                    return;
                }

                //Set the Route Connection Value for Corner Support
                int routeConnectionValue = Configuration;

                double cornerPipeDiameter = 0, verticalSectionEndOverLength = 0;
                bool isLeftPipe = false;
                if (isCornerSupport == true)
                {
                    double routeStructConfigAngle = RefPortHelper.PortConfigurationAngle("Route", "Struct_2", PortAxisType.Y);
                    if (!(sectionType == 4))
                    {
                        if (routeStructConfigAngle < 0.001 && routeStructConfigAngle > -0.001)
                        {
                            if (Configuration == 2 || routeConnectionValue == 4)
                                routeConnectionValue = 3;
                            else
                                routeConnectionValue = 1;
                        }
                        else
                        {
                            if (routeConnectionValue == 1 || routeConnectionValue == 3)
                                routeConnectionValue = 2;
                            else
                                routeConnectionValue = 4;
                        }
                    }
                    else
                    {
                        if (routeStructConfigAngle < 0.001 && routeStructConfigAngle > -0.001)
                        {
                            if (routeConnectionValue == 3)
                                routeConnectionValue = 1;
                            else if (routeConnectionValue == 4)
                                routeConnectionValue = 2;
                        }
                        else
                        {
                            if (routeConnectionValue == 1)
                                routeConnectionValue = 3;
                            else if (routeConnectionValue == 2)
                                routeConnectionValue = 4;
                        }
                    }

                    //Set the Corner pipe Dia
                    if (sectionType == 4)
                    {
                        if (routeConnectionValue == 1 || routeConnectionValue == 2)
                        {
                            cornerPipeDiameter = leftPipeDiameter;
                            isLeftPipe = true;
                        }
                        else
                        {
                            cornerPipeDiameter = rightPipeDiameter;
                            isLeftPipe = false;
                        }
                    }
                    else
                    {
                        if (routeConnectionValue == 1 || routeConnectionValue == 3)
                        {
                            cornerPipeDiameter = leftPipeDiameter;
                            isLeftPipe = true;
                        }
                        else
                        {
                            cornerPipeDiameter = rightPipeDiameter;
                            isLeftPipe = false;
                        }
                    }
                }
                //========================================
                // Set Values of Part Occurance Attributes
                //========================================
                string[] arrayOfKeys = new string[2];
                arrayOfKeys[0] = HORIZONTALSECTION;
                arrayOfKeys[1] = VERTICALSECTION;
                if (includeBrace == true)
                {
                    Array.Resize(ref arrayOfKeys, 3);
                    arrayOfKeys[2] = BRACESECTION;
                }

                //get Codelisted Property Values for Hgr Beam
                MarineAssemblyServices.BeamCLProperties beamCLProperties = MarineAssemblyServices.GetBeamCLProperties(componentDictionary);
                //Intialize the Hgr beam properties
                MarineAssemblyServices.IntializeBeamProperties(componentDictionary, support, beamCLProperties, arrayOfKeys);

                //For L section Initialize Snipped Properties
                if (frameSnipToFlge == true || frameSnipToWeb == true)
                {
                    string[] arrayOfKeys1 = new string[1];
                    arrayOfKeys1[0] = HORIZONTALSECTION;

                    if (connection == 3)
                    {
                        Array.Resize(ref arrayOfKeys1, 2);
                        arrayOfKeys1[0] = HORIZONTALSECTION;
                        arrayOfKeys1[1] = VERTICALSECTION;
                    }
                    for (int index = 0; index < arrayOfKeys1.Length; index++)
                    {
                        componentDictionary[arrayOfKeys1[index]].SetPropertyValue(beamCLProperties.HgrBeamType2.Value, "IJOAHsHgrBeamType", "HgrBeamType");
                        componentDictionary[arrayOfKeys1[index]].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackBeginAngle1");
                        componentDictionary[arrayOfKeys1[index]].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackBeginAngle2");
                        componentDictionary[arrayOfKeys1[index]].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackEndAngle1");
                        componentDictionary[arrayOfKeys1[index]].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackEndAngle2");
                        componentDictionary[arrayOfKeys1[index]].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "BeginOffsetAlongFlange");
                        componentDictionary[arrayOfKeys1[index]].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "BeginOffsetAlongWeb");
                        componentDictionary[arrayOfKeys1[index]].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "EndOffsetAlongFlange");
                        componentDictionary[arrayOfKeys1[index]].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "EndOffsetAlongWeb");
                    }
                }
                //set the overlength for the vertcal section
                PropertyValueCodelist fromOrientCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnFrmOrient", "FrmOrient");
                string fromOrient = fromOrientCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(fromOrientCodelist.PropValue).ShortDisplayName;
                if (slopedSteel == true && SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (sectionType == 4)
                        {
                            if (routeConnectionValue == 1 || routeConnectionValue == 2)
                                componentDictionary[VERTICALSECTION].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackEndAngle");
                            else
                                componentDictionary[VERTICALSECTION].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackEndAngle");
                        }
                        else
                        {
                            if (routeConnectionValue == 1)
                            {
                                if (connection == 1)
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackEndAngle");
                                }
                                else
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    if (connection == 2)
                                        componentDictionary[VERTICALSECTION].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                    else
                                        componentDictionary[VERTICALSECTION].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                }
                            }
                            else if (routeConnectionValue == 2)
                            {
                                if (connection == 1)
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                }
                                else
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    if (connection == 2)
                                        componentDictionary[VERTICALSECTION].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackEndAngle");
                                    else
                                        componentDictionary[VERTICALSECTION].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackEndAngle");
                                }
                            }
                            else if (routeConnectionValue == 3)
                            {
                                if (connection == 1)
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                }
                                else
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    if (connection == 2)
                                        componentDictionary[VERTICALSECTION].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackEndAngle");
                                    else
                                        componentDictionary[VERTICALSECTION].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackEndAngle");
                                }
                            }
                            else if (routeConnectionValue == 4)
                            {
                                if (connection == 1)
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackEndAngle");
                                }
                                else
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    if (connection == 2)
                                        componentDictionary[VERTICALSECTION].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                    else
                                        componentDictionary[VERTICALSECTION].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                }
                            }
                        }
                    }
                }

                //Get the largest Pipe Dia
                double[] pipeDiameter = new double[routeCount];
                for (int index = 1; index <= routeCount; index++)
                {
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(index);
                    pipeDiameter[index - 1] = pipeInfo.OutsideDiameter;
                }

                //For UBolt
                string uBoltType = string.Empty;
                double[] H = new double[routeCount];

                //Set attribute values for the U-Bolt
                int[] uBoltTypes = new int[routeCount];
                string tempUBoltType = string.Empty;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                {
                    if (isUboltExist[routeIndex - 1] == true)
                    {
                        tempUBoltType = "U" + routeIndex;

                        PartClass hsMarineServiceDim = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_PAttachSel");
                        IEnumerable<BusinessObject> hsMarineServiceDimPart = hsMarineServiceDim.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        hsMarineServiceDimPart = hsMarineServiceDimPart.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvPAttachSel", "PipeAttachPart")).PropValue == uBoltPart1[routeIndex - 1]));
                        if (hsMarineServiceDimPart.Count() > 0)
                            uBoltTypes[routeIndex - 1] = (int)((PropertyValueCodelist)hsMarineServiceDimPart.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvPAttachSel", "PipeAttachType")).PropValue;

                        //Get the Ublot Dimensions
                        componentDictionary[uBolt[uBoltCount[routeIndex - 1]]].SetPropertyValue(pipeDiameter[routeIndex - 1], "IJOAhsPipeOD", "PipeOD");
                        if (uBoltTypes[routeIndex - 1] == 1)
                        {
                            if (sectionType == 4)
                                componentDictionary[uBolt[uBoltCount[routeIndex - 1]]].SetPropertyValue(sectionWidth, "IJOAhsSteelThickness", "SteelThickness");
                            else
                                componentDictionary[uBolt[uBoltCount[routeIndex - 1]]].SetPropertyValue(sectionThickness, "IJOAhsSteelThickness", "SteelThickness");
                        }
                        support.SetPropertyValue(uBoltTypes[routeIndex - 1], "IJOAhsMrnPAttach", tempUBoltType);
                    }
                }

                //Set attribute values for Pads
                string[] arrayOfPadKeys = new string[0];

                if (includeLeftPad == true)
                {
                    if (isCornerSupport == false)
                    {
                        Array.Resize(ref arrayOfPadKeys, 1);
                        verticalSectionEndOverLength = verticalSectionEndOverLength - padProperties.padThickness;
                        arrayOfPadKeys[0] = LEFTPAD;
                    }
                    else
                    {
                        Array.Resize(ref arrayOfPadKeys, 2);
                        verticalSectionEndOverLength = verticalSectionEndOverLength - padProperties.padThickness;
                        arrayOfPadKeys[0] = LEFTPAD;
                        arrayOfPadKeys[1] = HORIZONTALSECTIONPAD;
                    }
                }

                if (includeBracePad == true)
                {
                    if (includeLeftPad == true)
                    {
                        if (isCornerSupport == false)
                        {
                            Array.Resize(ref arrayOfPadKeys, 2);
                            arrayOfPadKeys[1] = BRACEPAD;
                        }
                        else
                        {
                            Array.Resize(ref arrayOfPadKeys, 3);
                            arrayOfPadKeys[2] = BRACEPAD;
                        }
                    }
                    else
                    {
                        Array.Resize(ref arrayOfPadKeys, 1);
                        arrayOfPadKeys[0] = BRACEPAD;
                    }
                }

                //Set the Pad properties
                MarineAssemblyServices.SetPadProperties(componentDictionary, support, padProperties, arrayOfPadKeys);

                //set the Include brace Pad
                if (includeBrace == false)
                    support.SetPropertyValue(false, "IJOAhsMrnBracePad", "IncludeBrPad");

                double sectionCutbackAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSecSnips", "SecCutbackAngle")).PropValue;
                bool frameCutbackToSection = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSecSnips", "FrameCutbacktoSection")).PropValue;
                if (frameCutbackToSection == true)
                    cutbackangle = sectionCutbackAngle;
                else
                    cutbackangle = 0;
                double flgSnipAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FlgSnipAngle")).PropValue;
                double webSnipAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "WebSnipAngle")).PropValue;
                double flgSnipOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FlgSnipOffset")).PropValue;
                double webSnipOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "WebSnipOffset")).PropValue;
                if (isCornerSupport == false)
                {
                    if (routeConnectionValue == 1 || routeConnectionValue == 4 || sectionType == 4)
                    {
                        if (frameSnipToFlge == false && frameSnipToWeb == false)
                        {
                            componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[HORIZONTALSECTION].SetPropertyValue(cutbackangle, "IJOAhsCutback", "CutbackBeginAngle");
                        }
                        else
                        {
                            if (frameSnipToFlge == true)
                            {
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(flgSnipAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle1");
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(flgSnipOffset, "IJOAhsCutbackOffset", "BeginOffsetAlongFlange");
                            }
                            if (frameSnipToWeb == true)
                            {
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(webSnipAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle2");
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(webSnipOffset, "IJOAhsCutbackOffset", "BeginOffsetAlongWeb");
                            }
                        }
                    }
                    else
                    {
                        if (frameSnipToFlge == false && frameSnipToWeb == false)
                        {
                            componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[HORIZONTALSECTION].SetPropertyValue(-cutbackangle, "IJOAhsCutback", "CutbackEndAngle");
                        }
                        else
                        {
                            if (frameSnipToFlge == true)
                            {
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(flgSnipAngle, "IJOAhsSnipedSteel", "CutbackEndAngle1");
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(flgSnipOffset, "IJOAhsCutbackOffset", "EndOffsetAlongFlange");
                            }
                            if (frameSnipToWeb == true)
                            {
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(webSnipAngle, "IJOAhsSnipedSteel", "CutbackEndAngle2");
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(webSnipOffset, "IJOAhsCutbackOffset", "EndOffsetAlongWeb");
                            }
                        }
                    }
                }
                //For UboltOffset
                support.SetPropertyValue(uBoltOffset, "IJOAhsMrnPAttachOff", "Offset");

                //Set Assembly Attributes
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;

                //Set Frame Orientation attribute
                CodelistItem codeList = metadataManager.GetCodelistInfo("hsMrnCLFrmOrient", "UDP").GetCodelistItem(fromOrient);
                support.SetPropertyValue(codeList.Value, "IJOAhsMrnFrmOrient", "FrmOrient");

                //Set Section Size properties
                codeList = metadataManager.GetCodelistInfo("hsMrnCLSecSize", "UDP").GetCodelistItem(sectionSize);
                support.SetPropertyValue(codeList.Value, "IJOAhsMrnSection", "SectionSize");

                //Set the overhang attributes
                if (isCornerSupport == true)
                {
                    if (overhangOption == 2)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Joints" + ": " + "WARNING: " + "Adjust to Supporting steel is not applicable for Corner Support. Resetting the Overhang Option to 'By Catalog Rule'", "", "TypeL.cs", 759);
                        overhangOption = 1;
                        support.SetPropertyValue(1, "IJOAhsMrnOHOption", "OverhangOpt");
                    }
                }

                double hangerOverHangLeft = 0, hangerOverHangRight = 0;
                Collection<object> hangerOverHangCollection;
                bool value = GenericHelper.GetDataByRule("hsMrnRLFrmOH", (BusinessObject)support, out hangerOverHangCollection);
                if (hangerOverHangCollection[0] == null)
                {
                    hangerOverHangLeft = (double)hangerOverHangCollection[1];
                    hangerOverHangRight = (double)hangerOverHangCollection[2];
                }
                else
                {
                    hangerOverHangLeft = (double)hangerOverHangCollection[0];
                    hangerOverHangRight = (double)hangerOverHangCollection[1];
                }

                if (overhangOption == 1)      //By Catalog Rule
                {
                    support.SetPropertyValue(hangerOverHangLeft, "IJOAhsMrnOHLeft", "OverhangLeft");
                    support.SetPropertyValue(hangerOverHangRight, "IJOAhsMrnOHRight", "OverhangRight");

                    overhangLeft = hangerOverHangLeft - rightPipeDiameter / 2;
                    overhangRight = hangerOverHangRight - leftPipeDiameter / 2;

                    //Set the corner overhang
                    if (isCornerSupport == true)
                    {
                        if (isCornerOverhangExists == true)
                        {
                            if (isLeftPipe == false)
                                support.SetPropertyValue(hangerOverHangLeft, "IJOAhsMrnCornerOH", "CornerOverhang");
                            else
                                support.SetPropertyValue(hangerOverHangRight, "IJOAhsMrnCornerOH", "CornerOverhang");
                        }
                    }
                }
                else if (overhangOption == 3)      //User Defined
                {
                    double userDefOverHangLeft = 0, userDefOverHangRight = 0, userDefOverHangCorner = 0;

                    if (isCornerSupport == false)
                    {
                        userDefOverHangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHLeft", "OverhangLeft")).PropValue;
                        userDefOverHangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHRight", "OverhangRight")).PropValue;

                        support.SetPropertyValue(userDefOverHangLeft, "IJOAhsMrnOHLeft", "OverhangLeft");
                        support.SetPropertyValue(userDefOverHangRight, "IJOAhsMrnOHRight", "OverhangRight");
                    }
                    else        //For Corner Support
                    {
                        if (isCornerOverhangExists == true)
                        {
                            userDefOverHangCorner = cornerOverhang;
                            userDefOverHangLeft = userDefOverHangCorner;
                            userDefOverHangRight = userDefOverHangCorner;
                            leftPipeDiameter = cornerPipeDiameter;
                            rightPipeDiameter = leftPipeDiameter;
                            support.SetPropertyValue(userDefOverHangCorner, "IJOAhsMrnCornerOH", "CornerOverhang");
                        }
                    }
                    overhangLeft = userDefOverHangLeft - rightPipeDiameter / 2;
                    overhangRight = userDefOverHangRight - leftPipeDiameter / 2;
                }

                //Set Length of horzontal member
                double perHorHighStruct2Dist, horLowStruct2Dist, horHighStruct2Dist, routeLowStruct2HorDist = 0, routeHighStruct2HorDist = 0, horSectionLength = 0, perHorLowStruct2Dist;

                double routeStructConfigAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    double tempDist = 0;
                    if (overhangOption == 1 || overhangOption == 3)   //By Catalog Ruleand user defined
                        tempDist = boundingBoxWidth;
                    else if (overhangOption == 2)     //adjust to structure
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Joints" + ": " + "WARNING: " + "Adjust to Supporting steel is applicable only when the support placement is By Point.", "", "TypeL.cs", 840);
                        return;
                    }

                    horSectionLength = tempDist + overhangLeft + overhangRight;
                    if (isCornerSupport == false)
                        componentDictionary[HORIZONTALSECTION].SetPropertyValue(horSectionLength, "IJUAHgrOccLength", "Length");
                }
                else        //for Place by point
                {
                    if ((Math.Abs(routeAngle[0]) < (0 + 0.01) && Math.Abs(routeAngle[0]) > (0 - 0.01)) || (Math.Abs(routeAngle[0]) < (PI + 0.01) && Math.Abs(routeAngle[0]) > (PI - 0.01)))
                    {
                        if (isVerticalRoute == true)
                        {
                            perHorLowStruct2Dist = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortDistanceType.Horizontal_Perpendicular);
                            perHorHighStruct2Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", "Structure", PortDistanceType.Horizontal_Perpendicular);
                            horLowStruct2Dist = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortDistanceType.Horizontal);
                            horHighStruct2Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", "Structure", PortDistanceType.Horizontal);

                            routeLowStruct2HorDist = Math.Sqrt(horLowStruct2Dist * horLowStruct2Dist - perHorLowStruct2Dist * perHorLowStruct2Dist);
                            routeHighStruct2HorDist = Math.Sqrt(horHighStruct2Dist * horHighStruct2Dist - perHorHighStruct2Dist * perHorHighStruct2Dist);
                        }
                        else
                        {
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortDistanceType.Horizontal);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", "Structure", PortDistanceType.Horizontal);
                        }
                    }
                    else
                    {
                        if ((Math.Abs(routeStructAngle) < (0 + 0.01) && Math.Abs(routeStructAngle) > (0 - 0.01)) || (Math.Abs(routeStructAngle) < (PI + 0.01) && Math.Abs(routeStructAngle) > (PI - 0.01)))
                        {
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortDistanceType.Vertical);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", "Structure", PortDistanceType.Vertical);
                        }
                        else if (Math.Abs(routeStructAngle) < (PI / 2 + 0.01) && Math.Abs(routeStructAngle) > (PI / 2 - 0.01))
                        {
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortDistanceType.Horizontal);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", "Structure", PortDistanceType.Horizontal);
                        }
                        else
                        {
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortDistanceType.Vertical);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", "Structure", PortDistanceType.Vertical);
                        }
                    }

                    if (supportingType == "Slab" || supportingType=="Wall")
                    {
                        if (overhangOption == 1 || overhangOption == 3)//By catalog OR User Defined
                            horSectionLength = boundingBoxWidth + overhangLeft + overhangRight;
                        else if (overhangOption == 2)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Joints" + ": " + "WARNING: " + "Adjust to Supporting steel is applicable only when the support placement is By Point and the supporting count is more than one.", "", "TypeL.cs", 893);
                            return;
                        }
                    }
                    else if (supportingType == "Steel")
                    {
                        if (overhangOption == 1 || overhangOption == 3)  //by catalog
                            horSectionLength = boundingBoxWidth + overhangLeft + overhangRight;
                        else if (overhangOption == 2)
                        {
                            if (routeStructConfigAngle1 > PI / 2)
                            {
                                overhangRight = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist) + sectionWidth / 2;
                                overhangLeft = hangerOverHangLeft - rightPipeDiameter / 2;

                                support.SetPropertyValue(overhangRight + leftPipeDiameter / 2, "IJOAhsMrnOHRight", "OverhangRight");
                                support.SetPropertyValue(hangerOverHangLeft, "IJOAhsMrnOHLeft", "OverhangLeft");
                            }
                            else
                            {
                                overhangLeft = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist) + sectionWidth / 2;
                                overhangRight = hangerOverHangRight - leftPipeDiameter / 2;

                                support.SetPropertyValue(hangerOverHangRight, "IJOAhsMrnOHRight", "OverhangRight");
                                support.SetPropertyValue(overhangLeft + rightPipeDiameter / 2, "IJOAhsMrnOHLeft", "OverhangLeft");
                            }
                            horSectionLength = boundingBoxWidth + overhangLeft + overhangRight;
                        }
                    }
                    else if (supportingType == "Steel-Slab" || supportingType == "Slab-Steel")
                    {
                        if (overhangOption == 1 || overhangOption == 3)
                            horSectionLength = boundingBoxWidth + overhangLeft + overhangRight;
                        else if (overhangOption == 2)      //adjust to structure
                        {
                            overhangRight = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist);
                            overhangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHLeft", "OverhangLeft")).PropValue;
                            horSectionLength = boundingBoxWidth + overhangLeft + overhangRight + sectionWidth;
                        }
                    }

                    if (isCornerSupport == false)
                        componentDictionary[HORIZONTALSECTION].SetPropertyValue(horSectionLength, "IJUAHgrOccLength", "Length");
                }
                //set the properties for connections

                if (connection == 1)
                {
                    widthOffset = 0;
                    heightOffset = 0;
                    if (!(sectionType == 4))
                    {
                        if (routeConnectionValue == 1 || routeConnectionValue == 4)
                        {
                            componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[VERTICALSECTION].SetPropertyValue(-PI / 4, "IJOAhsCutback", "CutbackBeginAngle");
                            if (frameSnipToFlge == false && frameSnipToWeb == false)
                            {
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(PI / 4, "IJOAhsCutback", "CutbackEndAngle");
                            }
                            else
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(PI / 4, "IJOAhsSnipedSteel", "CutbackEndAngle2");
                        }
                        else if (routeConnectionValue == 2 || routeConnectionValue == 3)
                        {
                            componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[VERTICALSECTION].SetPropertyValue(PI / 4, "IJOAhsCutback", "CutbackEndAngle");
                            if (frameSnipToFlge == false && frameSnipToWeb == false)
                            {
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(-PI / 4, "IJOAhsCutback", "CutbackBeginAngle");
                            }
                            else
                                componentDictionary[HORIZONTALSECTION].SetPropertyValue(PI / 4, "IJOAhsSnipedSteel", "CutbackBeginAngle2");
                        }
                    }
                }

                else if (connection == 2)     //For Wrapping Joint
                    componentDictionary[VERTICALSECTION].SetPropertyValue(PI, "IJOAhsBeginCap", "BeginCapRotZ");

                else if (connection == 3)      //For Welded joint
                {
                    widthOffset = 0;
                    heightOffset = 0;
                    if (!(sectionType == 4))
                    {
                        if (frameSnipToFlge == false && frameSnipToWeb == false)
                        {
                            if (routeConnectionValue == 1 || routeConnectionValue == 4)
                            {
                                componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                componentDictionary[VERTICALSECTION].SetPropertyValue(-cutbackangle, "IJOAhsCutback", "CutbackEndAngle");
                            }
                            else
                            {
                                componentDictionary[VERTICALSECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                componentDictionary[VERTICALSECTION].SetPropertyValue(cutbackangle, "IJOAhsCutback", "CutbackBeginAngle");
                            }
                        }
                        else
                        {
                            if (routeConnectionValue == 1 || routeConnectionValue == 4)
                            {
                                if (frameSnipToFlge == true)
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(flgSnipAngle, "IJOAhsSnipedSteel", "CutbackEndAngle1");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(flgSnipOffset, "IJOAhsCutbackOffset", "EndOffsetAlongFlange");
                                }
                                if (frameSnipToWeb == true)
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(webSnipAngle, "IJOAhsSnipedSteel", "CutbackEndAngle2");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(webSnipOffset, "IJOAhsCutbackOffset", "EndOffsetAlongWeb");
                                }
                            }
                            else
                            {
                                if (frameSnipToFlge == true)
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(flgSnipAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle1");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(flgSnipOffset, "IJOAhsCutbackOffset", "BeginOffsetAlongFlange");
                                }
                                if (frameSnipToWeb == true)
                                {
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(webSnipAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle2");
                                    componentDictionary[VERTICALSECTION].SetPropertyValue(webSnipOffset, "IJOAhsCutbackOffset", "BeginOffsetAlongWeb");
                                }
                            }
                        }
                    }
                }

                //=============
                //Create Joints
                //=============
                //Create a collection to hold the joints
                MarineAssemblyServices.ConfigIndex[] routeUboltConfigIndex = new MarineAssemblyServices.ConfigIndex[routeCount];
                MarineAssemblyServices.ConfigIndex horizontalVerticalConfigIndex = new MarineAssemblyServices.ConfigIndex(), routeHorizontalConfigIndex = new MarineAssemblyServices.ConfigIndex(), padConfigIndex = new MarineAssemblyServices.ConfigIndex();

                double horizontalRoutePlaneOffset = 0, horizontalRouteAxisOffset = 0, horizontalRouteOriginOffset = 0, horizonatlVertPlaneOffset1 = 0, horizontalVertAxisOffset1 = 0, boltRoutePlaneOffset = 0;
                string horizonatlSectionPort = string.Empty, verticalSectionPort = string.Empty, verticalSectionPadPort = string.Empty, verticalStructPort = string.Empty, bracePort = string.Empty, braceStructPort = string.Empty, braceCapPort = string.Empty, bracePadPort = string.Empty, routePort = string.Empty, horizontalSectionPort1 = string.Empty, padPort = string.Empty, frameLength = string.Empty;

                routePort = MarineAssemblyServices.GetRoutePort(this, support, slopedRoute, slopedSteel, isVerticalRoute, isVerticalStruct);

                if (includeBrace == true)
                {
                    if (sectionType == 4)
                    {
                        bracePort = "EndCap";
                        braceStructPort = "BeginCap";
                        braceCapPort = "BeginFace";
                        frameLength = "BeginOverLength";
                        bracePadPort = "Port2";
                    }
                    else
                    {
                        if (routeConnectionValue == 1 || routeConnectionValue == 4)
                        {
                            if (connection == 3 || connection == 2)
                            {
                                bracePort = "BeginCap";
                                braceStructPort = "EndCap";
                                braceCapPort = "EndFace";
                                frameLength = "EndOverLength";
                                bracePadPort = "Port1";
                            }
                            else
                            {
                                bracePort = "EndCap";
                                braceStructPort = "BeginCap";
                                braceCapPort = "BeginFace";
                                frameLength = "BeginOverLength";
                                bracePadPort = "Port2";
                            }
                        }
                        else if (routeConnectionValue == 2 || routeConnectionValue == 3)
                        {
                            if (connection == 3 || connection == 2)
                            {
                                bracePort = "EndCap";
                                braceStructPort = "BeginCap";
                                braceCapPort = "BeginFace";
                                frameLength = "BeginOverLength";
                                bracePadPort = "Port2";
                            }
                            else
                            {
                                bracePort = "BeginCap";
                                braceStructPort = "EndCap";
                                braceCapPort = "EndFace";
                                frameLength = "EndOverLength";
                                bracePadPort = "Port1";
                            }
                        }
                    }
                }

                if (routeConnectionValue == 1 || routeConnectionValue == 4)
                {
                    horizonatlSectionPort = "EndCap";
                    if (connection == 2 || connection == 3)
                    {
                        verticalSectionPort = "EndCap";
                        verticalSectionPadPort = "BeginFace";
                        verticalStructPort = "BeginCap";
                        padPort = "Port2";
                        horizonatlVertPlaneOffset1 = heightOffset;
                        horizontalVertAxisOffset1 = -widthOffset;
                    }
                    else
                    {
                        verticalSectionPort = "BeginCap";
                        verticalSectionPadPort = "EndFace";
                        verticalStructPort = "EndCap";
                        padPort = "Port1";
                        horizonatlVertPlaneOffset1 = 0;
                        horizontalVertAxisOffset1 = 0;
                    }
                }
                else if (routeConnectionValue == 2 || routeConnectionValue == 3)
                {
                    horizonatlSectionPort = "BeginCap";
                    if (connection == 2 || connection == 3)
                    {
                        verticalSectionPort = "BeginCap";
                        verticalSectionPadPort = "EndFace";
                        verticalStructPort = "EndCap";
                        padPort = "Port1";
                        horizonatlVertPlaneOffset1 = heightOffset;
                        horizontalVertAxisOffset1 = widthOffset;
                    }
                    else
                    {
                        verticalSectionPort = "EndCap";
                        verticalSectionPadPort = "BeginFace";
                        verticalStructPort = "BeginCap";
                        padPort = "Port2";
                        horizonatlVertPlaneOffset1 = 0;
                        horizontalVertAxisOffset1 = 0;
                    }
                }

                for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                {
                    if (isUboltExist[routeIndex - 1] == true)
                    {
                        if (!(sectionType == 4))
                            routeUboltConfigIndex[routeIndex - 1] = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeZ);
                        else
                            routeUboltConfigIndex[routeIndex - 1] = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ);
                    }
                }
                double shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnShoeHeight", "ShoeHeight")).PropValue;

                if (sectionType == 4)
                {
                    verticalSectionPort = "BeginCap";
                    verticalSectionPadPort = "EndFace";
                    verticalStructPort = "EndCap";
                    padPort = "Port1";
                    horizontalRoutePlaneOffset = sectionDepth / 2;
                    horizontalRouteOriginOffset = sectionWidth + shoeHeight;
                    horizontalSectionPort1 = "EndCap";
                    horizonatlSectionPort = "EndCap";
                    horizonatlVertPlaneOffset1 = sectionDepth;
                    boltRoutePlaneOffset = sectionDepth / 2 - uBoltOffset;
                    padConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.Y, Axis.NegativeY);

                    if (routeConnectionValue == 1 || routeConnectionValue == 2)
                    {
                        routeHorizontalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.YZ, Axis.X, Axis.NegativeZ);
                        horizontalVerticalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeZX, Axis.X, Axis.NegativeZ);
                        horizontalRouteAxisOffset = -overhangRight - boundingBoxWidth;
                    }
                    else if (routeConnectionValue == 3 || routeConnectionValue == 4)
                    {
                        routeHorizontalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeYZ, Axis.X, Axis.NegativeZ);
                        horizontalVerticalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeZX, Axis.X, Axis.NegativeZ);
                        horizontalRouteAxisOffset = -overhangRight;
                    }
                }
                else
                {
                    horizontalRouteOriginOffset = sectionWidth / 2;
                    boltRoutePlaneOffset = sectionWidth / 2 - uBoltOffset;
                    padConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeY);

                    if (routeConnectionValue == 1)
                    {
                        if (connection == 2 || connection == 3)
                            horizontalVerticalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X);
                        else
                            horizontalVerticalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);

                        routeHorizontalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        horizontalSectionPort1 = "EndCap";
                        horizontalRoutePlaneOffset = -shoeHeight;
                        horizontalRouteAxisOffset = -overhangRight - boundingBoxWidth;
                    }
                    else if (routeConnectionValue == 2)
                    {
                        if (connection == 2 || connection == 3)
                            horizontalVerticalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        else
                            horizontalVerticalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X);

                        horizontalSectionPort1 = "BeginCap";
                        routeHorizontalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        horizontalRoutePlaneOffset = -shoeHeight;
                        horizontalRouteAxisOffset = overhangLeft;
                    }
                    else if (routeConnectionValue == 3)
                    {
                        if (connection == 2 || connection == 3)
                            horizontalVerticalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        else
                            horizontalVerticalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X);

                        horizontalSectionPort1 = "BeginCap";
                        routeHorizontalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX);
                        horizontalRoutePlaneOffset = overhangRight + boundingBoxWidth;
                        horizontalRouteAxisOffset = -shoeHeight;
                    }
                    else if (routeConnectionValue == 4)
                    {
                        if (connection == 2 || connection == 3)
                            horizontalVerticalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X);
                        else
                            horizontalVerticalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);

                        routeHorizontalConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX);
                        horizontalSectionPort1 = "EndCap";
                        horizontalRoutePlaneOffset = -overhangLeft;
                        horizontalRouteAxisOffset = -shoeHeight;
                    }
                }

                //Add joint Between Horizontal Section and BoundingBox
                JointHelper.CreateRigidJoint(HORIZONTALSECTION, horizontalSectionPort1, "-1", routePort, routeHorizontalConfigIndex.A, routeHorizontalConfigIndex.B, routeHorizontalConfigIndex.C, routeHorizontalConfigIndex.D, horizontalRoutePlaneOffset, horizontalRouteAxisOffset, horizontalRouteOriginOffset);

                //Add Joint Between the Horizontal and Vertical Beams
                JointHelper.CreateRigidJoint(HORIZONTALSECTION, horizonatlSectionPort, VERTICALSECTION, verticalSectionPort, horizontalVerticalConfigIndex.A, horizontalVerticalConfigIndex.B, horizontalVerticalConfigIndex.C, horizontalVerticalConfigIndex.D, horizonatlVertPlaneOffset1, horizontalVertAxisOffset1, 0);

                //Add Joint Between the Ports of Vertical Beam
                JointHelper.CreatePrismaticJoint(VERTICALSECTION, "BeginCap", VERTICALSECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add Joint Between the Ports of Vertical Beam
                JointHelper.CreatePointOnPlaneJoint(VERTICALSECTION, verticalStructPort, "-1", "Structure", Plane.XY);

                string horizontalSectionStructPort = string.Empty, horizontalSectionPadPort = string.Empty, horizontalPadPort = string.Empty;
                if (isCornerSupport == true)
                {
                    if (routeConnectionValue == 1 || routeConnectionValue == 4 || sectionType == 4)
                    {
                        horizontalSectionStructPort = "BeginCap";
                        horizontalSectionPadPort = "BeginFace";
                        horizontalPadPort = "Port2";
                    }
                    else
                    {
                        horizontalSectionStructPort = "EndCap";
                        horizontalSectionPadPort = "EndFace";
                        horizontalPadPort = "Port1";
                    }
                    //Add Joint Between the Ports of Horizontal Beam
                    JointHelper.CreatePrismaticJoint(HORIZONTALSECTION, "BeginCap", HORIZONTALSECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    //Add Joint Between the Horizontal Beam and Structure 2
                    JointHelper.CreatePointOnPlaneJoint(HORIZONTALSECTION, horizontalSectionStructPort, "-1", "Struct_2", Plane.XY);
                }

                double braceAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnBraceAngle", "BraceAngle")).PropValue;
                if (includeBrace == true)
                {
                    double braceVerOffset = 0, braceHorOffset = 0, braceAxisOffset = 0, angleX = 0, angleY = 0, angleZ = 0, angleLength = 0;

                    if (sectionType == 4)
                        angleLength = sectionDepth / Math.Sin(braceAngle);
                    else
                        angleLength = sectionWidth / Math.Sin(braceAngle);

                    if (sectionType == 4)
                    {
                        componentDictionary[BRACESECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                        componentDictionary[BRACESECTION].SetPropertyValue(braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                        componentDictionary[BRACESECTION].SetPropertyValue(-(PI / 2 - braceAngle), "IJOAhsCutback", "CutbackEndAngle");
                        braceVerOffset = braceVerticalOffset + angleLength;
                        braceAxisOffset = braceHorizontalOffset;
                        braceHorOffset = -sectionWidth;
                        angleX = (PI - braceAngle);
                        angleZ = -PI;
                    }
                    else
                    {
                        if (connection == 2)
                            braceHorOffset = -braceHorizontalOffset;
                        else
                            braceHorOffset = braceHorizontalOffset;

                        if (routeConnectionValue == 1 || routeConnectionValue == 4)
                        {
                            if (connection == 2)
                            {
                                componentDictionary[BRACESECTION].SetPropertyValue(PI/2 - braceAngle, "IJOAhsCutback", "CutbackEndAngle");
                                componentDictionary[BRACESECTION].SetPropertyValue(-(braceAngle), "IJOAhsCutback", "CutbackBeginAngle");
                                braceVerOffset = -braceVerticalOffset - angleLength;
                                angleY = -(PI -braceAngle);
                                angleZ = 0;
                            }
                            else if (connection == 3)
                            {
                                componentDictionary[BRACESECTION].SetPropertyValue(braceAngle, "IJOAhsCutback", "CutbackEndAngle");
                                componentDictionary[BRACESECTION].SetPropertyValue(-(PI / 2 - braceAngle), "IJOAhsCutback", "CutbackBeginAngle");
                                braceVerOffset = -braceVerticalOffset - angleLength;
                                angleY = PI-braceAngle;
                                angleZ = PI;
                            }
                            else
                            {
                                componentDictionary[BRACESECTION].SetPropertyValue(-braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                componentDictionary[BRACESECTION].SetPropertyValue((PI / 2 - braceAngle), "IJOAhsCutback", "CutbackEndAngle");
                                braceVerOffset = braceVerticalOffset + angleLength;
                                angleY = -(PI - braceAngle);
                                angleZ = -PI;
                            }
                        }
                        else if (routeConnectionValue == 2 || routeConnectionValue == 3)
                        {
                            if (connection == 2)
                            {
                                componentDictionary[BRACESECTION].SetPropertyValue(-braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                componentDictionary[BRACESECTION].SetPropertyValue((PI / 2 - braceAngle), "IJOAhsCutback", "CutbackEndAngle");
                                braceVerOffset = braceVerticalOffset + angleLength;
                                angleY = (PI - braceAngle);
                                angleZ = -PI;
                            }
                            else if (connection == 3)
                            {
                                componentDictionary[BRACESECTION].SetPropertyValue(-braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                componentDictionary[BRACESECTION].SetPropertyValue((PI / 2 - braceAngle), "IJOAhsCutback", "CutbackEndAngle");
                                braceVerOffset = braceVerticalOffset + angleLength;
                                angleY = -(PI - braceAngle);
                                angleZ = -PI;
                            }
                            else
                            {
                                componentDictionary[BRACESECTION].SetPropertyValue(braceAngle, "IJOAhsCutback", "CutbackEndAngle");
                                componentDictionary[BRACESECTION].SetPropertyValue((PI / 2 - braceAngle), "IJOAhsCutback", "CutbackBeginAngle");
                                braceVerOffset = -braceVerticalOffset - angleLength;
                                angleY = -braceAngle;
                                angleZ = 0;
                            }
                        }
                    }
                    //Add Joint Between Brace Section1 and Vertical Section
                    JointHelper.CreateAngularRigidJoint(BRACESECTION, bracePort, VERTICALSECTION, verticalSectionPort, new Vector(braceHorOffset, braceAxisOffset, braceVerOffset), new Vector(angleX, angleY, angleZ));   //-(Pi - dBraceAngle), -Pi)

                    //Add Joint Between the Ports of  Brace section
                    JointHelper.CreatePrismaticJoint(BRACESECTION, "BeginCap", BRACESECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    //Add Joint Between the Ports of Bracesection and structure
                    JointHelper.CreatePointOnPlaneJoint(BRACESECTION, braceStructPort, "-1", "Structure", Plane.XY);
                }

                string routePortName = string.Empty;

                for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                {
                    if (isUboltExist[routeIndex - 1] == true)
                    {
                        if (routeIndex == 1)
                            routePortName = "Route";
                        else
                            routePortName = "Route_" + routeIndex;

                        //Add Joint Between the UBolt and Route
                        JointHelper.CreateTranslationalJoint(uBolt[uBoltCount[routeIndex - 1]], "Route", HORIZONTALSECTION, "Neutral", routeUboltConfigIndex[routeIndex - 1].A, routeUboltConfigIndex[routeIndex - 1].B, routeUboltConfigIndex[routeIndex - 1].C, routeUboltConfigIndex[routeIndex - 1].D, boltRoutePlaneOffset);

                        //Add Joint Between the UBolt and Route
                        JointHelper.CreatePointOnAxisJoint(uBolt[uBoltCount[routeIndex - 1]], "Route", "-1", routePortName, Axis.X);
                    }
                }

                //Add joint for Pad and section
                if (includeLeftPad == true)
                {
                    JointHelper.CreateRigidJoint(LEFTPAD, padPort, VERTICALSECTION, verticalSectionPadPort, padConfigIndex.A, padConfigIndex.B, padConfigIndex.C, padConfigIndex.D, 0, 0, 0);
                    if (isCornerSupport == true)
                        JointHelper.CreateRigidJoint(HORIZONTALSECTIONPAD, horizontalPadPort, HORIZONTALSECTION, horizontalSectionPadPort, padConfigIndex.A, padConfigIndex.B, padConfigIndex.C, padConfigIndex.D, 0, 0, 0);
                }

                //Set the Over Lengths

                double overLength;
                if (slopedSteel == true && connection == 3)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (frameSnipToFlge == true || frameSnipToWeb == true)
                        {
                            overLength = sectionDepth * Math.Tan(structAngle);

                            if (structAngle < 0)        //Negative Angle
                            {
                                if (routeConnectionValue == 2 || routeConnectionValue == 4)
                                    verticalSectionEndOverLength = verticalSectionEndOverLength - overLength;
                            }
                            else                            //Positive Angle
                            {
                                if (routeConnectionValue == 1 || routeConnectionValue == 3)
                                    verticalSectionEndOverLength = verticalSectionEndOverLength + overLength;
                            }
                        }
                    }
                }

                if (slopedSteel == true && supportingType == "Slab")
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (frameSnipToFlge == true || frameSnipToWeb == true)
                        {
                            overLength = sectionDepth * Math.Tan(structAngle);

                            if (structAngle < 0)      //Negative Angle
                            {
                                if (slopedSteelY == true)
                                {
                                    if (routeConnectionValue == 2 || routeConnectionValue == 4)
                                        verticalSectionEndOverLength = verticalSectionEndOverLength - overLength;
                                }
                                else
                                {
                                    if (routeConnectionValue == 1 || routeConnectionValue == 2)
                                        verticalSectionEndOverLength = verticalSectionEndOverLength - overLength;
                                }
                            }
                            else                        //Positive Angle
                            {
                                if (slopedSteelY == true)
                                {
                                    if (routeConnectionValue == 1 || routeConnectionValue == 3)
                                        verticalSectionEndOverLength = verticalSectionEndOverLength + overLength;
                                }
                                else
                                {
                                    if (routeConnectionValue == 3 || routeConnectionValue == 4)
                                        verticalSectionEndOverLength = verticalSectionEndOverLength + overLength;
                                }
                            }
                        }
                    }
                }

                if (slopedRoute == true && supportingType == "Slab")
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (frameSnipToFlge == false || frameSnipToWeb == false)
                        {
                            overLength = sectionDepth * Math.Tan(routeAngle[0]);
                            if (routeAngle[0] < 0)     //Negative Angle
                            {
                                if (routeConnectionValue == 3 || routeConnectionValue == 4)
                                    verticalSectionEndOverLength = verticalSectionEndOverLength - overLength;
                            }
                            else                        //Positive Angle
                            {
                                if (routeConnectionValue == 3 || routeConnectionValue == 2)
                                    verticalSectionEndOverLength = verticalSectionEndOverLength + overLength;
                            }
                        }
                    }
                }

                if (routeConnectionValue == 1 || routeConnectionValue == 4 || sectionType == 4)
                {
                    if (connection == 1 || sectionType == 4)
                        componentDictionary[VERTICALSECTION].SetPropertyValue(verticalSectionEndOverLength, "IJUAHgrOccOverLength", "EndOverLength");
                    else
                        componentDictionary[VERTICALSECTION].SetPropertyValue(verticalSectionEndOverLength, "IJUAHgrOccOverLength", "BeginOverLength");
                }
                else if (routeConnectionValue == 2 || routeConnectionValue == 3)
                {
                    if (connection == 1)
                        componentDictionary[VERTICALSECTION].SetPropertyValue(verticalSectionEndOverLength, "IJUAHgrOccOverLength", "BeginOverLength");
                    else
                        componentDictionary[VERTICALSECTION].SetPropertyValue(verticalSectionEndOverLength, "IJUAHgrOccOverLength", "EndOverLength");
                }

                if (isCornerSupport == true)
                {
                    if (routeConnectionValue == 1 || routeConnectionValue == 4 || sectionType == 4)
                        componentDictionary[HORIZONTALSECTION].SetPropertyValue(verticalSectionEndOverLength, "IJUAHgrOccOverLength", "BeginOverLength");
                    else if (routeConnectionValue == 2 || routeConnectionValue == 3)
                        componentDictionary[HORIZONTALSECTION].SetPropertyValue(verticalSectionEndOverLength, "IJUAHgrOccOverLength", "EndOverLength");

                }

                if (includeBrace == true)
                {
                    if (includeBracePad == true)
                    {
                        componentDictionary[HORIZONTALSECTION].SetPropertyValue(-padProperties.padThickness / Math.Cos(braceAngle), "IJUAHgrOccOverLength", frameLength);

                        //Add joint Between BracePad and Brace section
                        JointHelper.CreateRigidJoint(BRACEPAD, bracePadPort, BRACESECTION, braceCapPort, padConfigIndex.A, padConfigIndex.B, padConfigIndex.C, padConfigIndex.D, 0, 0, 0);
                    }
                }
            }

            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(HORIZONTALSECTION, 1)); // partindex, routeindex

                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Struct Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Structure connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    structConnections.Add(new ConnectionInfo(VERTICALSECTION, 1)); // partindex, routeindex

                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                string bomDescriptionValue = (string)((PropertyValueString)support.GetPropertyValue("IJOAhsMrnBOMDesc", "BOM_DESC")).PropValue;

                if (string.IsNullOrEmpty(bomDescriptionValue))
                    bomDescription = (string)((PropertyValueString)support.GetPropertyValue("IJOAhsMrnSupType", "SupType")).PropValue + "-" + MarineAssemblyServices.GetLargePipeDiameter(this);
                else
                    bomDescription = bomDescriptionValue;

                return bomDescription;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - RectHorTypeB" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}

