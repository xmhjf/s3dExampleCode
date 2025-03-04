//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   TypeU.cs
//   Marine_Assy,Ingr.SP3D.Content.Support.Rules.TypeU
//   Author       :Vijaya
//   Creation Date:5.Aug.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  5.Aug.2013     Vijaya    CR-CP-224470 Convert HS_Marine_Assy to C# .Net
//  04.Dec.2013   Rajeswari  DI-CP-241804 Modified the code as part of hardening
//  11.Dec.2014      PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    public class TypeU : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string HORSECTION1 = "HORSECTION1";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";
        private const string LEFTPAD = "LEFTPAD";
        private const string RIGHTPAD = "RIGHTPAD";
        private const string BRACESECTION1 = "BRACESECTION1";
        private const string BRACESECTION2 = "BRACESECTION2";
        private const string BRACEPAD1 = "BRACEPAD1";
        private const string BRACEPAD2 = "BRACEPAD2";
        private double PI = 0;
        static int routeCount = 0;
        string[] uBolts = new string[routeCount];
        private string[] uBoltPart1 = new string[routeCount];
        int[] ubolt = new int[routeCount];
        string sectionSize, supportingType;
        bool includeLeftPad, includeRightPad, includeBracePad, includeBrace, frameSnipToFlge, frameSnipToWeb, isVerticalRoute, isVerticalStruct, slopedSteel, slopedRoute, slopedSteelY;
        double structAngle = 0.0, cutbackangle = 0.0, uBoltOffset, leftStructAngle = 0.0, rightStructAngle = 0.0;
        double[] routeAngle;
        bool[] isUboltExist;
        MarineAssemblyServices.PADProperties padProperties;
        PropertyValueCodelist sectionTypeCodeList, connectionCodeList;

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
                    sectionTypeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnSection", "SectionType");
                    PropertyValueCodelist sectionSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnSection", "SectionSize");
                    includeRightPad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnRightPad", "IncludeRightPad")).PropValue;
                    includeLeftPad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnLeftPad", "IncludeLeftPad")).PropValue;
                    bool uBoltOption = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnPAttachOpt", "PipeAttachOption")).PropValue;
                    uBoltOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnPAttachOff", "Offset")).PropValue;
                    bool uBoltOffSetFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnPAttachOff", "OffsetFrmRule")).PropValue;
                    connectionCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnFrmConn", "Connection");
                    PropertyValueCodelist padShapeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnPadShape", "PadShape");
                    includeBracePad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnBracePad", "IncludeBrPad")).PropValue;
                    includeBrace = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnBrace", "IncludeBrace")).PropValue;
                    bool frameCutbackToSection = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSecSnips", "FrameCutbacktoSection")).PropValue;
                    double sectionCutbackAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSecSnips", "SecCutbackAngle")).PropValue;
                    frameSnipToFlge = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FrameSniptoFlange")).PropValue;
                    frameSnipToWeb = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "FrameSniptoWeb")).PropValue;

                    routeCount = SupportHelper.SupportedObjects.Count;
                    //Initialize for the existing symbol
                    if (frameSnipToFlge != true && frameSnipToFlge != false)
                        frameSnipToFlge = false;

                    if (frameSnipToWeb != true && frameSnipToWeb != false)
                        frameSnipToWeb = false;

                    //Initialize the cutback angle
                    if (frameCutbackToSection != true && frameCutbackToSection != false)
                    {
                        cutbackangle = PI / 4;     //For backward Compatibility
                        frameCutbackToSection = false;
                    }
                    if (routeCount > 10)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Number of Pipes should be less than or equal to 10", "", "TypeU.cs", 107);
                        return null;
                    }

                    //Warning message for flatbar if wrapping or welded connection selected
                    if (sectionTypeCodeList.PropValue == 4)
                    {
                        if (connectionCodeList.PropValue == 2 || connectionCodeList.PropValue == 3)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Wrapping and Welded connections are not applicable for Flatbar Section", "", "TypeU.cs", 116);
                        }
                    }

                    if (connectionCodeList.PropValue != 3)  //Welded joint
                    {
                        if (frameSnipToFlge == true || frameSnipToWeb == true)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Parts" + ": " + "WARNING: " + "Properties that control the Snip to Frame are applicable only for welded connection. Resetting the values.", "", "TypeU.cs", 125);
                            frameSnipToFlge = false;
                            frameSnipToWeb = false;
                            support.SetPropertyValue(frameSnipToFlge, "IJOAhsMrnSnipAlongFlg", "FrameSniptoFlange");
                            support.SetPropertyValue(frameSnipToWeb, "IJOAhsMrnSnipAlongWeb", "FrameSniptoWeb");
                        }
                    }

                    //Get the Section Size
                    sectionSize = MarineAssemblyServices.GetSectionSize(this);
                    //Get the UBolt
                    string[] uBoltPart = MarineAssemblyServices.GetUboltPart(this);

                    int[] pipeAttachment = new int[routeCount];
                    uBoltPart1 = new string[routeCount];
                    isUboltExist = new bool[routeCount];
                    string tempPipeAttachment = string.Empty;
                    Array.Resize(ref ubolt, routeCount);
                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        tempPipeAttachment = "U" + routeIndex;
                        pipeAttachment[routeIndex - 1] = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnPAttach", tempPipeAttachment)).PropValue;
                        if (routeIndex == 1)
                            ubolt[routeIndex - 1] = 2 + routeIndex;     //Initialize to 1
                        else
                            ubolt[routeIndex - 1] = ubolt[routeIndex - 1 - 1];
                        if (uBoltOption == false)      //User defined
                        {
                            if (pipeAttachment[routeIndex - 1] == 3)       //None
                                isUboltExist[routeIndex - 1] = false;
                            else
                            {
                                ubolt[routeIndex - 1] = ubolt[routeIndex - 1] + 1;
                                isUboltExist[routeIndex - 1] = true;
                            }
                        }
                        else        //Default Ubolt is True
                        {
                            ubolt[routeIndex - 1] = ubolt[routeIndex - 1] + 1;
                            isUboltExist[routeIndex - 1] = true;
                        }
                    }

                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        if (isUboltExist[routeIndex - 1] == true)
                        {
                            string[] temp = uBoltPart[routeIndex - 1].Split('-');
                            uBoltPart1[routeIndex - 1] = temp[0] + "-";
                        }
                    }

                    //set the cutback angle
                    if (frameCutbackToSection == true)
                        cutbackangle = sectionCutbackAngle;

                    //Set the Snip Angle
                    if (frameSnipToFlge == true || frameSnipToWeb == true)
                        if (sectionTypeCodeList.PropValue != 1)    //Snipped Steel is only applicable for L Section
                        {
                            frameSnipToFlge = false;
                            frameSnipToWeb = false;
                        }

                    routeAngle = MarineAssemblyServices.GetRoutAngleAndIsRouteVertical(this, routeCount, out isVerticalRoute);
                    if (routeCount > 1)
                    {
                        for (int index = 1; index <= routeCount - 1; index++)
                        {
                            if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle[index - 1], 3), Math.Round(routeAngle[index], 3)) == false)
                            {
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe slopes are not equal. Please check", "", "TypeU.cs", 196);
                                return null;
                            }
                        }
                    }
                    supportingType = MarineAssemblyServices.GetSupportingTypes(this);

                    // Checking the angles of Stucture X and Stucture Y axes with Global Z axis
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
                    {
                        isVerticalStruct = true;
                    }
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
                    MarineAssemblyServices.GetLeftRightStructAngleAndIsVerticalStructure(this, out isVerticalStruct, out leftStructAngle, out rightStructAngle);

                    if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle[0], 3), 0) == false) //for sloped Route
                        slopedRoute = true;
                    else if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(leftStructAngle, 3), 0) == false || Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(rightStructAngle, 3), 0) == false)    //for sloped steel
                        slopedSteel = true;
                    if (supportingType.ToUpper() == "SLAB")     //for sloped Slab
                        if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(structAngle, 2), 0)== false)
                            slopedSteel = true;

                    //Get Pad part and Dimensions
                    string sectionCode, steelStd;
                    padProperties = MarineAssemblyServices.GetPadPartAndDimensions(this, sectionSize, out sectionCode, out steelStd);

                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    //Get the UboltOffset
                    if (uBoltOffSetFromRule == true)
                    {
                        PartClass hsMarineServiceDim = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_PAttachOffset");
                        IEnumerable<BusinessObject> hsMarineServiceDimPart = hsMarineServiceDim.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        hsMarineServiceDimPart = hsMarineServiceDimPart.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvPAttachOff", "SectionSize")).PropValue == sectionCode));
                        if (hsMarineServiceDimPart.Count() > 0)
                            uBoltOffset = (double)((PropertyValueDouble)hsMarineServiceDimPart.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvPAttachOff", "PipeAttachOffset")).PropValue;
                    }

                    //Create the list of parts                  
                    //GetSteelStandard    
                    string steelStandard = MarineAssemblyServices.GetSteelStandard(this, steelStd);

                    parts.Add(new PartInfo(HORSECTION1, sectionSize + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION1, sectionSize + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION2, sectionSize + " " + steelStandard));

                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        if (isUboltExist[routeIndex - 1] == true)
                        {
                            Array.Resize(ref uBolts, parts.Count + 2);
                            uBolts[parts.Count + 1] = "sUbolt_" + routeIndex + 1;
                            parts.Add(new PartInfo(uBolts[ubolt[routeIndex - 1]], uBoltPart[routeIndex - 1]));
                        }
                    }

                    if (includeLeftPad == true)
                        parts.Add(new PartInfo(LEFTPAD, padProperties.padPart));

                    if (includeRightPad == true)
                        parts.Add(new PartInfo(RIGHTPAD, padProperties.padPart));

                    if (includeBrace == true)
                    {
                        parts.Add(new PartInfo(BRACESECTION1, sectionSize + " " + steelStandard));
                        parts.Add(new PartInfo(BRACESECTION2, sectionSize + " " + steelStandard));
                        if (includeBracePad == true)
                        {
                            parts.Add(new PartInfo(BRACEPAD1, padProperties.padPart));
                            parts.Add(new PartInfo(BRACEPAD2, padProperties.padPart));
                        }
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
                double boundingBoxWidth = 0.0, boundingBoxHeight = 0.0, leftPipeDiameter = 0.0, rightPipeDiameter = 0.0;
                MarineAssemblyServices.GetBoundingBoxDimensionsAndPipeDiameter(this, routeCount, ref boundingBoxWidth, ref boundingBoxHeight, ref leftPipeDiameter, ref rightPipeDiameter);

                //Auto Dimensioning of Supports
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                int routeIndex = 0;
                for (routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    if (isUboltExist[routeIndex - 1] == true)
                        MarineAssemblyServices.CreateDimensionNote(this, "Dim " + routeIndex, componentDictionary[uBolts[ubolt[routeIndex - 1]]], "Route");

                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + routeIndex, componentDictionary[HORSECTION1], "BeginCap");
                routeIndex = routeIndex + 1;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + routeIndex, componentDictionary[HORSECTION1], "EndCap");

                PropertyValueCodelist cardinalCodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJOAhsSteelCP", "CP2");
                cardinalCodeList.PropValue = 1;
                componentDictionary[VERTSECTION1].SetPropertyValue(cardinalCodeList.PropValue, "IJOAhsSteelCP", "CP2");
                cardinalCodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJOAhsSteelCP", "CP4");
                cardinalCodeList.PropValue = 3;
                componentDictionary[VERTSECTION1].SetPropertyValue(cardinalCodeList.PropValue, "IJOAhsSteelCP", "CP4");

                cardinalCodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION2].GetPropertyValue("IJOAhsSteelCP", "CP1");
                cardinalCodeList.PropValue = 1;
                componentDictionary[VERTSECTION2].SetPropertyValue(cardinalCodeList.PropValue, "IJOAhsSteelCP", "CP1");
                cardinalCodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION2].GetPropertyValue("IJOAhsSteelCP", "CP3");
                cardinalCodeList.PropValue = 3;
                componentDictionary[VERTSECTION2].SetPropertyValue(cardinalCodeList.PropValue, "IJOAhsSteelCP", "CP3");
                string verticalSec1Port = string.Empty, verticalSec2Port = string.Empty;
                if (leftStructAngle <= 0)
                    verticalSec1Port = "EndCap";
                else
                    verticalSec1Port = "EndFace";

                if (rightStructAngle <= 0)
                    verticalSec2Port = "BeginCap";
                else
                    verticalSec2Port = "BeginFace";

                routeIndex = routeIndex + 1;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + routeIndex, componentDictionary[VERTSECTION1], verticalSec1Port);

                routeIndex = routeIndex + 1;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + routeIndex, componentDictionary[VERTSECTION2], verticalSec2Port);

                //Get Section Structure dimensions
                BusinessObject sectionPart = componentDictionary[HORSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                //Get SteelWidth and SteelDepth 
                double sectionWidth = 0.0, sectionDepth = 0.0, steelThickness = 0.0;
                sectionWidth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                sectionDepth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                if (sectionTypeCodeList.PropValue != 4)
                    steelThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                double shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnShoeHeight", "ShoeHeight")).PropValue;
                //Check Width Offset and Height Offset values
                double widthOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnWrappingOff", "WidthOffset")).PropValue;
                double heightOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnWrappingOff", "HeightOffset")).PropValue;
                if (widthOffset > sectionDepth || heightOffset > sectionDepth)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Frame offsets should not be more than frame depth, resetting the values", "", "TypeU.cs", 381);
                    widthOffset = 0.0;
                    heightOffset = 0.0;
                    support.SetPropertyValue(0.0, "IJOAhsMrnWrappingOff", "WidthOffset");
                    support.SetPropertyValue(0.0, "IJOAhsMrnWrappingOff", "HeightOffset");
                    return;
                }

                //Get the Current Location in the Route Connection Cycle
                //get Codelisted Property Values for Hgr Beam
                MarineAssemblyServices.BeamCLProperties beamCLProperties = MarineAssemblyServices.GetBeamCLProperties(componentDictionary);
                string[] partNames = new string[5];
                partNames[0] = HORSECTION1;
                partNames[1] = VERTSECTION1;
                partNames[2] = VERTSECTION2;
                if (includeBrace == true)
                {
                    partNames[3] = BRACESECTION1;
                    partNames[4] = BRACESECTION2;
                }
                //set HGR Beam Properties.
                MarineAssemblyServices.IntializeBeamProperties(componentDictionary, support, beamCLProperties, partNames);

                // Set Values of Part Occurance Attributes
                //For L section Initialize Snipped Properties
                if (connectionCodeList.PropValue == 3)
                {
                    if (frameSnipToFlge == true || frameSnipToWeb == true)
                    {
                        string[] arrayOfKeys1 = new string[2];
                        arrayOfKeys1[0] = VERTSECTION1;
                        arrayOfKeys1[1] = VERTSECTION2;

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
                }

                //set the overlength for the vertcal section
                PropertyValueCodelist orientCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnFrmOrient", "FrmOrient");
                string fromOrient = orientCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(orientCodeList.PropValue).DisplayName;
                if (slopedSteel == true && SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (sectionTypeCodeList.PropValue == 4)
                        {
                            componentDictionary[VERTSECTION1].SetPropertyValue(-leftStructAngle, "IJOAhsCutback", "CutbackEndAngle");
                            componentDictionary[VERTSECTION2].SetPropertyValue(-rightStructAngle, "IJOAhsCutback", "CutbackBeginAngle");
                        }
                        else
                        {
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            if (Configuration == 1 || Configuration == 4)
                            {
                                if (connectionCodeList.PropValue == 1 || connectionCodeList.PropValue == 3)
                                {
                                    componentDictionary[VERTSECTION1].SetPropertyValue(-leftStructAngle, "IJOAhsCutback", "CutbackEndAngle");
                                    componentDictionary[VERTSECTION2].SetPropertyValue(-rightStructAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                }
                                else
                                {
                                    componentDictionary[VERTSECTION1].SetPropertyValue(leftStructAngle, "IJOAhsCutback", "CutbackEndAngle");
                                    componentDictionary[VERTSECTION2].SetPropertyValue(rightStructAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                }
                            }
                            else if (Configuration == 2 || Configuration == 3)
                            {
                                if (connectionCodeList.PropValue == 1 || connectionCodeList.PropValue == 3)
                                {
                                    componentDictionary[VERTSECTION1].SetPropertyValue(leftStructAngle, "IJOAhsCutback", "CutbackEndAngle");
                                    componentDictionary[VERTSECTION2].SetPropertyValue(rightStructAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                }
                                else
                                {
                                    componentDictionary[VERTSECTION1].SetPropertyValue(-leftStructAngle, "IJOAhsCutback", "CutbackEndAngle");
                                    componentDictionary[VERTSECTION2].SetPropertyValue(-rightStructAngle, "IJOAhsCutback", "CutbackBeginAngle");
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
                for (int indxRoute = 1; indxRoute <= routeCount; indxRoute++)
                {
                    if (isUboltExist[indxRoute - 1] == true)
                    {
                        tempUBoltType = "U" + indxRoute;

                        PartClass hsMarineServiceDim = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_PAttachSel");
                        IEnumerable<BusinessObject> hsMarineServiceDimPart = hsMarineServiceDim.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        hsMarineServiceDimPart = hsMarineServiceDimPart.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvPAttachSel", "PipeAttachPart")).PropValue == uBoltPart1[indxRoute - 1]));
                        if (hsMarineServiceDimPart.Count() > 0)
                            uBoltTypes[indxRoute - 1] = (int)((PropertyValueCodelist)hsMarineServiceDimPart.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvPAttachSel", "PipeAttachType")).PropValue;

                        //Get the Ublot Dimensions
                        componentDictionary[uBolts[ubolt[indxRoute - 1]]].SetPropertyValue(pipeDiameter[indxRoute - 1], "IJOAhsPipeOD", "PipeOD");
                        if (uBoltTypes[indxRoute - 1] == 1)
                        {
                            if (sectionTypeCodeList.PropValue == 4)
                                componentDictionary[uBolts[ubolt[indxRoute - 1]]].SetPropertyValue(sectionWidth, "IJOAhsSteelThickness", "SteelThickness");
                            else
                                componentDictionary[uBolts[ubolt[indxRoute - 1]]].SetPropertyValue(steelThickness, "IJOAhsSteelThickness", "SteelThickness");
                        }
                        support.SetPropertyValue(uBoltTypes[indxRoute - 1], "IJOAhsMrnPAttach", tempUBoltType);
                    }
                }

                //Set attribute values for Pads
                double vertSec1EndOverLength = 0.0, vertSec2BeginOverLength = 0.0;
                string[] padPartNames = new string[4];
                if (includeLeftPad == false && includeRightPad == true)
                {
                    vertSec1EndOverLength = vertSec1EndOverLength - padProperties.padThickness;
                    padPartNames[0] = RIGHTPAD;
                }
                else if (includeLeftPad == true && includeRightPad == false)
                {
                    vertSec2BeginOverLength = vertSec2BeginOverLength - padProperties.padThickness;
                    padPartNames[0] = LEFTPAD;
                }
                else if (includeLeftPad == true && includeRightPad == true)
                {
                    vertSec2BeginOverLength = vertSec2BeginOverLength - padProperties.padThickness;
                    vertSec1EndOverLength = vertSec1EndOverLength - padProperties.padThickness;
                    padPartNames[0] = RIGHTPAD;
                    padPartNames[1] = LEFTPAD;
                }

                if (includeBracePad == true)
                {
                    if (includeLeftPad == true && includeRightPad == true)
                    {
                        padPartNames[2] = BRACEPAD1;
                        padPartNames[3] = BRACEPAD2;
                    }
                    else if (includeLeftPad == false && includeRightPad == false)
                    {
                        padPartNames[0] = BRACEPAD1;
                        padPartNames[1] = BRACEPAD2;
                    }
                    else
                    {
                        padPartNames[1] = BRACEPAD1;
                        padPartNames[2] = BRACEPAD2;
                    }
                }

                MarineAssemblyServices.SetPadProperties(componentDictionary, support, padProperties, padPartNames);
                //set the Include brace Pad
                if (includeBrace == false)
                    support.SetPropertyValue(false, "IJOAhsMrnBracePad", "IncludeBrPad");

                //Set Assembly Attributes
                //For UboltOffset
                support.SetPropertyValue(uBoltOffset, "IJOAhsMrnPAttachOff", "Offset");
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;

                //Set Frame Orientation attribute
                CodelistItem codeList = metadataManager.GetCodelistInfo("hsMrnCLFrmOrient", "UDP").GetCodelistItem(fromOrient);
                support.SetPropertyValue(codeList.Value, "IJOAhsMrnFrmOrient", "FrmOrient");

                //Set Section Size properties
                codeList = metadataManager.GetCodelistInfo("hsMrnCLSecSize", "UDP").GetCodelistItem(sectionSize);
                support.SetPropertyValue(codeList.Value, "IJOAhsMrnSection", "SectionSize");

                //Set Overhang attributes
                PropertyValueCodelist overHangCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnOHOption", "OverhangOpt");
                double hgrOverHangLeft = 0.0, hgrOverHangRight = 0.0, userDefOverHangLeft = 0.0, userDefOverHangRight = 0.0, overHangLeft = 0.0, overHangRight = 0.0;

                Collection<object> overHangCollecion = new Collection<object>();
                GenericHelper.GetDataByRule("hsMrnRLFrmOH", (BusinessObject)support, out overHangCollecion);
                if (overHangCollecion[0] == null)
                {
                    hgrOverHangLeft = (double)overHangCollecion[1];
                    hgrOverHangRight = (double)overHangCollecion[2];
                }
                else
                {
                    hgrOverHangLeft = (double)overHangCollecion[0];
                    hgrOverHangRight = (double)overHangCollecion[1];
                }
                if (overHangCodeList.PropValue == 1)    //By Catalog Rule
                {
                    overHangLeft = hgrOverHangLeft - rightPipeDiameter / 2;
                    overHangRight = hgrOverHangRight - leftPipeDiameter / 2;
                    support.SetPropertyValue(hgrOverHangLeft, "IJOAhsMrnOHLeft", "OverhangLeft");
                    support.SetPropertyValue(hgrOverHangRight, "IJOAhsMrnOHRight", "OverhangRight");
                }
                else if (overHangCodeList.PropValue == 3) //User Defined
                {
                    userDefOverHangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHLeft", "OverhangLeft")).PropValue;
                    userDefOverHangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHRight", "OverhangRight")).PropValue;

                    support.SetPropertyValue(userDefOverHangLeft, "IJOAhsMrnOHLeft", "OverhangLeft");
                    support.SetPropertyValue(userDefOverHangRight, "IJOAhsMrnOHRight", "OverhangRight");

                    overHangLeft = userDefOverHangLeft - rightPipeDiameter / 2;
                    overHangRight = userDefOverHangRight - leftPipeDiameter / 2;
                }
                //Do Something if more than one Structure
                //Get the current steel face number
                bool IsSingleSteel = false;
                if (supportingType.ToUpper() == "STEEL")
                    if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                        if (SupportHelper.SupportingObjects.Count < 2)
                            IsSingleSteel = true;

                //Set Length of horzontal member
                double horSectionLength = 0.0, routeLowStruct2HorDist, routeHighStruct2HorDist, routeLowStruct1HorDist, routeHighStruct1HorDist;
                double perHorLowStruct1Dist, perHorLowStruct2Dist, perHorHighStruct1Dist, perHorHighStruct2Dist, distanceBetweenStruct, horizontalLowStruct1Dist, horizontalLowStruct2Dist, horizontalHighStruct1Dist, horizontalHighStruct2Dist;

                //For Sloped Structure
                bool[] isOffsetApplied = new bool[2];
                isOffsetApplied = MarineAssemblyServices.GetIsLugEndOffsetApplied(this);

                string[] structPort = new string[2];
                structPort = MarineAssemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0];
                string rightStructPort = structPort[1];

                //Get the Route Direction Angle
                double routeDirAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_X);
                //Get the Route Struct Config angle
                double routeStructConfigAng = MarineAssemblyServices.GetRouteStructConfigAngle(this, "Route", "Structure", PortAxisType.Y);
                double  routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    double tempDistance = 0.0;
                    if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 3)      //By Catalog Rule
                        tempDistance = boundingBoxWidth;
                    else if (overHangCodeList.PropValue == 2)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Joints" + ": " + "WARNING: " + "Adjust to Supporting steel is applicable only when the support placement is By Point.", "", "TypeU.cs", 643);
                        return;
                    }

                    horSectionLength = tempDistance + overHangLeft + overHangRight;
                    componentDictionary[HORSECTION1].SetPropertyValue(horSectionLength, "IJUAHgrOccLength", "Length");
                }
                else                        //for Place by point
                {
                    if ((Math.Abs(routeAngle[0]) < (0 + 0.01) && Math.Abs(routeAngle[0]) > (0 - 0.01)) || (Math.Abs(routeAngle[0]) < (PI + 0.01) && Math.Abs(routeAngle[0]) > (PI - 0.01)))
                    {
                        if (isVerticalRoute == true)
                        {
                            perHorLowStruct1Dist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Horizontal_Perpendicular);
                            perHorLowStruct2Dist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Horizontal_Perpendicular);
                            perHorHighStruct1Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Horizontal_Perpendicular);
                            perHorHighStruct2Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Horizontal_Perpendicular);
                            distanceBetweenStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
                            horizontalLowStruct1Dist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Horizontal);
                            horizontalLowStruct2Dist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Horizontal);
                            horizontalHighStruct1Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Horizontal);
                            horizontalHighStruct2Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Horizontal);

                            routeLowStruct1HorDist = Math.Sqrt((Math.Pow(horizontalLowStruct1Dist, 2) - Math.Pow(perHorLowStruct1Dist, 2)));
                            routeLowStruct2HorDist = Math.Sqrt((Math.Pow(horizontalLowStruct2Dist, 2) - Math.Pow(perHorLowStruct2Dist, 2)));
                            routeHighStruct1HorDist = Math.Sqrt((Math.Pow(horizontalHighStruct1Dist, 2) - Math.Pow(perHorHighStruct1Dist, 2)));
                            routeHighStruct2HorDist = Math.Sqrt((Math.Pow(horizontalHighStruct2Dist, 2) - Math.Pow(perHorHighStruct2Dist, 2)));
                        }
                        else
                        {
                            routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Horizontal);
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Horizontal);
                            routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Horizontal);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Horizontal);
                            distanceBetweenStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
                        }
                    }
                    else
                    {
                        if ((Math.Abs(routeStructAngle) < (0 + 0.01) && Math.Abs(routeStructAngle) > (0 - 0.01)) || (Math.Abs(routeStructAngle) < (PI + 0.01) && Math.Abs(routeStructAngle) > (PI - 0.01)))
                        {
                            routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Vertical);
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Vertical);
                            routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Vertical);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Vertical);
                            distanceBetweenStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Vertical);
                        }
                        else if (Math.Abs(routeStructAngle) < (PI / 2 + 0.01) && Math.Abs(routeStructAngle) > (PI / 2 - 0.01))
                        {
                            routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Horizontal);
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Horizontal);
                            routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Horizontal);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Horizontal);
                            distanceBetweenStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
                        }
                        else
                        {
                            routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Vertical);
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Vertical);
                            routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Vertical);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Vertical);
                            distanceBetweenStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Vertical);
                        }
                    }

                    if (supportingType.ToUpper() == "SLAB" || supportingType.ToUpper() == "WALL" || IsSingleSteel == true)
                    {
                        if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 3)          //By catalog OR User Defined
                            horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                        else if (overHangCodeList.PropValue == 2)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Adjust to Supporting steel is applicable only when the support placement is By Point and the supporting count is more than one.", "", "TypeT.cs", 714);
                            return;
                        }
                    }
                    else if (supportingType.ToUpper() == "STEEL")
                    {
                        if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 3)         //By catalog
                            horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                        else if (overHangCodeList.PropValue == 2)
                        {
                            if ((Math.Abs(routeDirAngle) < (0 + 0.01) && Math.Abs(routeDirAngle) > (0 - 0.01)) || (Math.Abs(routeDirAngle) < (PI + 0.01) && Math.Abs(routeDirAngle) > (PI - 0.01)))
                            {
                                overHangRight = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist) + sectionWidth / 2;
                                overHangLeft = Math.Min(routeLowStruct1HorDist, routeHighStruct1HorDist) + sectionWidth / 2;
                            }
                            else
                            {
                                if (routeStructConfigAng < PI / 2)
                                {
                                    overHangLeft = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist) + sectionWidth / 2;
                                    overHangRight = Math.Min(routeLowStruct1HorDist, routeHighStruct1HorDist) + sectionWidth / 2;
                                    rightStructPort = leftStructPort;
                                    leftStructPort = rightStructPort;
                                }
                                else
                                {
                                    overHangRight = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist) + sectionWidth / 2;
                                    overHangLeft = Math.Min(routeLowStruct1HorDist, routeHighStruct1HorDist) + sectionWidth / 2;
                                }
                            }

                            horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                            support.SetPropertyValue(overHangLeft + rightPipeDiameter / 2, "IJOAhsMrnOHLeft", "OverhangLeft");
                            support.SetPropertyValue(overHangRight + leftPipeDiameter / 2, "IJOAhsMrnOHRight", "OverhangRight");
                        }
                    }
                    else if (supportingType.ToUpper() == "STEEL-SLAB" || supportingType.ToUpper() == "SLAB-STEEL")
                    {
                        if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 3)         //by catalog
                            horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                        else if (overHangCodeList.PropValue == 2)
                        {
                            if ((Math.Abs(routeDirAngle) < (0 + 0.01) && Math.Abs(routeDirAngle) > (0 - 0.01)) || (Math.Abs(routeDirAngle) < (PI + 0.01) && Math.Abs(routeDirAngle) > (PI - 0.01)))
                            {
                                overHangRight = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist) + sectionWidth / 2;
                                overHangLeft = Math.Min(routeLowStruct1HorDist, routeHighStruct1HorDist) + sectionWidth / 2;
                            }
                            else
                            {
                                if (routeStructConfigAng <= PI / 2 + 0.0001)
                                {
                                    overHangLeft = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist) + sectionWidth / 2;
                                    overHangRight = Math.Min(routeLowStruct1HorDist, routeHighStruct1HorDist) + sectionWidth / 2;
                                    rightStructPort = leftStructPort;
                                    leftStructPort = rightStructPort;
                                }
                                else
                                {
                                    overHangRight = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist) + sectionWidth / 2;
                                    overHangLeft = Math.Min(routeLowStruct1HorDist, routeHighStruct1HorDist) + sectionWidth / 2;
                                }
                            }
                            horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                            support.SetPropertyValue(overHangLeft + rightPipeDiameter / 2, "IJOAhsMrnOHLeft", "OverhangLeft");
                            support.SetPropertyValue(overHangRight + leftPipeDiameter / 2, "IJOAhsMrnOHRight", "OverhangRight");
                        }
                    }
                    componentDictionary[HORSECTION1].SetPropertyValue(horSectionLength, "IJUAHgrOccLength", "Length");
                }

                //set the properties for connections
                double flgSnipAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FlgSnipAngle")).PropValue;
                double webSnipAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "WebSnipAngle")).PropValue;
                double flgSnipOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FlgSnipOffset")).PropValue;
                double webSnipOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "WebSnipOffset")).PropValue;

                if (connectionCodeList.PropValue == 1)    //For Miter Joint
                {
                    widthOffset = 0.0;
                    heightOffset = 0.0;
                    if (sectionTypeCodeList.PropValue != 4)
                    {
                        if (Configuration == 1 || Configuration == 2)
                        {
                            componentDictionary[HORSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[HORSECTION1].SetPropertyValue(-PI / 4, "IJOAhsCutback", "CutbackBeginAngle");
                            componentDictionary[HORSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[HORSECTION1].SetPropertyValue(PI / 4, "IJOAhsCutback", "CutbackEndAngle");
                        }
                        else if (Configuration == 3 || Configuration == 4)
                        {
                            componentDictionary[HORSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[HORSECTION1].SetPropertyValue(PI / 4, "IJOAhsCutback", "CutbackBeginAngle");
                            componentDictionary[HORSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt1AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[HORSECTION1].SetPropertyValue(-PI / 4, "IJOAhsCutback", "CutbackEndAngle");
                        }
                        componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                        componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                        componentDictionary[VERTSECTION1].SetPropertyValue(-PI / 4, "IJOAhsCutback", "CutbackBeginAngle");
                        componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                        componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                        componentDictionary[VERTSECTION2].SetPropertyValue(PI / 4, "IJOAhsCutback", "CutbackEndAngle");
                    }
                }
                else if (connectionCodeList.PropValue == 2)    //For Wrapping Joint
                {
                    componentDictionary[VERTSECTION1].SetPropertyValue(PI, "IJOAhsBeginCap", "BeginCapRotZ");
                    componentDictionary[VERTSECTION2].SetPropertyValue(-PI, "IJOAhsBeginCap", "BeginCapRotZ");
                }
                else if (connectionCodeList.PropValue == 3)    //For Welded joint
                {
                    widthOffset = 0.0;
                    heightOffset = 0.0;
                    if (frameSnipToFlge == false && frameSnipToWeb == false)
                    {
                        if (sectionTypeCodeList.PropValue != 4)
                        {
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[VERTSECTION1].SetPropertyValue(-cutbackangle, "IJOAhsCutback", "CutbackBeginAngle");
                            componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[VERTSECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[VERTSECTION2].SetPropertyValue(cutbackangle, "IJOAhsCutback", "CutbackEndAngle");
                        }
                    }
                    else
                    {
                        if (frameSnipToFlge == true)
                        {
                            componentDictionary[VERTSECTION1].SetPropertyValue(flgSnipAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle1");
                            componentDictionary[VERTSECTION1].SetPropertyValue(flgSnipOffset, "IJOAhsCutbackOffset", "BeginOffsetAlongFlange");
                            componentDictionary[VERTSECTION2].SetPropertyValue(flgSnipAngle, "IJOAhsSnipedSteel", "CutbackEndAngle1");
                            componentDictionary[VERTSECTION2].SetPropertyValue(flgSnipOffset, "IJOAhsCutbackOffset", "EndOffsetAlongFlange");
                        }
                        if (frameSnipToWeb == true)
                        {
                            componentDictionary[VERTSECTION1].SetPropertyValue(webSnipAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle2");
                            componentDictionary[VERTSECTION1].SetPropertyValue(webSnipOffset, "IJOAhsCutbackOffset", "BeginOffsetAlongWeb");
                            componentDictionary[VERTSECTION2].SetPropertyValue(webSnipAngle, "IJOAhsSnipedSteel", "CutbackEndAngle2");
                            componentDictionary[VERTSECTION2].SetPropertyValue(webSnipOffset, "IJOAhsCutbackOffset", "EndOffsetAlongWeb");
                        }
                    }
                }

                //Create Joints
                double horizontalRoutePlaneOffset = 0.0, horizontalRouteAxisOffset = 0.0, horizontalRouteOriginOffset = 0.0, boltRoutePlaneOffset = 0.0, horizontalVertPlaneOffset1 = 0.0, horizontalVertPlaneOffset2 = 0.0, horizontalVertAxisOffset1 = 0.0, horizontalVertAxisOffset2 = 0.0, horizontalVertOriginOffset1 = 0.0, horizontalVertOriginOffset2 = 0.0;
                string horizontalSectionPort = string.Empty, horizontalSectionPort1 = string.Empty, horizontalSectionPort2 = string.Empty;

                //Get Route Port
                string routePort = MarineAssemblyServices.GetRoutePort(this, support, slopedRoute, slopedSteel, isVerticalRoute, isVerticalStruct);
                horizontalRouteOriginOffset = sectionWidth / 2;
                MarineAssemblyServices.ConfigIndex horizontalorVertConfigIndex1 = new MarineAssemblyServices.ConfigIndex(), horizonatlVertConfigIndex2 = new MarineAssemblyServices.ConfigIndex(), routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(), padConfigIndex = new MarineAssemblyServices.ConfigIndex();
                MarineAssemblyServices.ConfigIndex[] routeUboltConfigIndex = new MarineAssemblyServices.ConfigIndex[routeCount];

                for (routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                {
                    if (isUboltExist[routeIndex - 1] == true)
                    {
                        if (sectionTypeCodeList.PropValue != 4)
                            routeUboltConfigIndex[routeIndex - 1] = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.YZ, Axis.Z, Axis.Y);
                        else
                            routeUboltConfigIndex[routeIndex - 1] = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ);
                    }
                }
                if (sectionTypeCodeList.PropValue == 4)       //For RS
                {
                    horizontalRoutePlaneOffset = sectionDepth / 2;
                    horizontalRouteAxisOffset = -overHangRight - boundingBoxWidth;
                    horizontalRouteOriginOffset = sectionWidth + shoeHeight;
                    boltRoutePlaneOffset = -sectionDepth / 2 + uBoltOffset;

                    horizontalSectionPort = "EndCap";
                    horizontalSectionPort1 = "EndCap";
                    horizontalSectionPort2 = "BeginCap";
                    horizontalVertPlaneOffset1 = sectionDepth;
                    horizontalVertPlaneOffset2 = sectionDepth;
                    routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.YZ, Axis.X, Axis.NegativeZ);
                    horizontalorVertConfigIndex1 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeZX, Axis.X, Axis.NegativeZ);
                    horizonatlVertConfigIndex2 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeZX, Axis.X, Axis.Z);
                    padConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.Y, Axis.NegativeY);
                }
                else
                {
                    padConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeY);
                    boltRoutePlaneOffset = sectionWidth / 2 - uBoltOffset;
                    if (Configuration == 1 || Configuration == 2)
                    {
                        horizontalSectionPort1 = "EndCap";
                        horizontalSectionPort2 = "BeginCap";
                        horizontalorVertConfigIndex1 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        horizonatlVertConfigIndex2 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X);
                        if (connectionCodeList.PropValue == 2)
                        {
                            horizontalVertPlaneOffset1 = heightOffset;
                            horizontalVertAxisOffset1 = -horSectionLength + widthOffset;
                            horizontalVertPlaneOffset2 = heightOffset;
                            horizontalVertAxisOffset2 = horSectionLength - widthOffset;
                        }
                        else if (connectionCodeList.PropValue == 3)
                        {
                            horizontalVertAxisOffset1 = -horSectionLength;
                            horizontalVertAxisOffset2 = horSectionLength;
                        }
                    }
                    else if (Configuration == 3 || Configuration == 4)
                    {
                        horizontalSectionPort1 = "BeginCap";
                        horizontalSectionPort2 = "EndCap";
                        horizontalorVertConfigIndex1 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X);
                        horizonatlVertConfigIndex2 = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        if (connectionCodeList.PropValue == 1)
                        {
                            horizontalVertPlaneOffset1 = sectionDepth;
                            horizontalVertAxisOffset1 = 0.0;
                            horizontalVertPlaneOffset2 = sectionDepth;
                            horizontalVertAxisOffset2 = 0.0;
                        }
                        else if (connectionCodeList.PropValue == 2)
                        {
                            horizontalVertPlaneOffset1 = sectionDepth - heightOffset;
                            horizontalVertAxisOffset1 = horSectionLength - widthOffset;
                            horizontalVertPlaneOffset2 = sectionDepth - heightOffset;
                            horizontalVertAxisOffset2 = -horSectionLength + widthOffset;
                        }
                        else if (connectionCodeList.PropValue == 3)
                        {
                            horizontalVertAxisOffset1 = horSectionLength;
                            horizontalVertAxisOffset2 = -horSectionLength;
                            horizontalVertPlaneOffset1 = sectionDepth;
                            horizontalVertPlaneOffset2 = sectionDepth;
                        }
                    }

                    if (Configuration == 1)
                    {
                        if (routePort == "BBR_High" || routePort == "BBSR_High")
                        {
                            horizontalSectionPort = "EndCap";
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                            horizontalRoutePlaneOffset = -shoeHeight;
                            horizontalRouteAxisOffset = -overHangRight - boundingBoxWidth;
                        }
                        else
                        {
                            horizontalSectionPort = "BeginCap";
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX);
                            horizontalRoutePlaneOffset = overHangRight + boundingBoxWidth;
                            horizontalRouteAxisOffset = -shoeHeight;
                        }
                    }
                    else if (Configuration == 2)
                    {
                        if (routePort == "BBR_High" || routePort == "BBSR_High")
                        {
                            horizontalSectionPort = "BeginCap";
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX);
                            horizontalRoutePlaneOffset = overHangRight + boundingBoxWidth;
                            horizontalRouteAxisOffset = -shoeHeight;
                        }
                        else
                        {
                            horizontalSectionPort = "EndCap";
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                            horizontalRoutePlaneOffset = -shoeHeight;
                            horizontalRouteAxisOffset = -overHangRight - boundingBoxWidth;
                        }
                    }
                    else if (Configuration == 3)
                    {
                        if (routePort == "BBR_High" || routePort == "BBSR_High")
                        {
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX);
                            horizontalSectionPort = "BeginCap";
                            horizontalRoutePlaneOffset = -boundingBoxHeight - shoeHeight;
                            horizontalRouteAxisOffset = overHangLeft;
                        }
                        else
                        {
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X);
                            horizontalSectionPort = "EndCap";
                            horizontalRoutePlaneOffset = -boundingBoxHeight - shoeHeight;
                            horizontalRouteAxisOffset = -overHangLeft;
                        }
                    }
                    else if (Configuration == 4)
                    {
                        if (routePort == "BBR_High" || routePort == "BBSR_High")
                        {
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X);
                            horizontalSectionPort = "EndCap";
                            horizontalRoutePlaneOffset = -boundingBoxHeight - shoeHeight;
                            horizontalRouteAxisOffset = -overHangLeft;
                        }
                        else
                        {
                            routeHorConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX);
                            horizontalSectionPort = "BeginCap";
                            horizontalRoutePlaneOffset = -boundingBoxHeight - shoeHeight;
                            horizontalRouteAxisOffset = overHangLeft;
                        }
                    }
                }

                //Add joint Between Horizontal Section and BoundingBox
                JointHelper.CreateRigidJoint(HORSECTION1, horizontalSectionPort, "-1", routePort, routeHorConfigIndex.A, routeHorConfigIndex.B, routeHorConfigIndex.C, routeHorConfigIndex.D, horizontalRoutePlaneOffset, horizontalRouteAxisOffset, horizontalRouteOriginOffset);

                //Add Joint Between the Horizontal and Vertical Beams
                JointHelper.CreateRigidJoint(HORSECTION1, horizontalSectionPort1, VERTSECTION1, "BeginCap", horizontalorVertConfigIndex1.A, horizontalorVertConfigIndex1.B, horizontalorVertConfigIndex1.C, horizontalorVertConfigIndex1.D, horizontalVertPlaneOffset1, horizontalVertAxisOffset1, horizontalVertOriginOffset1);

                //Add Joint Between the Horizontal and Vertical Beams
                JointHelper.CreateRigidJoint(HORSECTION1, horizontalSectionPort2, VERTSECTION2, "EndCap", horizonatlVertConfigIndex2.A, horizonatlVertConfigIndex2.B, horizonatlVertConfigIndex2.C, horizonatlVertConfigIndex2.D, horizontalVertPlaneOffset2, horizontalVertAxisOffset2, horizontalVertOriginOffset2);

                //Add Joint Between the Ports of Vertical Beam 1
                JointHelper.CreatePrismaticJoint(VERTSECTION1, "BeginCap", VERTSECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add Joint Between the Ports of Vertical Beam 2
                JointHelper.CreatePrismaticJoint(VERTSECTION2, "BeginCap", VERTSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add Joint Between the Ports of Vertical Beam 1
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION1, "EndCap", "-1", leftStructPort, Plane.XY);

                //Add Joint Between the Ports of Vertical Beam 2
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION2, "BeginCap", "-1", rightStructPort, Plane.XY);

                //Joint for Ubolt
                string strRefPortName;
                for (routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                {
                    if (isUboltExist[routeIndex - 1] == true)
                    {
                        if (routeIndex == 1)
                            strRefPortName = "Route";
                        else
                            strRefPortName = "Route_" + routeIndex;

                        //Add Joint Between the UBolt and Route
                        JointHelper.CreateTranslationalJoint(uBolts[ubolt[routeIndex - 1]], "Route", HORSECTION1, "Neutral", routeUboltConfigIndex[routeIndex - 1].A, routeUboltConfigIndex[routeIndex - 1].B, routeUboltConfigIndex[routeIndex - 1].C, routeUboltConfigIndex[routeIndex - 1].D, boltRoutePlaneOffset);

                        //Add Joint Between the UBolt and Route
                        JointHelper.CreatePointOnAxisJoint(uBolts[ubolt[routeIndex - 1]], "Route", "-1", strRefPortName, Axis.X);//1
                    }
                }

                if (includeLeftPad == false && includeRightPad == true)
                    //Add Joint Between the Left and the Vertical Section 1
                    JointHelper.CreateRigidJoint(RIGHTPAD, "Port1", VERTSECTION1, "EndFace", padConfigIndex.A, padConfigIndex.B, padConfigIndex.C, padConfigIndex.D, 0, 0, 0);

                else if (includeLeftPad == true && includeRightPad == false)
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(LEFTPAD, "Port2", VERTSECTION2, "BeginFace", padConfigIndex.A, padConfigIndex.B, padConfigIndex.C, padConfigIndex.D, 0, 0, 0);

                else if (includeRightPad == true && includeLeftPad == true)
                {
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(RIGHTPAD, "Port1", VERTSECTION1, "EndFace", padConfigIndex.A, padConfigIndex.B, padConfigIndex.C, padConfigIndex.D, 0, 0, 0);

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(LEFTPAD, "Port2", VERTSECTION2, "BeginFace", padConfigIndex.A, padConfigIndex.B, padConfigIndex.C, padConfigIndex.D, 0, 0, 0);

                }

                //Set the Over Length for Snipped Steel for Sloped Structure
                double overLength;
                if (slopedSteel == true && connectionCodeList.PropValue == 3)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (frameSnipToFlge == true || frameSnipToWeb == true)
                        {
                            overLength = sectionDepth * Math.Tan(leftStructAngle);

                            if (leftStructAngle < 0)        //Negative Angle
                            {
                                if (Configuration == 2 || Configuration == 3)
                                    vertSec1EndOverLength = vertSec1EndOverLength - overLength;
                                else if (Configuration == 1 || Configuration == 4)
                                    vertSec2BeginOverLength = vertSec2BeginOverLength - overLength;
                            }
                            else                            //Positive Angle
                            {
                                if (Configuration == 1 || Configuration == 4)
                                    vertSec1EndOverLength = vertSec1EndOverLength + overLength;
                                else if (Configuration == 2 || Configuration == 3)
                                    vertSec2BeginOverLength = vertSec2BeginOverLength + overLength;
                            }
                        }
                    }
                }

                if (slopedSteel == true && supportingType.ToUpper() == "SLAB")
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (frameSnipToFlge == false || frameSnipToWeb == false)
                        {
                            overLength = sectionDepth * Math.Tan(structAngle);
                            if (structAngle < 0)      //Negative Angle
                            {
                                if (slopedSteelY == true)
                                {
                                    if (Configuration == 1 || Configuration == 4)
                                        vertSec1EndOverLength = vertSec1EndOverLength - overLength;
                                    else if (Configuration == 2 || Configuration == 3)
                                        vertSec2BeginOverLength = vertSec2BeginOverLength - overLength;
                                }
                                else
                                    if (Configuration == 2 || Configuration == 3)
                                    {
                                        vertSec1EndOverLength = vertSec1EndOverLength - overLength;
                                        vertSec2BeginOverLength = vertSec2BeginOverLength - overLength;
                                    }
                            }
                            else                            //Positive Angle
                            {
                                if (slopedSteelY == true)
                                {
                                    if (Configuration == 1 || Configuration == 4)
                                        vertSec2BeginOverLength = vertSec2BeginOverLength + overLength;
                                    else if (Configuration == 2 || Configuration == 3)
                                        vertSec1EndOverLength = vertSec1EndOverLength + overLength;
                                }
                                else
                                {
                                    if (Configuration == 1 || Configuration == 4)
                                    {
                                        vertSec1EndOverLength = vertSec1EndOverLength + overLength;
                                        vertSec2BeginOverLength = vertSec2BeginOverLength + overLength;
                                    }
                                }
                            }
                        }
                    }
                }

                if (slopedRoute == true && supportingType.ToUpper() == "SLAB")
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (frameSnipToFlge == false || frameSnipToWeb == false)
                        {
                            overLength = sectionDepth * Math.Tan(routeAngle[0]);
                            if (routeAngle[0] < 0)      //Negative Angle
                            {
                                if (Configuration == 2 || Configuration == 3)
                                {
                                    vertSec1EndOverLength = vertSec1EndOverLength - overLength;
                                    vertSec2BeginOverLength = vertSec2BeginOverLength - overLength;
                                }
                            }
                            else                            //Positive Angle
                            {
                                if (Configuration == 1 || Configuration == 4)
                                {
                                    vertSec2BeginOverLength = vertSec2BeginOverLength + overLength;
                                    vertSec1EndOverLength = vertSec1EndOverLength + overLength;
                                }
                            }
                        }
                    }
                }
                componentDictionary[VERTSECTION1].SetPropertyValue(vertSec1EndOverLength, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(vertSec2BeginOverLength, "IJUAHgrOccOverLength", "BeginOverLength");

                //For  Brace
                double braceAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnBraceAngle", "BraceAngle")).PropValue;
                //get Brace Offset
                double braceHorizontalOffset = 0.0, braceVerticalOffset = 0.0;
                MarineAssemblyServices.GetBraceOffset(this, sectionSize, ref braceHorizontalOffset, ref braceVerticalOffset);
                if (includeBrace == true)
                {
                    double braceVerOffset = 0.0, braceHorOffset = 0.0, braceAxisOffset = 0.0, angleX1 = 0.0, angleX2 = 0.0, angleY1 = 0.0, angleZ1 = 0.0, angleY2 = 0.0, angleZ2 = 0.0, angleLength = 0.0;

                    if (sectionTypeCodeList.PropValue == 4)
                        angleLength = sectionDepth / Math.Sin(braceAngle);
                    else
                        angleLength = sectionWidth / Math.Sin(braceAngle);

                    braceVerOffset = braceVerticalOffset + angleLength;
                    if (sectionTypeCodeList.PropValue == 4)
                    {
                        componentDictionary[BRACESECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                        componentDictionary[BRACESECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                        componentDictionary[BRACESECTION2].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                        componentDictionary[BRACESECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                        componentDictionary[BRACESECTION1].SetPropertyValue(-braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                        componentDictionary[BRACESECTION1].SetPropertyValue((PI / 2 - braceAngle), "IJOAhsCutback", "CutbackEndAngle");
                        componentDictionary[BRACESECTION2].SetPropertyValue(braceAngle, "IJOAhsCutback", "CutbackEndAngle");
                        componentDictionary[BRACESECTION2].SetPropertyValue(-(PI / 2 - braceAngle), "IJOAhsCutback", "CutbackBeginAngle");

                        braceAxisOffset = braceHorizontalOffset;
                        braceHorOffset = -sectionWidth;
                        angleX1 = -(PI - braceAngle);
                        angleX2 = PI - braceAngle;
                        angleZ1 = -PI;
                        angleZ2 = -PI;
                    }
                    else
                    {
                        if (connectionCodeList.PropValue == 2)
                        {
                            componentDictionary[BRACESECTION1].SetPropertyValue(PI, "IJOAhsBeginCap", "BeginCapRotZ");
                            componentDictionary[BRACESECTION2].SetPropertyValue(-PI, "IJOAhsBeginCap", "BeginCapRotZ");
                            componentDictionary[BRACESECTION1].SetPropertyValue(-braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                            componentDictionary[BRACESECTION1].SetPropertyValue((PI / 2 - braceAngle), "IJOAhsCutback", "CutbackEndAngle");
                            componentDictionary[BRACESECTION2].SetPropertyValue(braceAngle, "IJOAhsCutback", "CutbackEndAngle");
                            componentDictionary[BRACESECTION2].SetPropertyValue(-(PI / 2 - braceAngle), "IJOAhsCutback", "CutbackBeginAngle");

                            angleY1 = -(PI - braceAngle);
                            angleY2 = PI - braceAngle;
                            angleZ1 = 0.0;
                            angleZ2 = 0.0;
                        }
                        else
                        {
                            componentDictionary[BRACESECTION1].SetPropertyValue(-braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                            componentDictionary[BRACESECTION1].SetPropertyValue((PI / 2 - braceAngle), "IJOAhsCutback", "CutbackEndAngle");
                            componentDictionary[BRACESECTION2].SetPropertyValue(braceAngle, "IJOAhsCutback", "CutbackEndAngle");
                            componentDictionary[BRACESECTION2].SetPropertyValue(-(PI / 2 - braceAngle), "IJOAhsCutback", "CutbackBeginAngle");

                            angleY1 = -(PI - braceAngle);
                            angleY2 = PI - braceAngle;
                            angleZ1 = -PI;
                            angleZ2 = -PI;
                        }
                        if (connectionCodeList.PropValue == 2)
                        {
                            braceAxisOffset = -horSectionLength;
                            braceHorOffset = -braceHorizontalOffset;
                        }
                        else
                        {
                            braceAxisOffset = 0.0;
                            braceHorOffset = braceHorizontalOffset;
                        }
                    }

                    //Add Joint Between Brace Section1 and Vertical Section1
                    JointHelper.CreateAngularRigidJoint(BRACESECTION1, "EndCap", VERTSECTION1, "BeginCap", new Vector(braceHorOffset, braceAxisOffset, braceVerOffset), new Vector(angleX1, angleY1, angleZ1));

                    //Add Joint Between Brace Section2 and Vertical Section2
                    JointHelper.CreateAngularRigidJoint(BRACESECTION2, "BeginCap", VERTSECTION2, "EndCap", new Vector(braceHorOffset, braceAxisOffset, -braceVerOffset), new Vector(angleX2, angleY2, angleZ2));

                    //Add Joint Between the Ports of  Brace section1
                    JointHelper.CreatePrismaticJoint(BRACESECTION1, "BeginCap", BRACESECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);// 11757)

                    //Add Joint Between the Ports of Brace section2
                    JointHelper.CreatePrismaticJoint(BRACESECTION2, "BeginCap", BRACESECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);//11757)

                    //Add Joint Between the Ports of Bracesection 1
                    JointHelper.CreatePointOnPlaneJoint(BRACESECTION1, "BeginCap", "-1", leftStructPort, Plane.XY);//4

                    //Add Joint Between the Ports of Bracesection 2
                    JointHelper.CreatePointOnPlaneJoint(BRACESECTION2, "EndCap", "-1", rightStructPort, Plane.XY);// 4)

                    if (includeBracePad == true)
                    {
                        componentDictionary[BRACESECTION1].SetPropertyValue(-padProperties.padThickness / Math.Cos(braceAngle), "IJUAHgrOccOverLength", "BeginOverLength");
                        componentDictionary[BRACESECTION2].SetPropertyValue(-padProperties.padThickness / Math.Cos(braceAngle), "IJUAHgrOccOverLength", "EndOverLength");

                        //Add joint Between BracePad and Brace section
                        JointHelper.CreateRigidJoint(BRACEPAD1, "Port2", BRACESECTION1, "BeginFace", padConfigIndex.A, padConfigIndex.B, padConfigIndex.C, padConfigIndex.D, 0, 0, 0);

                        //Add joint Between BracePad and Brace section
                        JointHelper.CreateRigidJoint(BRACEPAD2, "Port1", BRACESECTION2, "EndFace", padConfigIndex.A, padConfigIndex.B, padConfigIndex.C, padConfigIndex.D, 0, 0, 0);
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
                    routeConnections.Add(new ConnectionInfo(HORSECTION1, 1)); // partindex, routeindex

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
                    structConnections.Add(new ConnectionInfo(VERTSECTION1, 1)); // partindex, routeindex

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