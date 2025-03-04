//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   TypeL.cs
//   Marine_Assy,Ingr.SP3D.Content.Support.Rules.TypeI
//   Author       :Vijay
//   Creation Date:30.July.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  30.July.2013     Vijay    CR-CP-224470 Convert HS_Marine_Assy to C# .Net
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
    public class TypeI : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string SECTION = "SECTION";
        private const string CANTIPAD = "CANTIPAD";
        private const string BRACEPAD = "BRACEPAD";
        private const string BRACESECTION = "BRACESECTION";
        private string sectionSize = string.Empty, supportingType = string.Empty;
        private int overhangOption = 0;
        private double uBoltOffset = 0, structAngle = 0, braceHorizontelOffset = 0, braceVerticalOffset = 0;
        private double PI = 0;
        private bool includeCantiPad, includeBracePad, includeBrace, frameSnipToFlge, frameSnipToWeb, isVerticalRoute, isVerticalStruct, slopedSteelY, slopedRoute, slopedSteel;
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
                    overhangOption = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnOHOption", "OverhangOpt")).PropValue;
                    includeCantiPad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnCantiPad", "IncludeCantiPad")).PropValue;
                    uBoltOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnPAttachOff", "Offset")).PropValue;
                    bool uBoltOption = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnPAttachOpt", "PipeAttachOption")).PropValue;
                    bool uBoltOffsetFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnPAttachOff", "OffsetFrmRule")).PropValue;
                    includeBracePad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnBracePad", "IncludeBrPad")).PropValue;
                    includeBrace = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnBrace", "IncludeBrace")).PropValue;
                    frameSnipToFlge = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FrameSniptoFlange")).PropValue;
                    frameSnipToWeb = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "FrameSniptoWeb")).PropValue;

                    routeCount = SupportHelper.SupportedObjects.Count;

                    if (routeCount > 10)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Number of Pipes should be less than or equal to 10", "", "TypeI.cs", 91);
                        return null;
                    }

                    if (overhangOption == 2)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Adjust to Supporting Steel option is not applicable for this support type.", "", "TypeI.cs", 97);
                        return null;
                    }

                    //Get the Section Size
                    sectionSize = MarineAssemblyServices.GetSectionSize(this);
                    PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnSection", "SectionType");
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

                    string tempStr = string.Empty;
                    Array.Resize(ref pipeAttachment, routeCount);
                    Array.Resize(ref uBoltCount, routeCount);
                    Array.Resize(ref isUboltExist, routeCount);

                    Array.Resize(ref routeAngle, routeCount);
                    Array.Resize(ref uBoltPart1, routeCount);
                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        tempStr = "U" + routeIndex;
                        pipeAttachment[routeIndex - 1] = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnPAttach", tempStr)).PropValue;
                        if (routeIndex == 1)
                            uBoltCount[routeIndex - 1] = routeIndex;     //Initialize to 1
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
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe slopes are not equal. Please check", "", "TypeI.cs", 185);
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

                    if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle[0], 3), 0) == false)
                        slopedRoute = true;
                    else if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(structAngle, 3), 0) == false)
                        slopedSteel = true;

                    //Get Pad part and Dimensions

                    string sectionCode = string.Empty, steelStd = string.Empty;
                    padProperties = MarineAssemblyServices.GetPadPartAndDimensions(this, sectionSize, out sectionCode, out steelStd);

                    //get Brace Offset
                    MarineAssemblyServices.GetBraceOffset(this, sectionSize, ref braceHorizontelOffset, ref braceVerticalOffset);

                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    //Get the UboltOffset
                    if (uBoltOffsetFromRule == true)
                    {
                        PartClass hsMarineServiceDim = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_PAttachOffset");
                        IEnumerable<BusinessObject> hsMarineServiceDimPart = hsMarineServiceDim.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        hsMarineServiceDimPart = hsMarineServiceDimPart.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsMrnSrvPAttachOff", "SectionSize")).PropValue == sectionCode));
                        if (hsMarineServiceDimPart.Count() > 0)
                            uBoltOffset = (double)((PropertyValueDouble)hsMarineServiceDimPart.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvPAttachOff", "PipeAttachOffset")).PropValue;
                    }

                    //Create the list of parts
                    string steelStandard = MarineAssemblyServices.GetSteelStandard(this, steelStd);

                    parts.Add(new PartInfo(SECTION, sectionSize + " " + steelStandard));
                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        if (isUboltExist[routeIndex - 1] == true)
                        {
                            Array.Resize(ref uBolt, parts.Count + 2);
                            uBolt[parts.Count + 1] = "sUbolt_" + routeIndex + 1;
                            parts.Add(new PartInfo(uBolt[uBoltCount[routeIndex - 1]], uBoltPart[routeIndex - 1]));
                        }
                    }

                    if (includeCantiPad == true)
                        parts.Add(new PartInfo(CANTIPAD, padProperties.padPart));

                    if (includeBrace == true)
                    {
                        parts.Add(new PartInfo(BRACESECTION, sectionSize + " " + steelStandard));

                        if (includeBracePad == true)
                        {
                            parts.Add(new PartInfo(BRACEPAD, padProperties.padPart));
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
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                double boundingBoxWidth = 0, boundingBoxHeight = 0, leftPipeDiameter = 0, rightPipeDiameter = 0, overhangLeft = 0, cutbackangle = 0;
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
                PropertyValueCodelist braceCodelistcp2 = (PropertyValueCodelist)componentDictionary[SECTION].GetPropertyValue("IJOAhsSteelCP", "CP2");
                componentDictionary[SECTION].SetPropertyValue(braceCodelistcp2.PropValue = 1, "IJOAhsSteelCP", "CP2");
                PropertyValueCodelist braceCodelistcp4 = (PropertyValueCodelist)componentDictionary[SECTION].GetPropertyValue("IJOAhsSteelCP", "CP4");
                componentDictionary[SECTION].SetPropertyValue(braceCodelistcp4.PropValue = 3, "IJOAhsSteelCP", "CP4");

                string verticalSectionPort = string.Empty;
                if (structAngle <= 0)
                    verticalSectionPort = "EndCap";
                else
                    verticalSectionPort = "EndFace";

                textIndex = textIndex + 1;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, componentDictionary[SECTION], verticalSectionPort);

                //Get Section Structure dimensions
                BusinessObject sectionPart = componentDictionary[SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double sectionWidth = 0.0, sectionDepth = 0.0, sectionThickness = 0.0, verticalSectionEndOverLength  = 0;
                int sectionType = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnSection", "SectionType")).PropValue;
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

                //Get the Current Location in the Route Connection Cycle
                string[] partNames = new string[1];
                partNames[0] = SECTION;
                if (includeBrace == true)
                {
                    Array.Resize(ref partNames, 2);
                    partNames[1] = BRACESECTION;
                }

                //get Codelisted Property Values for Hgr Beam
                MarineAssemblyServices.BeamCLProperties beamCLProperties = MarineAssemblyServices.GetBeamCLProperties(componentDictionary);
                //Intialize the Hgr beam properties
                MarineAssemblyServices.IntializeBeamProperties(componentDictionary, support, beamCLProperties, partNames);

                //For L section Initialize Snipped Properties
                if (frameSnipToFlge == true || frameSnipToWeb == true)
                {
                    componentDictionary[SECTION].SetPropertyValue(beamCLProperties.HgrBeamType2.Value, "IJOAHsHgrBeamType", "HgrBeamType");
                    componentDictionary[SECTION].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackBeginAngle1");
                    componentDictionary[SECTION].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackBeginAngle2");
                    componentDictionary[SECTION].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackEndAngle1");
                    componentDictionary[SECTION].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackEndAngle2");
                    componentDictionary[SECTION].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "BeginOffsetAlongFlange");
                    componentDictionary[SECTION].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "BeginOffsetAlongWeb");
                    componentDictionary[SECTION].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "EndOffsetAlongFlange");
                    componentDictionary[SECTION].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "EndOffsetAlongWeb");
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
                            if (Configuration == 1 || Configuration == 4)
                                componentDictionary[SECTION].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackEndAngle");
                            else
                                componentDictionary[SECTION].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackEndAngle");
                        }
                        else
                        {
                            if (Configuration == 1)
                            {
                                componentDictionary[SECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                componentDictionary[SECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                componentDictionary[SECTION].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackEndAngle");
                            }
                            else if (Configuration == 2)
                            {
                                componentDictionary[SECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                componentDictionary[SECTION].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackBeginAngle");
                            }
                            else if (Configuration == 3)
                            {
                                componentDictionary[SECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                componentDictionary[SECTION].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackBeginAngle");
                            }
                            else if (Configuration == 4)
                            {
                                componentDictionary[SECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                componentDictionary[SECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                componentDictionary[SECTION].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackEndAngle");
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
                int[] uBoltTypeCount = new int[routeCount];
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
                            uBoltTypeCount[routeIndex - 1] = (int)((PropertyValueCodelist)hsMarineServiceDimPart.ElementAt(0).GetPropertyValue("IJUAhsMrnSrvPAttachSel", "PipeAttachType")).PropValue;

                        //Get the Ublot Dimensions
                        componentDictionary[uBolt[uBoltCount[routeIndex - 1]]].SetPropertyValue(pipeDiameter[routeIndex - 1], "IJOAhsPipeOD", "PipeOD");
                        if (uBoltTypeCount[routeIndex - 1] == 1)
                        {
                            if (sectionType == 4)
                                componentDictionary[uBolt[uBoltCount[routeIndex - 1]]].SetPropertyValue(sectionWidth, "IJOAhsSteelThickness", "SteelThickness");
                            else
                                componentDictionary[uBolt[uBoltCount[routeIndex - 1]]].SetPropertyValue(sectionThickness, "IJOAhsSteelThickness", "SteelThickness");
                        }

                        support.SetPropertyValue(uBoltTypeCount[routeIndex - 1], "IJOAhsMrnPAttach", tempUBoltType);
                    }
                }

                //Set attribute values for Pads
                string[] arrayOfPadKeys = new string[0];

                if (includeCantiPad == true)
                {
                    Array.Resize(ref arrayOfPadKeys, 1);
                    verticalSectionEndOverLength = verticalSectionEndOverLength - padProperties.padThickness;
                    arrayOfPadKeys[0] = CANTIPAD;
                }

                if (includeBracePad == true)
                {
                    if (includeCantiPad == true)
                    {
                        Array.Resize(ref arrayOfPadKeys, 2);
                        arrayOfPadKeys[1] = BRACEPAD;
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

                //Set the overhang
                if (overhangOption == 1)
                {
                    Collection<object> collectionOverHang;
                    bool value = GenericHelper.GetDataByRule("hsMrnRLFrmOH", (BusinessObject)support, out collectionOverHang);
                    double hangerOverHangLeft = 0;
                    if (collectionOverHang[0] == null)
                        hangerOverHangLeft = (double)collectionOverHang[1];
                    else
                        hangerOverHangLeft = (double)collectionOverHang[0];
                    support.SetPropertyValue(hangerOverHangLeft, "IJOAhsMrnOHLeft", "OverhangLeft");
                    overhangLeft = hangerOverHangLeft - rightPipeDiameter / 2;
                }
                else if (overhangOption == 3)
                {
                    double userDefOverHangLeft = 0;
                    userDefOverHangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHLeft", "OverhangLeft")).PropValue;
                    support.SetPropertyValue(userDefOverHangLeft, "IJOAhsMrnOHLeft", "OverhangLeft");

                    overhangLeft = userDefOverHangLeft - rightPipeDiameter / 2;
                }

                //set the cutback angle
                bool frameCutbackToSection = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSecSnips", "FrameCutbacktoSection")).PropValue;
                double sectionCutbackAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSecSnips", "SecCutbackAngle")).PropValue;
                if (frameCutbackToSection == true)
                    cutbackangle = sectionCutbackAngle;
                else
                    cutbackangle = 0;
                //set the cutback angle for the section
                if (frameSnipToFlge == false && frameSnipToWeb == false)
                {
                    if (!(sectionType == 4))
                    {
                        if (Configuration == 1 || Configuration == 4)
                        {
                            componentDictionary[SECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[SECTION].SetPropertyValue(-cutbackangle, "IJOAhsCutback", "CutbackBeginAngle");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[SECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            componentDictionary[SECTION].SetPropertyValue(cutbackangle, "IJOAhsCutback", "CutbackEndAngle");
                        }
                    }
                }
                else
                {
                    double flgSnipAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FlgSnipAngle")).PropValue;
                    double webSnipAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "WebSnipAngle")).PropValue;
                    double flgSnipOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FlgSnipOffset")).PropValue;
                    double webSnipOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "WebSnipOffset")).PropValue;
                    if (Configuration == 1 || Configuration == 4)
                    {
                        if (frameSnipToFlge == true)
                        {
                            componentDictionary[SECTION].SetPropertyValue(flgSnipAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle1");
                            componentDictionary[SECTION].SetPropertyValue(flgSnipOffset, "IJOAhsCutbackOffset", "BeginOffsetAlongFlange");
                        }
                        if (frameSnipToWeb == true)
                        {
                            componentDictionary[SECTION].SetPropertyValue(webSnipAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle2");
                            componentDictionary[SECTION].SetPropertyValue(webSnipOffset, "IJOAhsCutbackOffset", "BeginOffsetAlongWeb");
                        }
                    }
                    else if (Configuration == 2 || Configuration == 3)
                    {
                        if (frameSnipToFlge == true)
                        {
                            componentDictionary[SECTION].SetPropertyValue(flgSnipAngle, "IJOAhsSnipedSteel", "CutbackEndAngle1");
                            componentDictionary[SECTION].SetPropertyValue(flgSnipOffset, "IJOAhsCutbackOffset", "EndOffsetAlongFlange");
                        }
                        if (frameSnipToWeb == true)
                        {
                            componentDictionary[SECTION].SetPropertyValue(webSnipAngle, "IJOAhsSnipedSteel", "CutbackEndAngle2");
                            componentDictionary[SECTION].SetPropertyValue(webSnipOffset, "IJOAhsCutbackOffset", "EndOffsetAlongWeb");
                        }
                    }
                }

                string routePort = string.Empty,sectionPort = string.Empty,sectionFacePort = string.Empty,padPort = string.Empty,bracePadPort = string.Empty,sectionCapPort = string.Empty,braceCapPort1 = string.Empty,braceCapPort2 = string.Empty,braceFacePort = string.Empty;

                MarineAssemblyServices.ConfigIndex[] routeUboltConfigIndex = new MarineAssemblyServices.ConfigIndex[routeCount];
                MarineAssemblyServices.ConfigIndex routeSectionConfigIndex = new MarineAssemblyServices.ConfigIndex(), padConfigIndex = new MarineAssemblyServices.ConfigIndex();
                double sectionRoutePlaneOffset = 0, sectionRouteAxisOffset = 0, sectionRouteOriginOffset = 0;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                    {
                        if (slopedRoute == true)
                        {
                            if (isVerticalRoute == true)
                                routePort = "BBSV_Low";
                            else
                            {
                                if (isVerticalStruct == true)
                                    routePort = "BBSV_Low";
                                else
                                    routePort = "BBSR_Low";
                            }
                        }
                        else if (slopedSteel == true)
                            routePort = "BBSR_Low";
                        else
                            routePort = "BBSR_Low";
                    }
                    else if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (slopedRoute == true)
                            routePort = "BBR_Low";
                        else if (slopedSteel == true)
                        {
                            if (isVerticalRoute == true)
                                routePort = "BBSR_Low";
                            else
                            {
                                if (isVerticalStruct == true)
                                    routePort = "BBSR_Low";
                                else
                                    routePort = "BBSV_Low";
                            }
                        }
                        else
                            routePort = "BBR_Low";
                    }
                }
                else
                {
                    if (slopedRoute == true)
                    {
                        if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                        {
                            if (isVerticalRoute == true)
                                routePort = "BBR_Low";
                            else
                                routePort = "BBRV_Low";
                        }
                        else
                            routePort = "BBR_Low";
                    }
                    else if (slopedSteel == true)
                    {
                        if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                        {
                            MarineAssemblyServices.CreateBoundingBox(this, "MarineBBX", false);            //for a route which is not vertical
                            routePort = "MarineBBX_Low";
                        }
                        else
                            routePort = "BBRV_Low";
                    }
                    else
                    {
                        if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                        {
                            if (isVerticalRoute == true)
                            {
                                MarineAssemblyServices.CreateBoundingBox(this, "MarineBBX", isVerticalRoute);       //for a route which is  vertical
                                routePort = "MarineBBX_Low";
                            }
                            else
                                routePort = "BBR_Low";
                        }
                        else
                            routePort = "BBR_Low";
                    }
                }

                for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                {
                    if (isUboltExist[routeIndex - 1] == true)
                    {
                        if (sectionType == 4)
                            routeUboltConfigIndex[routeIndex - 1] = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeZX, Axis.Y, Axis.NegativeZ);
                        else
                            routeUboltConfigIndex[routeIndex - 1] = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeZ);
                    }
                }

                if (Configuration == 1 || Configuration == 4)
                {
                    sectionPort = "BeginCap";
                    sectionCapPort = "EndCap";
                    sectionFacePort = "EndFace";
                    padPort = "Port1";
                }
                else if (Configuration == 2 || Configuration == 3)
                {
                    sectionPort = "EndCap";
                    sectionCapPort = "BeginCap";
                    sectionFacePort = "BeginFace";
                    padPort = "Port2";
                }

                double shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnShoeHeight", "ShoeHeight")).PropValue;
                if (sectionType == 4)
                {
                    sectionPort = "BeginCap";
                    sectionCapPort = "EndCap";
                    sectionFacePort = "EndFace";
                    padPort = "Port1";
                    padConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.Y, Axis.NegativeY);
                  
                    if (Configuration == 1 || Configuration == 4)
                    {
                        routeSectionConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeY);
                        sectionRouteOriginOffset = -shoeHeight;
                        sectionRoutePlaneOffset = overhangLeft;
                        sectionRouteAxisOffset = 0;
                    }
                    else
                    {
                        routeSectionConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.Y);
                        sectionRouteOriginOffset = -boundingBoxWidth - shoeHeight;
                        sectionRoutePlaneOffset = overhangLeft;
                        sectionRouteAxisOffset = 0;
                    }
                }
                else
                {
                    padConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeY);

                    if (Configuration == 1)
                    {
                        routeSectionConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeZX, Axis.X, Axis.NegativeX);
                        sectionRouteOriginOffset = sectionWidth / 2;
                        sectionRoutePlaneOffset = -shoeHeight;
                        sectionRouteAxisOffset = overhangLeft;
                    }
                    else if (Configuration == 2)
                    {
                        routeSectionConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeZX, Axis.X, Axis.X);
                        sectionRouteOriginOffset = sectionWidth / 2;
                        sectionRoutePlaneOffset = -shoeHeight;
                        sectionRouteAxisOffset = -overhangLeft;
                    }
                    else if (Configuration == 3)
                    {
                        routeSectionConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX);
                        sectionRouteOriginOffset = sectionWidth / 2;
                        sectionRoutePlaneOffset = -boundingBoxWidth - shoeHeight;
                        sectionRouteAxisOffset = -overhangLeft;
                    }
                    else if (Configuration == 4)
                    {
                        routeSectionConfigIndex = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.ZX, Axis.X, Axis.X);
                        sectionRouteOriginOffset = sectionWidth / 2;
                        sectionRoutePlaneOffset = -boundingBoxWidth - shoeHeight;
                        sectionRouteAxisOffset = overhangLeft;
                    }
                }

                //Add the Joint for the Pad
                if (includeCantiPad == true)
                    //Add Joint Between the Plate and the Beam
                    JointHelper.CreateRigidJoint(CANTIPAD, padPort, SECTION, sectionFacePort, padConfigIndex.A, padConfigIndex.B, padConfigIndex.C, padConfigIndex.D, 0, 0, 0);

                //Set the Over Length for Snipped Steel for Sloped Structure
                double overLength = 0;
                if (slopedSteel == true)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (frameSnipToFlge == true || frameSnipToWeb == true)
                        {
                            overLength = sectionDepth * Math.Tan(structAngle);
                            if (structAngle < 0)       //Negative Angle
                            {
                                if (Configuration == 3 || Configuration == 4)
                                    verticalSectionEndOverLength = verticalSectionEndOverLength - overLength;
                            }
                            else            //Positive Angle
                            {
                                if (Configuration == 1 || Configuration == 2)
                                    verticalSectionEndOverLength = verticalSectionEndOverLength + overLength;
                            }
                        }

                        if (supportingType == "Slab")
                        {
                            if (frameSnipToFlge == false || frameSnipToWeb == false)
                            {
                                overLength = sectionDepth * Math.Tan(structAngle);
                                if (structAngle < 0)       //Negative Angle
                                {
                                    if (slopedSteelY == true)
                                    {
                                        if (Configuration == 1 || Configuration == 2)
                                            verticalSectionEndOverLength = verticalSectionEndOverLength - overLength;
                                    }
                                    else
                                    {
                                        if (Configuration == 2 || Configuration == 4)
                                            verticalSectionEndOverLength = verticalSectionEndOverLength - overLength;
                                    }
                                }
                                else            //Positive Angle
                                {
                                    if (slopedSteelY == true)
                                    {
                                        if (Configuration == 3 || Configuration == 4)
                                            verticalSectionEndOverLength = verticalSectionEndOverLength + overLength;
                                    }
                                    else
                                    {
                                        if (Configuration == 1 || Configuration == 3)
                                            verticalSectionEndOverLength = verticalSectionEndOverLength + overLength;
                                    }
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
                            if (routeAngle[1] < 0)     //Negative Angle
                            {
                                if (Configuration == 1 || Configuration == 3)
                                    verticalSectionEndOverLength = verticalSectionEndOverLength - overLength;
                            }
                            else            //Positive Angle
                            {
                                if (Configuration == 2 || Configuration == 4)
                                    verticalSectionEndOverLength = verticalSectionEndOverLength + overLength;
                            }
                        }
                    }
                }

                if (Configuration == 1 || Configuration == 4 || sectionType == 4)
                    componentDictionary[SECTION].SetPropertyValue(verticalSectionEndOverLength, "IJUAHgrOccOverLength", "EndOverLength");
                else
                    componentDictionary[SECTION].SetPropertyValue(verticalSectionEndOverLength, "IJUAHgrOccOverLength", "BeginOverLength");

                //Add joint Between Section and BoundingBox
                JointHelper.CreateRigidJoint(SECTION, sectionPort, "-1", routePort, routeSectionConfigIndex.A, routeSectionConfigIndex.B, routeSectionConfigIndex.C, routeSectionConfigIndex.D, sectionRoutePlaneOffset, sectionRouteAxisOffset, sectionRouteOriginOffset);

                //Add Joint Between the Ports of Vertical Beam
                JointHelper.CreatePrismaticJoint(SECTION, "BeginCap", SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add Joint Between the Ports of Vertical Beam
                JointHelper.CreatePointOnPlaneJoint(SECTION, sectionCapPort, "-1", "Structure", Plane.XY);

                double braceAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnBraceAngle", "BraceAngle")).PropValue;
                if (includeBrace == true)
                {
                    double braceVerticalOffset = 0, braceHorizontalOffset = 0, braceAxisOffset = 0, angleX = 0, angleY = 0, angleZ = 0, angleLength = 0;
                    string frameLength = string.Empty;

                    if (sectionType == 4)
                        angleLength = (sectionDepth / Math.Sin(braceAngle)) + boundingBoxHeight + overhangLeft;
                    else
                        angleLength = (sectionWidth / Math.Sin(braceAngle)) + boundingBoxHeight + overhangLeft;

                    if (Configuration == 2 && Configuration == 3)
                    {
                        braceCapPort1 = "BeginCap";
                        braceCapPort2 = "EndCap";
                        braceFacePort = "EndFace";
                        frameLength = "EndOverLength";
                        bracePadPort = "Port1";
                    }
                    else
                    {
                        braceCapPort1 = "EndCap";
                        braceCapPort2 = "BeginCap";
                        braceFacePort = "BeginFace";
                        frameLength = "BeginOverLength";
                        bracePadPort = "Port2";
                    }

                    if (sectionType == 4)
                    {
                        componentDictionary[BRACESECTION].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                        componentDictionary[BRACESECTION].SetPropertyValue(-braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                        componentDictionary[BRACESECTION].SetPropertyValue((PI / 2 - braceAngle), "IJOAhsCutback", "CutbackEndAngle");
                        angleX = -(PI - braceAngle);
                        angleZ = -PI;
                        braceHorizontalOffset = -sectionWidth;
                        braceVerticalOffset = braceVerticalOffset + angleLength;
                        braceAxisOffset = braceHorizontelOffset;
                    }
                    else
                    {
                        braceHorizontalOffset = braceHorizontelOffset;
                        if (Configuration == 1 || Configuration == 4)
                        {
                            componentDictionary[BRACESECTION].SetPropertyValue(-braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                            componentDictionary[BRACESECTION].SetPropertyValue((PI / 2 - braceAngle), "IJOAhsCutback", "CutbackEndAngle");
                            braceVerticalOffset = braceVerticalOffset + angleLength;
                            angleY = -(PI - braceAngle);
                            angleZ = -PI;
                        }
                        else
                        {
                            componentDictionary[BRACESECTION].SetPropertyValue(-braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                            componentDictionary[BRACESECTION].SetPropertyValue((PI / 2 - braceAngle), "IJOAhsCutback", "CutbackEndAngle");
                            braceVerticalOffset = -braceVerticalOffset - angleLength;
                            angleY = -(PI - braceAngle);
                            angleZ = 0;
                        }
                    }
                    //Add Joint Between Brace Section and Section
                    JointHelper.CreateAngularRigidJoint(BRACESECTION, braceCapPort1, SECTION, sectionPort, new Vector(braceHorizontalOffset, braceAxisOffset, braceVerticalOffset), new Vector(angleX, angleY, angleZ));

                    //Add Joint Between the Ports of  Brace section
                    JointHelper.CreatePrismaticJoint(BRACESECTION, "BeginCap", BRACESECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    //Add Joint Between the Ports of Bracesection
                    JointHelper.CreatePointOnPlaneJoint(BRACESECTION, braceCapPort2, "-1", "Structure", Plane.XY);

                    if (includeBracePad == true)
                    {
                        componentDictionary[BRACESECTION].SetPropertyValue(-padProperties.padThickness / Math.Cos(braceAngle), "IJUAHgrOccOverLength", frameLength);

                        //Add joint Between BracePad and Brace section
                        JointHelper.CreateRigidJoint(BRACEPAD, bracePadPort, BRACESECTION, braceFacePort, padConfigIndex.A, padConfigIndex.B, padConfigIndex.C, padConfigIndex.D, 0, 0, 0);
                    }
                }

                //Joint for Ubolt
                string strRefPortName = string.Empty;
                for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                {
                    if (isUboltExist[routeIndex - 1] == true)
                    {
                        if (routeIndex == 1)
                            strRefPortName = "Route";
                        else
                            strRefPortName = "Route_" + routeIndex;

                        //Add Joint Between the UBolt and Route
                        JointHelper.CreateTranslationalJoint(uBolt[uBoltCount[routeIndex - 1]], "Route", SECTION, "Neutral", routeUboltConfigIndex[routeIndex - 1].A, routeUboltConfigIndex[routeIndex - 1].B, routeUboltConfigIndex[routeIndex - 1].C, routeUboltConfigIndex[routeIndex - 1].D, 0);

                        //Add Joint Between the UBolt and Route
                        JointHelper.CreatePointOnAxisJoint(uBolt[uBoltCount[routeIndex - 1]], "Route", "-1", strRefPortName, Axis.X);
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
                    routeConnections.Add(new ConnectionInfo(SECTION, 1)); // partindex, routeindex

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
                    structConnections.Add(new ConnectionInfo(SECTION, 1)); // partindex, routeindex

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

