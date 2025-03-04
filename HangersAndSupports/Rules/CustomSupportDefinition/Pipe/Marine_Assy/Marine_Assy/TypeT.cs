//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   TypeT.cs
//   Marine_Assy,Ingr.SP3D.Content.Support.Rules.TypeT
//   Author       :Vijaya
//   Creation Date:23.July.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  1.Aug.2013     Vijaya    CR-CP-224470 Convert HS_Marine_Assy to C# .Net
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
    public class TypeT : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string HORSECTION1 = "HORSECTION1";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string LEFTPAD = "LEFTPAD";
        private string[] ubolts = new string[10], uBoltPart1;
        private double PI = 0;
        string sectionSize = string.Empty, supportingType = string.Empty;
        bool includeLeftPad, frameSnipToFlge, frameSnipToWeb, isVerticalRoute, isVerticalStruct, slopedSteel, slopedRoute, slopedSteelY;
        double structAngle = 0.0, uBoltOffset = 0.0;
        int routeCount = 0;
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
                    includeLeftPad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnLeftPad", "IncludeLeftPad")).PropValue;
                    bool uBoltOption = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnPAttachOpt", "PipeAttachOption")).PropValue;
                    bool uBoltOffSetFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnPAttachOff", "OffsetFrmRule")).PropValue;
                    uBoltOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnPAttachOff", "Offset")).PropValue;
                    connectionCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnFrmConn", "Connection");
                    PropertyValueCodelist padShapeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnPadShape", "PadShape");
                    frameSnipToFlge = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FrameSniptoFlange")).PropValue;
                    frameSnipToWeb = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "FrameSniptoWeb")).PropValue;
                    
                    routeCount = SupportHelper.SupportedObjects.Count;
                    //Initialize for the existing symbol
                    if (frameSnipToFlge != true && frameSnipToFlge != false)
                        frameSnipToFlge = false;

                    if (frameSnipToWeb != true && frameSnipToWeb != false)
                        frameSnipToWeb = false;

                    if (routeCount > 10)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Number of Pipes should be less than or equal to 10", "", "TypeT.cs", 87);
                        return null;
                    }

                    if (connectionCodeList.PropValue == 3)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "This Connection Type is not appplicable for this assembly", "", "TypeT.cs", 93);
                        return null;
                    }

                    //Get the Section Size
                    sectionSize = MarineAssemblyServices.GetSectionSize(this);

                    //Set the Snip Angle
                    if (frameSnipToFlge == true || frameSnipToWeb == true)
                    {
                        if (sectionTypeCodeList.PropValue != 1)    //Snipped Steel is only applicable for L Section
                        {
                            frameSnipToFlge = false;
                            frameSnipToWeb = false;
                        }
                    }
                    //Get the UBolt
                    string[] uBoltPart = MarineAssemblyServices.GetUboltPart(this);

                    string[] PipeAttachment = new string[routeCount];
                    uBoltPart1 = new string[routeCount];
                    isUboltExist = new bool[routeCount];
                    for (int index = 1; index <= routeCount; index++)
                    {
                        PipeAttachment[index - 1] = "U" + index;

                        if (index == 1)
                            ubolts[index - 1] = "UBOLT_" + index;
                        else
                            ubolts[index - 1] = "UBOLT_" + index;

                        if (uBoltOption == false)        //User defined
                            if (PipeAttachment[index - 1] == "3")   //None
                                isUboltExist[index - 1] = false;
                            else
                            {
                                ubolts[index - 1] = "UBOLT_" + (index);
                                isUboltExist[index - 1] = true;
                            }

                        else            //Default Ubolt is True
                        {
                            ubolts[index - 1] = "UBOLT_" + (index);
                            isUboltExist[index - 1] = true;
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

                    routeAngle = MarineAssemblyServices.GetRoutAngleAndIsRouteVertical(this, routeCount, out isVerticalRoute);

                    if (routeCount > 1)
                    {
                        for (int index = 1; index < routeCount; index++)
                        {
                            if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle[index - 1], 3), Math.Round(routeAngle[index], 3))== false)
                            {
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe slopes are not equal. Please check", "", "TypeT.cs", 157);
                                return null;
                            }
                        }
                    }
                    //CommonAssembly commonAssembly = new CommonAssembly();
                    supportingType = MarineAssemblyServices.GetSupportingTypes(this);
                    //Checking the angles of Stucture X and Stucture Y axes with Global Z axis
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

                    if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle[0], 3), 0) == false) //for sloped Route
                        slopedRoute = true;
                    else if (Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(structAngle, 3), 0) == false)    //for sloped steel
                        slopedSteel = true;

                    //Get Pad part and Dimensions
                    string sectionCode, steelStd;
                    padProperties = MarineAssemblyServices.GetPadPartAndDimensions(this, sectionSize, out sectionCode, out steelStd);

                    //Get the UboltOffset
                    if (uBoltOffSetFromRule == true)
                    {
                        CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                        ReadOnlyCollection<BusinessObject> classItems;
                        PartClass auxilaryTable = (PartClass)catalogBaseHelper.GetPartClass("hsMrnSrv_PAttachOffset");
                        classItems = auxilaryTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                        foreach (BusinessObject classItem in classItems)
                        {
                            if (((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsMrnSrvPAttachOff", "SectionSize")).PropValue == sectionCode))
                            {
                                uBoltOffset = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsMrnSrvPAttachOff", "PipeAttachOffset")).PropValue;
                                break;
                            }
                        }
                    }

                    //Create the list of parts                  
                    //GetSteelStandard    
                    string steelStandard = MarineAssemblyServices.GetSteelStandard(this, steelStd);

                    parts.Add(new PartInfo(HORSECTION1, sectionSize + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION1, sectionSize + " " + steelStandard));

                    for (int routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                        if (isUboltExist[routeIndex - 1] == true)
                            parts.Add(new PartInfo(ubolts[routeIndex - 1], uBoltPart[routeIndex - 1]));

                    if (includeLeftPad == true)
                        parts.Add(new PartInfo(LEFTPAD, padProperties.padPart));

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
                return 2;
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
                //load standard bounding box.
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                //BoundingBox boundingBox;
                double boundingBoxWidth = 0.0, boundingBoxHeight = 0.0, leftPipeDiameter = 0.0, rightPipeDiameter = 0.0;
                MarineAssemblyServices.GetBoundingBoxDimensionsAndPipeDiameter(this, routeCount, ref boundingBoxWidth, ref boundingBoxHeight, ref leftPipeDiameter, ref rightPipeDiameter);

                //Auto Dimensioning of Supports
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                int routeIndex = 0;
                for (routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    if (isUboltExist[routeIndex - 1] == true)
                        MarineAssemblyServices.CreateDimensionNote(this, "Dim " + routeIndex, componentDictionary[ubolts[routeIndex - 1]], "Route");

                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + routeIndex, componentDictionary[HORSECTION1], "BeginCap");
                routeIndex = routeIndex + 1;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + routeIndex, componentDictionary[HORSECTION1], "EndCap");

                PropertyValueCodelist cardinalCodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJOAhsSteelCP", "CP2");
                cardinalCodeList.PropValue = 1;
                componentDictionary[VERTSECTION1].SetPropertyValue(cardinalCodeList.PropValue, "IJOAhsSteelCP", "CP2");
                cardinalCodeList = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJOAhsSteelCP", "CP4");
                cardinalCodeList.PropValue = 3;
                componentDictionary[VERTSECTION1].SetPropertyValue(cardinalCodeList.PropValue, "IJOAhsSteelCP", "CP4");

                string verticalSectionPort = string.Empty;
                if (structAngle <= 0.0)
                    verticalSectionPort = "EndCap";
                else
                    verticalSectionPort = "EndFace";

                routeIndex = routeIndex + 1;
                MarineAssemblyServices.CreateDimensionNote(this, "Dim " + routeIndex, componentDictionary[VERTSECTION1], verticalSectionPort);
                //Get Section Structure dimensions
                BusinessObject sectionPart = componentDictionary[HORSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                //Get SteelWidth and SteelDepth 
                double sectionWidth = 0.0, sectionDepth = 0.0, steelThickness = 0.0;//, vertSectionWidth = 0.0, vertSectionDepth = 0.0, vertSteelThickness = 0.0;
                sectionWidth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                sectionDepth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                if (sectionTypeCodeList.PropValue != 4)
                    steelThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                double dShoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnShoeHeight", "ShoeHeight")).PropValue;

                //Get the Current Location in the Route Connection Cycle
                string[] partNames = new string[2];
                partNames[0] = HORSECTION1;
                partNames[1] = VERTSECTION1;

                //get Codelisted Property Values for Hgr Beam
                MarineAssemblyServices.BeamCLProperties beamCLProperties = MarineAssemblyServices.GetBeamCLProperties(componentDictionary);
                //set HGR Beam Properties.
                MarineAssemblyServices.IntializeBeamProperties(componentDictionary, support, beamCLProperties, partNames);
                // Set Values of Part Occurance Attributes
                //For L section Change to Snipped type

                if (frameSnipToFlge == true || frameSnipToWeb == true)
                {
                    componentDictionary[HORSECTION1].SetPropertyValue(beamCLProperties.HgrBeamType2.Value, "IJOAHsHgrBeamType", "HgrBeamType");
                    componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackBeginAngle1");
                    componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackBeginAngle2");
                    componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackEndAngle1");
                    componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsSnipedSteel", "CutbackEndAngle2");
                    componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "BeginOffsetAlongFlange");
                    componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "BeginOffsetAlongWeb");
                    componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "EndOffsetAlongFlange");
                    componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsCutbackOffset", "EndOffsetAlongWeb");
                }

                //set the overlength for the vertcal section
                PropertyValueCodelist orientCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsMrnFrmOrient", "FrmOrient");
                string fromOrient = orientCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(orientCodeList.PropValue).DisplayName;
                if (slopedSteel == true && SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (sectionTypeCodeList.PropValue == 4)
                            componentDictionary[VERTSECTION1].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackEndAngle");
                        else
                        {
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                            componentDictionary[VERTSECTION1].SetPropertyValue(beamCLProperties.VertSecCBAncPt2AlWeb.Value, "IJOAhsCutback", "EndCutbackAnchorPoint");
                            if (Configuration == 1)
                            {
                                if (connectionCodeList.PropValue == 1)
                                    componentDictionary[VERTSECTION1].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackEndAngle");
                                else if (connectionCodeList.PropValue == 2)
                                    componentDictionary[VERTSECTION1].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackEndAngle");
                            }
                            else if (Configuration == 2)
                            {

                                if (connectionCodeList.PropValue == 1)
                                    componentDictionary[VERTSECTION1].SetPropertyValue(structAngle, "IJOAhsCutback", "CutbackEndAngle");
                                else if (connectionCodeList.PropValue == 2)
                                    componentDictionary[VERTSECTION1].SetPropertyValue(-structAngle, "IJOAhsCutback", "CutbackEndAngle");
                            }

                        }
                    }
                }

                if (slopedRoute == true && SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (Configuration == 1)
                        {
                            if (connectionCodeList.PropValue == 1)
                                componentDictionary[VERTSECTION1].SetPropertyValue(-routeAngle[0], "IJOAhsCutback", "CutbackEndAngle");
                            else if (connectionCodeList.PropValue == 2)
                                componentDictionary[VERTSECTION1].SetPropertyValue(routeAngle[0], "IJOAhsCutback", "CutbackEndAngle");
                        }
                        else if (Configuration == 2)
                        {
                            if (connectionCodeList.PropValue == 1)
                                componentDictionary[VERTSECTION1].SetPropertyValue(routeAngle[0], "IJOAhsCutback", "CutbackEndAngle");
                            else if (connectionCodeList.PropValue == 2)
                                componentDictionary[VERTSECTION1].SetPropertyValue(-routeAngle[0], "IJOAhsCutback", "CutbackEndAngle");
                        }
                    }
                }

                //Set the Snipped angle for L section
                double flgSnipAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FlgSnipAngle")).PropValue;
                double webSnipAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "WebSnipAngle")).PropValue;
                double flgSnipOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongFlg", "FlgSnipOffset")).PropValue;
                double webSnipOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnSnipAlongWeb", "WebSnipOffset")).PropValue;

                if (frameSnipToFlge == true || frameSnipToWeb == true)
                {
                    if (frameSnipToFlge == true)
                    {
                        componentDictionary[HORSECTION1].SetPropertyValue(flgSnipAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle1");
                        componentDictionary[HORSECTION1].SetPropertyValue(flgSnipAngle, "IJOAhsSnipedSteel", "CutbackEndAngle1");
                        componentDictionary[HORSECTION1].SetPropertyValue(flgSnipOffset, "IJOAhsCutbackOffset", "BeginOffsetAlongFlange");
                        componentDictionary[HORSECTION1].SetPropertyValue(flgSnipOffset, "IJOAhsCutbackOffset", "EndOffsetAlongFlange");
                    }
                    if (frameSnipToWeb == true)
                    {
                        componentDictionary[HORSECTION1].SetPropertyValue(webSnipAngle, "IJOAhsSnipedSteel", "CutbackBeginAngle2");
                        componentDictionary[HORSECTION1].SetPropertyValue(webSnipAngle, "IJOAhsSnipedSteel", "CutbackEndAngle2");
                        componentDictionary[HORSECTION1].SetPropertyValue(webSnipOffset, "IJOAhsCutbackOffset", "BeginOffsetAlongWeb");
                        componentDictionary[HORSECTION1].SetPropertyValue(webSnipOffset, "IJOAhsCutbackOffset", "EndOffsetAlongWeb");
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
                        componentDictionary[ubolts[indxRoute - 1]].SetPropertyValue(pipeDiameter[indxRoute - 1], "IJOAhsPipeOD", "PipeOD");
                        if (uBoltTypes[indxRoute - 1] == 1)
                        {
                            if (sectionTypeCodeList.PropValue == 4)
                                componentDictionary[ubolts[indxRoute - 1]].SetPropertyValue(sectionWidth, "IJOAhsSteelThickness", "SteelThickness");
                            else
                                componentDictionary[ubolts[indxRoute - 1]].SetPropertyValue(steelThickness, "IJOAhsSteelThickness", "SteelThickness");
                        }
                        support.SetPropertyValue(uBoltTypes[indxRoute - 1], "IJOAhsMrnPAttach", tempUBoltType);
                    }
                }

                //Set attribute values for Pads
                double verticalSectionEndOverLength = 0.0;

                string[] padPartNames = new string[1];
                if (includeLeftPad == true)
                {
                    verticalSectionEndOverLength = verticalSectionEndOverLength - padProperties.padThickness;
                    padPartNames[0] = LEFTPAD;
                }
                MarineAssemblyServices.SetPadProperties(componentDictionary, support, padProperties, padPartNames);

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
                
                Collection<object> overHangCollection = new Collection<object>();
                GenericHelper.GetDataByRule("hsMrnRLFrmOH", (BusinessObject)support, out overHangCollection);
                if (overHangCollection[0] == null)
                {
                    hgrOverHangLeft = (double)overHangCollection[1];
                    hgrOverHangRight = (double)overHangCollection[2];
                }
                else
                {
                    hgrOverHangLeft = (double)overHangCollection[0];
                    hgrOverHangRight = (double)overHangCollection[1];
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

                //Set Length of horzontal member
                double horSectionLength = 0.0, routeLowStruct2HorDist, routeHighStruct2HorDist;
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    double tempDistance = 0.0;
                    if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 3)      //By Catalog Ruleand user defined
                        tempDistance = boundingBoxWidth;
                    else if (overHangCodeList.PropValue == 2)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Joints" + ": " + "WARNING: " + "Adjust to Supporting steel is applicable only when the support placement is By Point.", "", "TypeT.cs", 520);
                        return;
                    }

                    horSectionLength = tempDistance + overHangLeft + overHangRight;
                    componentDictionary[HORSECTION1].SetPropertyValue(horSectionLength, "IJUAHgrOccLength", "Length");
                }
                else                        //for Place by point
                {
                    if ((Math.Abs(routeAngle[0]) < (0 + 0.01) && Math.Abs(routeAngle[0]) > (0 - 0.01)) || (Math.Abs(routeAngle[0]) < (PI + 0.01) && Math.Abs(routeAngle[0]) > (PI - 0.01)))
                    {
                        routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortDistanceType.Horizontal);
                        routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", "Structure", PortDistanceType.Horizontal);
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
                    if (supportingType.ToUpper() == "SLAB" || supportingType.ToUpper() == "WALL")
                    {
                        if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 3)          //By catalog OR User Defined
                            horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                        else if (overHangCodeList.PropValue == 2)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Joints" + ": " + "WARNING: " + "Adjust to Supporting steel is applicable only when the support placement is By Point and the supporting count is more than one.", "", "TypeT.cs", 558);
                            return;
                        }
                    }
                    else if (supportingType.ToUpper() == "STEEL")
                    {
                        if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 3)         //By catalog
                            horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                        else if (overHangCodeList.PropValue == 2)
                        {
                            overHangRight = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist) + sectionWidth / 2;
                            overHangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHLeft", "OverhangLeft")).PropValue;
                            horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                            support.SetPropertyValue(overHangRight, "IJOAhsMrnOHRight", "OverhangRight");
                        }
                    }
                    else if (supportingType.ToUpper() == "STEEL-SLAB" || supportingType.ToUpper() == "SLAB-STEEL")
                    {
                        if (overHangCodeList.PropValue == 1 || overHangCodeList.PropValue == 3)         //by catalog
                            horSectionLength = boundingBoxWidth + overHangLeft + overHangRight;
                        else if (overHangCodeList.PropValue == 2)
                        {
                            overHangRight = Math.Min(routeLowStruct2HorDist, routeHighStruct2HorDist);
                            overHangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsMrnOHLeft", "OverhangLeft")).PropValue;
                            horSectionLength = boundingBoxWidth + overHangLeft + overHangRight + sectionWidth;
                        }
                    }
                    componentDictionary[HORSECTION1].SetPropertyValue(horSectionLength, "IJUAHgrOccLength", "Length");
                }

                Double dOverLength = 0.0;
                if (slopedSteel == true && supportingType.ToUpper() == "SLAB")
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    {
                        if (frameSnipToFlge == false || frameSnipToWeb == false)
                        {
                            dOverLength = sectionDepth * Math.Tan(structAngle);
                            if (structAngle < 0)   //Negative Angle
                            {
                                if (slopedSteelY == true)
                                {
                                    if (Configuration == 1)
                                        verticalSectionEndOverLength = verticalSectionEndOverLength - dOverLength;
                                }
                                else
                                {
                                    if (Configuration == 2)
                                        verticalSectionEndOverLength = verticalSectionEndOverLength - dOverLength;
                                }
                            }

                            else                            //Positive Angle
                            {
                                if (slopedSteelY == true)
                                {
                                    if (Configuration == 2)
                                        verticalSectionEndOverLength = verticalSectionEndOverLength + dOverLength;
                                }
                                else
                                {
                                    if (Configuration == 1)
                                        verticalSectionEndOverLength = verticalSectionEndOverLength + dOverLength;
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
                            dOverLength = sectionDepth * Math.Tan(routeAngle[0]);
                            if (routeAngle[0] < 0)        //Negative Angle
                            {
                                if (Configuration == 2)
                                    verticalSectionEndOverLength = verticalSectionEndOverLength - dOverLength;
                            }

                            else                            //Positive Angle
                            {
                                if (Configuration == 1)
                                    verticalSectionEndOverLength = verticalSectionEndOverLength + dOverLength;
                            }
                        }
                    }
                }
                componentDictionary[VERTSECTION1].SetPropertyValue(verticalSectionEndOverLength, "IJUAHgrOccOverLength", "EndOverLength");

                //Create Joints
                double horizontalVertOriginOffset = 0.0, horizontalVertAxisOffset = 0.0;
                if (connectionCodeList.PropValue == 1)  //Mitered Joint
                {
                    horizontalVertOriginOffset = 0.0;
                    if (Configuration == 1)
                        horizontalVertAxisOffset = -horSectionLength / 2 + sectionDepth / 2;
                    else if (Configuration == 2)
                        horizontalVertAxisOffset = horSectionLength / 2 + sectionDepth / 2;

                }
                else if (connectionCodeList.PropValue == 2)   //Wrapping joint
                {
                    componentDictionary[VERTSECTION1].SetPropertyValue(PI, "IJOAhsBeginCap", "BeginCapRotZ");
                    horizontalVertOriginOffset = 0.0;
                    if (Configuration == 1)
                        horizontalVertAxisOffset = -horSectionLength / 2 - sectionDepth / 2;
                    else if (Configuration == 2)
                        horizontalVertAxisOffset = horSectionLength / 2 - sectionDepth / 2;
                }

                //Get Route Port
                string routePort = MarineAssemblyServices.GetRoutePort(this, support, slopedRoute, slopedSteel, isVerticalRoute, isVerticalStruct);
                string horizonatlSectionPort = string.Empty, verticalSectionPort1 = string.Empty;

                if (Configuration == 1)
                {
                    horizonatlSectionPort = "EndCap";
                    verticalSectionPort1 = "BeginCap";
                }
                else if (Configuration == 2)
                {
                    horizonatlSectionPort = "BeginCap";
                    verticalSectionPort1 = "BeginCap";
                }

                MarineAssemblyServices.ConfigIndex padPlane = new MarineAssemblyServices.ConfigIndex(), routeHorPlane = new MarineAssemblyServices.ConfigIndex(), horVertPlane = new MarineAssemblyServices.ConfigIndex();
                MarineAssemblyServices.ConfigIndex[] routeUboltPlane = new MarineAssemblyServices.ConfigIndex[routeCount];


                for (int idexRoute = 1; idexRoute <= routeCount; idexRoute++)
                {
                    if (isUboltExist[idexRoute - 1] == true)
                    {
                        if (sectionTypeCodeList.PropValue != 4)
                            routeUboltPlane[idexRoute - 1] = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeZ);
                        else
                            routeUboltPlane[idexRoute - 1] = new MarineAssemblyServices.ConfigIndex(Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ);
                    }
                }

                string horizotalSectionPort1 = string.Empty;
                double horRoutePlaneOffset = 0.0, horizontalRouteOriginOffset = 0.0, horizontalRouteAxisOffset = 0.0, horizontalVertPlaneOffset = 0.0, boltRoutePlaneOffset = 0.0;
                if (sectionTypeCodeList.PropValue == 4)
                {
                    horizotalSectionPort1 = "EndCap";
                    horizonatlSectionPort = "EndCap";

                    padPlane = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.Y, Axis.NegativeY);
                    routeHorPlane = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.YZ, Axis.X, Axis.NegativeZ);
                    horVertPlane = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeZX, Axis.X, Axis.NegativeZ);

                    verticalSectionPort1 = "BeginCap";
                    horRoutePlaneOffset = sectionDepth / 2;
                    horizontalRouteOriginOffset = sectionWidth + dShoeHeight;
                    horizontalRouteAxisOffset = -overHangRight - boundingBoxWidth;
                    horizontalVertAxisOffset = -horSectionLength / 2 + sectionWidth / 2;
                    horizontalVertOriginOffset = sectionWidth;
                    horizontalVertPlaneOffset = sectionDepth;
                    boltRoutePlaneOffset = sectionDepth / 2 - uBoltOffset;
                }
                else
                {
                    horizontalRouteOriginOffset = sectionWidth / 2;
                    boltRoutePlaneOffset = sectionWidth / 2 - uBoltOffset;
                    
                    padPlane = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeY);

                    if (Configuration == 1)
                    {
                        horVertPlane = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        if (routePort == "BBR_High" || routePort == "BBSR_High")
                        {
                            routeHorPlane = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                            horizotalSectionPort1 = "EndCap";
                            horRoutePlaneOffset = -dShoeHeight;
                            horizontalRouteAxisOffset = -overHangRight - boundingBoxWidth;
                        }
                        else
                        {
                            routeHorPlane = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX);
                            horizotalSectionPort1 = "BeginCap";
                            horRoutePlaneOffset = overHangRight + boundingBoxWidth;
                            horizontalRouteAxisOffset = -dShoeHeight;
                        }
                    }
                    else if (Configuration == 2)
                    {
                        horVertPlane = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        if (routePort == "BBR_High" || routePort == "BBSR_High")
                        {
                            routeHorPlane = new MarineAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX);
                            horizotalSectionPort1 = "BeginCap";
                            horRoutePlaneOffset = overHangRight + boundingBoxWidth;
                            horizontalRouteAxisOffset = -dShoeHeight;
                        }
                        else
                        {
                            routeHorPlane = new MarineAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                            horizotalSectionPort1 = "EndCap";
                            horRoutePlaneOffset = -dShoeHeight;
                            horizontalRouteAxisOffset = -overHangRight - boundingBoxWidth;
                        }
                    }
                }
                //Add joint Between Horizontal Section and BoundingBox
                JointHelper.CreateRigidJoint(HORSECTION1, horizotalSectionPort1, "-1", routePort, routeHorPlane.A, routeHorPlane.B, routeHorPlane.C, routeHorPlane.D, horRoutePlaneOffset, horizontalRouteAxisOffset, horizontalRouteOriginOffset);

                //Add Joint Between the Horizontal and Vertical Beam
                JointHelper.CreateRigidJoint(HORSECTION1, horizonatlSectionPort, VERTSECTION1, verticalSectionPort1, horVertPlane.A, horVertPlane.B, horVertPlane.C, horVertPlane.D, horizontalVertPlaneOffset, horizontalVertAxisOffset, horizontalVertOriginOffset);    //-dHorSecLength / 2 + SectionDepth / 2)
                
                //Add Joint Between the Ports of Vertical Beam
                JointHelper.CreatePrismaticJoint(VERTSECTION1, "BeginCap", VERTSECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0.0, 0.0);//11757

                //Add Joint Between the Ports of Vertical Beam
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION1, "EndCap", "-1", "Structure", Plane.XY);//4

                string strRefPortName;
                for (int routeIndex1 = 1; routeIndex1 <= routeCount; routeIndex1++)
                {
                    if (isUboltExist[routeIndex1 - 1] == true)
                    {
                        if (routeIndex1 == 1)
                            strRefPortName = "Route";
                        else
                            strRefPortName = "Route_" + routeIndex1;

                        //Add Joint Between the UBolt and Route
                        JointHelper.CreateTranslationalJoint(ubolts[routeIndex1 - 1], "Route", HORSECTION1, "Neutral", routeUboltPlane[routeIndex1 - 1].A, routeUboltPlane[routeIndex1 - 1].B, routeUboltPlane[routeIndex1 - 1].C, routeUboltPlane[routeIndex1 - 1].D, boltRoutePlaneOffset);

                        //Add Joint Between the UBolt and Route
                        JointHelper.CreatePointOnAxisJoint(ubolts[routeIndex1 - 1], "Route", "-1", strRefPortName, Axis.X);
                    }
                }

                if (includeLeftPad == true)
                    JointHelper.CreateRigidJoint(LEFTPAD, "Port1", VERTSECTION1, "EndFace", padPlane.A, padPlane.B, padPlane.C, padPlane.D, 0.0, 0.0, 0.0);

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