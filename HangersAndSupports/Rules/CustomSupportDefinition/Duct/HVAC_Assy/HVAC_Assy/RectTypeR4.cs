﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectTypeR4.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectTypeR4
//   Author       :Vijaya
//   Creation Date:13.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  13.Jun.2013     Vijaya   CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//  16.Oct.2014     Chethan  TR-CP-237154  Some .Net HS_HVAC_Assy(FrameType R10, R3, R4, R7, R8) are behaving incorretly.  
//  31.Oct.2014     PVK      TR-CP-260301	Resolve coverity issues found in August 2014 report
//  22-Jan-2015     PVK      TR-CP-264951  Resolve coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle.Services;
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
    public class RectTypeR4 : CustomSupportDefinition
    {
        //Constants
        private const string HORSECTION1 = "HORSECTION1";
        private const string HORSECTION2 = "HORSECTION2";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";

        double offset;
        string sectionSize, sizeOfSection, steelStandard, supType, horSection1Size, sizeOfFlatBar;
        bool sectionFromRule, value;

        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    //Get the attributes from assembly                  
                    sectionFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssySecRule", "SectionFromRule")).PropValue;
                    PropertyValueCodelist sectionCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyLSize", "LSectionSize");
                    sectionSize = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionCodeList.PropValue).DisplayName;
                    offset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyOffset", "Offset")).PropValue;
                    supType = (string)((PropertyValueString)support.GetPropertyValue("IJOAHgrHVACAssySupType", "SupType")).PropValue;

                    if (sectionCodeList.PropValue == -1)
                        sectionCodeList.PropValue = 1;

                    BusinessObject supportObject = support;
                    Collection<object> sectionValue;
                    //Get the Section Size from Rule
                    if (sectionFromRule)
                    {
                        value = GenericHelper.GetDataByRule("HVACSectionSize", supportObject, out sectionValue);
                        if (sectionValue!=null)
                            sizeOfSection = (string)sectionValue[0];

                    }
                    else if (!sectionFromRule)
                    {
                        sizeOfSection = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSize).DisplayName;

                        if (sizeOfSection.ToUpper().Equals("NONE") || sizeOfSection.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RectTypeR4.cs", 75);
                            return null;
                        }
                    }

                    if (supType.Equals("R4"))
                        horSection1Size = sizeOfSection;
                    else if (supType.Equals("R4P"))
                    {
                        if (sectionFromRule)
                        {

                            value = GenericHelper.GetDataByRule("HVACSectionSize", supportObject, out sectionValue);
                            if (sectionValue != null)
                            sizeOfFlatBar = (string)sectionValue[1];

                        }  
                        else if (!sectionFromRule)
                        {
                            
                            PropertyValueCodelist flatBarSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyFBSize", "FBSectionSize");
                            sizeOfFlatBar = flatBarSizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(flatBarSizeCodeList.PropValue).DisplayName;

                            if (sizeOfFlatBar.ToUpper().Equals("NONE") || sizeOfFlatBar.Equals(string.Empty))
                            {
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RectTypeR4.cs", 94);
                                return null;
                            }
                        }
                        horSection1Size = sizeOfFlatBar;
                    }
                    //GetSteel Standard                   
                    string steelStd;
                    value = GenericHelper.GetDataByRule("HgrHVACSSteelStandardName", supportObject, out steelStd);

                    if (steelStd == "JIS-2005")
                        steelStandard = "Japan-2005";
                    else if (steelStd == "Russian")
                        steelStandard = "Russia";
                    else if (steelStd == "GB")
                        steelStandard = "CHINA-2006";
                    else if (steelStd == "ICHA-2000")
                        steelStandard = "Chile-2000";
                    else if (steelStd == "BS5950-1:2000")
                        steelStandard = "BS";
                    else if (steelStd == "AUST-OneSteel-05")
                        steelStandard = "AUST-05";
                    else if (steelStd == "AISC-METRIC")
                        steelStandard = "AISC-Metric";
                    else
                        steelStandard = steelStd;

                    //Create the list of parts
                    parts.Add(new PartInfo(HORSECTION1, horSection1Size + " " + steelStandard));
                    parts.Add(new PartInfo(HORSECTION2, sizeOfSection + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION1, sizeOfSection + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION2, sizeOfSection + " " + steelStandard));

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
                BoundingBoxHelper.CreateStandardBoundingBoxes(true);
                BoundingBox boundingBox;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                //Get Section Structure dimensions
                BusinessObject horizonta2SectionPart = componentDictionary[HORSECTION2].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection2 = (CrossSection)horizonta2SectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double sectionWidth = (double)((PropertyValueDouble)crosssection2.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double sectionDepth = (double)((PropertyValueDouble)crosssection2.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;

                double fbSectionWidth = 0, fbSectionDepth = 0;
                BusinessObject horizontalSectionPart = componentDictionary[HORSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection1 = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                if (supType.Equals("R4P"))
                {
                    fbSectionWidth = (double)((PropertyValueDouble)crosssection1.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                    fbSectionDepth = (double)((PropertyValueDouble)crosssection1.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                }
                //Get the Current Location in the Route Connection Cycle           
                // Set Values of Part Occurance Attributes
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");

                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");

                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");

                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");

                //Set SectionSize attribute value on the support
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (sectionFromRule)
                {
                    CodelistItem codeList = metadataManager.GetCodelistInfo("HVACAssyLSecSize", "UDP").GetCodelistItem(sizeOfSection);
                    support.SetPropertyValue(codeList.Value, "IJOAHgrHVACAssyLSize", "LSectionSize");
                }
                if (supType.Equals("R4P"))
                {
                    if (sectionFromRule)
                    {
                        CodelistItem codeList = metadataManager.GetCodelistInfo("HVACAssyFBSecSize", "UDP").GetCodelistItem(sizeOfFlatBar);
                        support.SetPropertyValue(codeList.Value, "IJOAHgrHVACAssyFBSize", "FBSectionSize");
                    }
                }
                // Miscallaneous
                String leftStructPort = string.Empty, rightStructPort = string.Empty, midStructPort = string.Empty, horSecPort = string.Empty, routePort, beginPort = string.Empty, endPort = string.Empty, structPort1, structPort2;
                Double distBetStructPort = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_4", PortDistanceType.Vertical);

                if (supType.Equals("R4"))
                    componentDictionary[HORSECTION1].SetPropertyValue(boundingBoxWidth + 2 * sectionWidth, "IJUAHgrOccLength", "Length");
                else if (supType.Equals("R4P"))
                    componentDictionary[HORSECTION1].SetPropertyValue(boundingBoxWidth, "IJUAHgrOccLength", "Length");

                double horDistBetStructPort1 = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
                double horDistBetStructPort2 = RefPortHelper.DistanceBetweenPorts("Struct_2", "Struct_3", PortDistanceType.Horizontal);
                double horDistBetStructPort3 = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_3", PortDistanceType.Horizontal);

                if ((horDistBetStructPort1 + horDistBetStructPort2) > (horDistBetStructPort3 - 0.001) && (horDistBetStructPort1 + horDistBetStructPort2) < (horDistBetStructPort3 + 0.001))
                {
                    leftStructPort = "Structure";
                    midStructPort = "Struct_2";
                    rightStructPort = "Struct_3";
                }
                else if ((horDistBetStructPort1 + horDistBetStructPort3) > (horDistBetStructPort2 - 0.001) && (horDistBetStructPort1 + horDistBetStructPort3) < (horDistBetStructPort2 + 0.001))
                {
                    leftStructPort = "Struct_2";
                    midStructPort = "Structure";
                    rightStructPort = "Struct_3";
                }
                else if ((horDistBetStructPort2 + horDistBetStructPort3) > (horDistBetStructPort1 - 0.001) && (horDistBetStructPort2 + horDistBetStructPort3) < (horDistBetStructPort1 + 0.001))
                {
                    leftStructPort = "Structure";
                    midStructPort = "Struct_3";
                    rightStructPort = "Struct_2";
                }
                //Set Values of EndOverLength for Vertical Sections
                componentDictionary[VERTSECTION1].SetPropertyValue(-offset, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(-offset, "IJUAHgrOccOverLength", "BeginOverLength");

                string[] structPortName = new string[2];
                structPortName[0] = leftStructPort;
                structPortName[1] = midStructPort;

                Double angle1;
                angle1 = HVACAssemblyServices.GetVectorProjectioBetweenPorts(this, "Route", "Structure", PortAxisType.Y);

                if (SupportHelper.SupportingObjects.Count > 1)
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                    {
                        if (angle1 < 0)
                        {
                            structPortName[0] = midStructPort;
                            structPortName[1] = leftStructPort;
                        }
                        else
                        {
                            structPortName[0] = leftStructPort;
                            structPortName[1] = midStructPort;
                        }
                    }
                 }


                if (supType.Equals("R4"))
                {
                    componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                }
                else if (supType.Equals("R4P"))
                {
                    componentDictionary[VERTSECTION2].SetPropertyValue(-sectionDepth, "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[VERTSECTION1].SetPropertyValue(-sectionDepth, "IJUAHgrOccOverLength", "BeginOverLength");
                }
                if (Configuration == 1)
                {
                    beginPort = "BeginCap";
                    endPort = "EndCap";
                }
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is placed only by Point.", "", "RectTypeR4.cs", 301);
                    return;
                }


                structPort1 = structPortName[0];
                structPort2 = structPortName[1];

                Double angle2;
                angle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, structPortName[0], PortAxisType.X, OrientationAlong.Direct);



                Plane horSec1RoutePlaneA = new Plane(), horSec1RoutePlaneB = new Plane(), horSec2RoutePlaneA = new Plane(), horSec2RoutePlaneB = new Plane(), horSec3RoutePlaneA = new Plane(), horSec3RoutePlaneB = new Plane();
                Axis horSec1RouteAxisA = new Axis(), horSec1RouteAxisB = new Axis(), horSec2RouteAxisA = new Axis(), horSec2RouteAxisB = new Axis(), horSec3RouteAxisA = new Axis(), horSec3RouteAxisB = new Axis();

                Double horSec1RouteAxOffset = 0, horSec2RoutePlOffset = 0, horSec2RouteAxOffset = 0, horSec2RouteOrgOffset = 0;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    routePort = "BBSR_Low";
                else
                    routePort = "BBR_Low";


                if (HgrCompareDoubleService.cmpdbl(Math.Round(angle2,4) ,  Math.Round(Math.PI,4))==false)
                {

                    horSec3RoutePlaneA = Plane.YZ;
                    horSec3RoutePlaneB = Plane.XY;
                    horSec3RouteAxisA = Axis.Y;
                    horSec3RouteAxisB = Axis.X;

                }
                else
                {

                    horSec3RoutePlaneA = Plane.YZ;
                    horSec3RoutePlaneB = Plane.XY;
                    horSec3RouteAxisA = Axis.Y;
                    horSec3RouteAxisB = Axis.NegativeX;
                }

                if (Configuration == 1)
                {
                    horSec1RoutePlaneA = Plane.ZX;
                    horSec1RoutePlaneB = Plane.NegativeXY;
                    horSec1RouteAxisA = Axis.X;
                    horSec1RouteAxisB = Axis.X;
                    horSecPort = "BeginCap";
                    horSec1RouteAxOffset = sectionWidth;

                    horSec2RoutePlaneA = Plane.YZ;
                    horSec2RoutePlaneB = Plane.XY;
                    horSec2RouteAxisA = Axis.Y;
                    horSec2RouteAxisB = Axis.NegativeX;
                    horSec2RoutePlOffset = fbSectionWidth;
                    horSec2RouteAxOffset = boundingBoxWidth;
                    horSec2RouteOrgOffset = fbSectionDepth;
                    if (supType.Equals("R4P"))
                    {
                        componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                        componentDictionary[HORSECTION2].SetPropertyValue(Math.PI/2, "IJOAhsBeginCap", "BeginCapRotZ");
                        componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                        componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");


                    }
                }
                else if (Configuration == 2)
                {
                    if (supType.Equals("R4"))
                    {
                        horSec1RoutePlaneA = Plane.ZX;
                        horSec1RoutePlaneB = Plane.NegativeXY;
                        horSec1RouteAxisA = Axis.X;  
                        horSec1RouteAxisB = Axis.NegativeX;
                        horSecPort = "EndCap";
                        horSec1RouteAxOffset = -sectionWidth;
                    }
                    if (supType.Equals("R4P"))
                    {
                        horSec2RoutePlaneA = Plane.YZ;
                        horSec2RoutePlaneB = Plane.XY;
                        horSec2RouteAxisA = Axis.Y;
                        horSec2RouteAxisB = Axis.NegativeX;
                        horSecPort = "BeginCap";
                        horSec2RoutePlOffset = fbSectionWidth;
                        horSec2RouteAxOffset = boundingBoxWidth;
                        horSec2RouteOrgOffset = fbSectionDepth;
                        componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                        componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                        componentDictionary[VERTSECTION1].SetPropertyValue(-Math.PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                        componentDictionary[VERTSECTION2].SetPropertyValue(-Math.PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                    }
                }
                //Create Joints

                if (supType.Equals("R4"))
                {
                    //Add joint between Horizontal Section 1 and Route
                    JointHelper.CreateRigidJoint(HORSECTION1, horSecPort, "-1", routePort, horSec1RoutePlaneA, horSec1RoutePlaneB, horSec1RouteAxisA, horSec1RouteAxisB, 0, horSec1RouteAxOffset, 0);

                    //Joint between EndCap of top horizontal and Middle structure
                    //JointHelper.CreateRigidJoint(HORSECTION2, "BeginCap", "-1", midStructPort, Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX, distBetStructPort - offset, 0, 0);
                    JointHelper.CreateRigidJoint(HORSECTION2, "BeginCap", "-1", structPort1, horSec3RoutePlaneA, horSec3RoutePlaneB, horSec3RouteAxisA, horSec3RouteAxisB, distBetStructPort - offset, 0, 0);

                    //Joint between BeginCap and EndCap of top horizontal
                    JointHelper.CreatePrismaticJoint(HORSECTION2, "BeginCap", HORSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    //Joint between BeginCap and rightmost structure
                    JointHelper.CreatePointOnPlaneJoint(HORSECTION2, "EndCap", "-1", structPort2, Plane.ZX);

                    //Add joint between Horizontal Section 1 and Vertical Section 1
                    JointHelper.CreateRigidJoint(VERTSECTION1, "BeginCap", HORSECTION1, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX, sectionDepth, sectionDepth, 0);

                    //Add Joint between BeginCap and EndCap of Vertical Section 1
                    JointHelper.CreatePrismaticJoint(VERTSECTION1, "BeginCap", VERTSECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    //Add Joint between EndCap of Vertical Section 1 and Deck
                    JointHelper.CreatePointOnPlaneJoint(VERTSECTION1, "EndCap", "-1", "Struct_4", Plane.XY);

                    //Add joint between Horizontal Section 1 and Vertical Section 2
                    JointHelper.CreateRigidJoint(VERTSECTION2, "EndCap", HORSECTION1, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, sectionDepth, -sectionDepth, 0);

                    //Add Joint between BeginCap and EndCap of Vertical Section 2
                    JointHelper.CreatePrismaticJoint(VERTSECTION2, "BeginCap", VERTSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    //Add Joint between BeginCaps of Verical Section 2 and top horizontal
                    JointHelper.CreatePointOnPlaneJoint(VERTSECTION2, "BeginCap", "-1", "Struct_4", Plane.XY);
                }
                else if (supType.Equals("R4P"))
                {
                    //Add joint between Horizontal Section 1 and Route
                    JointHelper.CreateRigidJoint(HORSECTION1, horSecPort, "-1", routePort, horSec2RoutePlaneA, horSec2RoutePlaneB, horSec2RouteAxisA, horSec2RouteAxisB, horSec2RoutePlOffset, horSec2RouteAxOffset, horSec2RouteOrgOffset);

                    //Joint between EndCap of top horizontal and Middle structure      
                    //JointHelper.CreateRigidJoint(HORSECTION2, "BeginCap", "-1", midStructPort, Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX, distBetStructPort - offset, 0, 0);
                    JointHelper.CreateRigidJoint(HORSECTION2, "BeginCap", "-1", structPort1, horSec3RoutePlaneA, horSec3RoutePlaneB, horSec3RouteAxisA, horSec3RouteAxisB, distBetStructPort - offset, 0, 0);


                    //Joint between BeginCap and EndCap of top horizontal
                    JointHelper.CreatePrismaticJoint(HORSECTION2, "BeginCap", HORSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    //Joint between BeginCap and rightmost structure
                    JointHelper.CreatePointOnPlaneJoint(HORSECTION2, "EndCap", "-1", structPort2, Plane.ZX);

                    //Add joint between Horizontal Section 1 and Vertical Section 1
                    JointHelper.CreateRigidJoint(VERTSECTION1, "BeginCap", HORSECTION1, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0, sectionDepth, fbSectionDepth);

                    //Add Joint between BeginCap and EndCap of Vertical Section 1
                    JointHelper.CreatePrismaticJoint(VERTSECTION1, "BeginCap", VERTSECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    //Add Joint between EndCap of Vertical Section 1 and Deck
                    JointHelper.CreatePointOnPlaneJoint(VERTSECTION1, "EndCap", "-1", "Struct_4", Plane.XY);

                    //Add joint between Horizontal Section 1 and Vertical Section 2
                    JointHelper.CreateRigidJoint(VERTSECTION2, "EndCap", HORSECTION1, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionDepth, fbSectionDepth);

                    //Add Joint between BeginCap and EndCap of Vertical Section 2
                    JointHelper.CreatePrismaticJoint(VERTSECTION2, "BeginCap", VERTSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                    //Add Joint between BeginCaps of Verical Section 2 and top horizontal
                    JointHelper.CreatePointOnPlaneJoint(VERTSECTION2, "BeginCap", "-1", "Struct_4", Plane.XY);
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
    }
}

