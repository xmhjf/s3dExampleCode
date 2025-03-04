//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectTypeR7.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectTypeR7
//   Author       :Vijaya
//   Creation Date:12.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  12.Jun.2013     Vijaya   CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//  31.Oct.2014     PVK      TR-CP-260301	Resolve coverity issues found in August 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle.Services;

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
    public class RectTypeR7 : CustomSupportDefinition
    {
        //Constants
        private const string HORSECTION1 = "HORSECTION1";
        private const string HORSECTION2 = "HORSECTION2";
        private const string HORSECTION3 = "HORSECTION3";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";

        double offset;
        string sectionSize, sizeOfSection, steelStandard;
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
                    BusinessObject supportObject = support;
                    Collection<object> sectionValue;
                    //Get the Section Size from Rule
                    if (sectionFromRule)
                    {
                        value = GenericHelper.GetDataByRule("HVACSectionSize", supportObject, out sectionValue);
                        if (sectionValue != null)
                        {
                            if (sectionValue[0] == null)
                                sizeOfSection = (string)sectionValue[1];
                            else
                                sizeOfSection = (string)sectionValue[0];
                        }
                    }
                    else if (!sectionFromRule)
                    {
                        sizeOfSection = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSize).DisplayName;

                        if (sizeOfSection.ToUpper().Equals("NONE") || sizeOfSection.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RectTypeR7.cs", 72);
                            return null;
                        }
                    }
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
                    parts.Add(new PartInfo(HORSECTION1, sizeOfSection + " " + steelStandard));
                    parts.Add(new PartInfo(HORSECTION2, sizeOfSection + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION1, sizeOfSection + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION2, sizeOfSection + " " + steelStandard));
                    parts.Add(new PartInfo(HORSECTION3, sizeOfSection + " " + steelStandard));

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
                BusinessObject horizontalSectionPart = componentDictionary[HORSECTION2].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double sectionWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double sectionDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;

                //Get the Current Location in the Route Connection Cycle
                // Set Values of Part Occurance Attributes
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapYOffset");

                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapYOffset");

                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapYOffset");

                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapYOffset");

                componentDictionary[HORSECTION3].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION3].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HORSECTION3].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                componentDictionary[HORSECTION3].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapYOffset");

                //Set SectionSize attribute value on the support
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (sectionFromRule)
                {
                    CodelistItem codeList = metadataManager.GetCodelistInfo("HVACAssyLSecSize", "UDP").GetCodelistItem(sizeOfSection);
                    support.SetPropertyValue(codeList.Value, "IJOAHgrHVACAssyLSize", "LSectionSize");
                }
                // Miscallaneous
                String leftStructPort = string.Empty, rightStructPort = string.Empty, midStructPort = string.Empty, routePort;
                componentDictionary[HORSECTION1].SetPropertyValue(boundingBoxWidth + 2 * sectionWidth, "IJUAHgrOccLength", "Length");

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
                string[] portArray = new string[3];
                portArray[0] = leftStructPort;
                portArray[1] = midStructPort;
                portArray[2] = rightStructPort;
                int[] indexes = new int[3];

                for (int index = 0; index < 3; index++)
                {
                    if (portArray[index] == "Structure")
                        indexes[index] = 1;
                    else
                    {
                        string[] port = portArray[index].Split('_');
                        indexes[index] = int.Parse(port[1]);
                    }
                }
                double SteelWebThickLeft = SupportingHelper.SupportingObjectInfo(indexes[0]).WebThickness;
                double SteelWebThickMid = SupportingHelper.SupportingObjectInfo(indexes[1]).WebThickness;
                double SteelWebThickRight = SupportingHelper.SupportingObjectInfo(indexes[2]).WebThickness;

                double dDistBetStructPortMid = RefPortHelper.DistanceBetweenPorts(midStructPort, "Struct_4", PortDistanceType.Vertical);

                string strHorSec3Port1 = string.Empty, strHorSec3Port2 = string.Empty, strHorSec2Port1 = string.Empty, strHorSec2Port2 = string.Empty;
                if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                {
                    Double angle = HVACAssemblyServices.GetRouteStructConfigAngle(this, midStructPort, leftStructPort, PortAxisType.Y);

                    if (Math.Abs(angle) < Math.PI / 2)
                    {
                        strHorSec2Port1 = "BeginCap";
                        strHorSec2Port2 = "EndCap";
                        componentDictionary[HORSECTION2].SetPropertyValue(-SteelWebThickMid / 2, "IJUAHgrOccOverLength", "BeginOverLength");
                        componentDictionary[HORSECTION2].SetPropertyValue(-SteelWebThickLeft / 2, "IJUAHgrOccOverLength", "EndOverLength");
                    }
                    else
                    {
                        strHorSec2Port1 = "EndCap";
                        strHorSec2Port2 = "BeginCap";
                        componentDictionary[HORSECTION2].SetPropertyValue(-SteelWebThickLeft / 2, "IJUAHgrOccOverLength", "BeginOverLength");
                        componentDictionary[HORSECTION2].SetPropertyValue(-SteelWebThickMid / 2, "IJUAHgrOccOverLength", "EndOverLength");

                    }
                }

                if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                {
                    Double angle = HVACAssemblyServices.GetRouteStructConfigAngle(this, midStructPort, rightStructPort, PortAxisType.Y);
                    if (Math.Abs(angle) < Math.PI / 2)
                    {
                        strHorSec3Port1 = "BeginCap";
                        strHorSec3Port2 = "EndCap";
                        componentDictionary[HORSECTION3].SetPropertyValue(-SteelWebThickMid / 2, "IJUAHgrOccOverLength", "BeginOverLength");
                        componentDictionary[HORSECTION3].SetPropertyValue(-SteelWebThickRight / 2, "IJUAHgrOccOverLength", "EndOverLength");
                    }
                    else
                    {
                        strHorSec3Port1 = "EndCap";
                        strHorSec3Port2 = "BeginCap";
                        componentDictionary[HORSECTION3].SetPropertyValue(-SteelWebThickRight / 2, "IJUAHgrOccOverLength", "BeginOverLength");
                        componentDictionary[HORSECTION3].SetPropertyValue(-SteelWebThickMid / 2, "IJUAHgrOccOverLength", "EndOverLength");
                    }
                }
                //Set Values of EndOverLength for Vertical Sections
                componentDictionary[VERTSECTION1].SetPropertyValue(-offset, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(-offset, "IJUAHgrOccOverLength", "BeginOverLength");

                double structAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, midStructPort, PortAxisType.X, OrientationAlong.Direct);

                Plane horSec1RoutePlaneA = new Plane(), horSec1RoutePlaneB = new Plane();
                Axis horSec1RouteAxisA = new Axis(), horSec1RouteAxisB = new Axis();

                Double horSec1RouteAxOffset = 0, horSec1RouteOrgOffset = 0;

                if (Configuration == 1)
                {
                    horSec1RoutePlaneA = Plane.YZ;
                    horSec1RoutePlaneB = Plane.NegativeXY;
                    horSec1RouteAxisA = Axis.Y;
                    horSec1RouteAxisB = Axis.NegativeX;
                    horSec1RouteAxOffset = sectionWidth;
                    horSec1RouteOrgOffset = 0;
                    if (structAngle >= Math.PI / 2)
                    {
                        componentDictionary[HORSECTION2].SetPropertyValue(Math.PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                        componentDictionary[HORSECTION3].SetPropertyValue(Math.PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                    }
                    else
                    {
                        componentDictionary[HORSECTION2].SetPropertyValue(Math.PI, "IJOAhsBeginCap", "BeginCapRotZ");
                        componentDictionary[HORSECTION3].SetPropertyValue(Math.PI, "IJOAhsBeginCap", "BeginCapRotZ");
                    }
                }
                else if (Configuration == 2)
                {
                    horSec1RoutePlaneA = Plane.YZ;
                    horSec1RoutePlaneB = Plane.NegativeXY;
                    horSec1RouteAxisA = Axis.Y;
                    horSec1RouteAxisB = Axis.X;
                    horSec1RouteAxOffset = boundingBoxWidth + sectionWidth;
                    horSec1RouteOrgOffset = 0;

                    if (structAngle >= Math.PI / 2)
                    {
                        componentDictionary[HORSECTION2].SetPropertyValue(Math.PI, "IJOAhsBeginCap", "BeginCapRotZ");
                        componentDictionary[HORSECTION3].SetPropertyValue(Math.PI, "IJOAhsBeginCap", "BeginCapRotZ");
                    }
                    else
                    {
                        componentDictionary[HORSECTION2].SetPropertyValue(Math.PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                        componentDictionary[HORSECTION3].SetPropertyValue(Math.PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                    }
                }
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    routePort = "BBSR_Low";
                else
                    routePort = "BBR_Low";
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "This support is placed only by Point.", "", "RectTypeR7.cs", 328);
                    return;
                }
                //Create Joints
                //Add joint between Horizontal Section 1 and Route
                JointHelper.CreateRigidJoint(HORSECTION1, "BeginCap", "-1", routePort, horSec1RoutePlaneA, horSec1RoutePlaneB, horSec1RouteAxisA, horSec1RouteAxisB, 0, horSec1RouteAxOffset, horSec1RouteOrgOffset);

                //Joint between Horizontal Section 2 and Left Structure port
                JointHelper.CreateRigidJoint(HORSECTION2, strHorSec2Port1, "-1", midStructPort, Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.Y, -(dDistBetStructPortMid - offset), 0, 0);

                //Joint between BeginCap and EndCap of Horizontal Section 2
                JointHelper.CreatePrismaticJoint(HORSECTION2, "BeginCap", HORSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Joint between Horizontal Section 2 and Middle Structure Port
                JointHelper.CreatePointOnPlaneJoint(HORSECTION2, strHorSec2Port2, "-1", leftStructPort, Plane.ZX);

                //Joint between Horizontal Section 3 and Middle structure port
                JointHelper.CreateRigidJoint(HORSECTION3, strHorSec3Port1, "-1", midStructPort, Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.Y, -(dDistBetStructPortMid - offset), 0, 0);

                //Joint between BeginCap and EndCap of Horizontal Section 3
                JointHelper.CreatePrismaticJoint(HORSECTION3, "BeginCap", HORSECTION3, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                // Joint between BeginCap and Struct_2
                JointHelper.CreatePointOnPlaneJoint(HORSECTION3, strHorSec3Port2, "-1", rightStructPort, Plane.ZX);

                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(VERTSECTION1, "BeginCap", HORSECTION1, "BeginCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeY, 0, sectionDepth, 0);

                //Add Joint between BeginCap and EndCap of Vertical Section 1
                JointHelper.CreatePrismaticJoint(VERTSECTION1, "BeginCap", VERTSECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add Joint between EndCap of Vertical Section 1 and End Cap of top horizontal
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION1, "EndCap", "-1", "Struct_4", Plane.XY);

                //Add joint between Horizontal Section 1 and Vertical Section 2
                JointHelper.CreateRigidJoint(VERTSECTION2, "EndCap", HORSECTION1, "EndCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0, -sectionDepth, 0);

                //Add Joint between BeginCap and EndCap of Vertical Section 2
                JointHelper.CreatePrismaticJoint(VERTSECTION2, "BeginCap", VERTSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add Joint between BeginCaps of Verical Section 1 and top horizontal
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION2, "BeginCap", "-1", "Struct_4", Plane.XY);
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

