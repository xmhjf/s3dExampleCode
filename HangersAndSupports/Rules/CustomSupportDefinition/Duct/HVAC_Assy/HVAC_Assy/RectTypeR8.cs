//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectTypeR8.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectTypeR8
//   Author       :Vijaya
//   Creation Date:13.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  13.Jun.2013     Vijaya   CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//  31.Oct.2014     PVK      TR-CP-260301	Resolve coverity issues found in August 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Linq;
using Ingr.SP3D.Route.Middle;
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

    public class RectTypeR8 : CustomSupportDefinition
    {
        //Constants 
        private const string HORSECTION1 = "HORSECTION1";
        private const string HORSECTION2 = "HORSECTION2";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";
        private const string HORSECTION3 = "HORSECTION3";
        private const string HORSECTION4 = "HORSECTION4";

        double offset, horOffset;
        string sectionSize, sizeOfSection = string.Empty, steelStandard;
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
                    horOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyHorOffset", "HorOffset")).PropValue;

                    if (sectionCodeList.PropValue == -1)
                        sectionCodeList.PropValue = 1;

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
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RectTypeR8.cs", 81);
                            return null;
                        }
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
                    parts.Add(new PartInfo(HORSECTION1, sizeOfSection + " " + steelStandard));
                    parts.Add(new PartInfo(HORSECTION2, sizeOfSection + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION1, sizeOfSection + " " + steelStandard));
                    parts.Add(new PartInfo(VERTSECTION2, sizeOfSection + " " + steelStandard));
                    parts.Add(new PartInfo(HORSECTION3, sizeOfSection + " " + steelStandard));
                    parts.Add(new PartInfo(HORSECTION4, sizeOfSection + " " + steelStandard));

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

                ReadOnlyCollection<BusinessObject> ductObjects = boundingBox.SupportedObjectsAtEdge(BoundingBoxEdge.Left);
                IRouteFeatureWithCrossSection ductInfo = (IRouteFeatureWithCrossSection)ductObjects.First();
                double ductWidth1 = ductInfo.Depth;
                ductObjects = boundingBox.SupportedObjectsAtEdge(BoundingBoxEdge.Right);
                ductInfo = (IRouteFeatureWithCrossSection)ductObjects.First();
                double ductWidth2 = ductInfo.Depth;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                //Get Section Structure dimensions
                BusinessObject horizonta1SectionPart = componentDictionary[HORSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection1 = (CrossSection)horizonta1SectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double sectionWidth = (double)((PropertyValueDouble)crosssection1.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double sectionDepth = (double)((PropertyValueDouble)crosssection1.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;

                // Set Values of Part Occurance Attributes
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                componentDictionary[HORSECTION3].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION3].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                componentDictionary[HORSECTION4].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION4].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                //Set SectionSize attribute value on the support
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (sectionFromRule)
                {
                    CodelistItem codeList = metadataManager.GetCodelistInfo("HVACAssyLSecSize", "UDP").GetCodelistItem(sizeOfSection);
                    support.SetPropertyValue(codeList.Value, "IJOAHgrHVACAssyLSize", "LSectionSize");
                }
                //have to do
                // Miscallaneous
                componentDictionary[HORSECTION1].SetPropertyValue(boundingBoxWidth + 2 * sectionWidth, "IJUAHgrOccLength", "Length");
                componentDictionary[HORSECTION2].SetPropertyValue(boundingBoxWidth + 2 * sectionWidth, "IJUAHgrOccLength", "Length");

                double distBetStructPort = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_3", PortDistanceType.Vertical);
                double horDistBetStructPort = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
                double horDist1BetStructRoute = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_Low", PortDistanceType.Horizontal);
                double horDist2BetStructRoute = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_High", PortDistanceType.Horizontal);

                if (horOffset > horDistBetStructPort)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "HorOffset value should be less than the horizontal distance between first two structures selected.", "", "RectTypeR8.cs", 208);
                    return;
                }

                componentDictionary[HORSECTION4].SetPropertyValue(horDistBetStructPort, "IJUAHgrOccLength", "Length");

                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                double routeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                string horSec3Port1 = string.Empty, horSec3Port2 = string.Empty;
                double structAngle = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, OrientationAlong.Global_X);
                if (SupportHelper.SupportingObjects.Count > 1)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        Double angle1 = HVACAssemblyServices.GetRouteStructConfigAngle(this, "Structure", "Struct_2", PortAxisType.Y);
                        //the port is the right structure port
                        if (Math.Abs(angle1) < Math.PI / 2)
                        {
                            horSec3Port1 = "BeginCap";
                            horSec3Port2 = "EndCap";
                        }
                        else
                        {
                            horSec3Port1 = "EndCap";
                            horSec3Port2 = "BeginCap";
                        }
                    }
                }
                if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is placed only by Structure.", "", "RectTypeR8.cs", 238);
                    return;
                }
                double angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                Boolean isRouteVertical;
                string routePort = string.Empty;
                //when pipe is vertical
                if (angle < Math.PI / 4)
                    isRouteVertical = true;
                else if (angle > 3 * Math.PI / 4)
                    isRouteVertical = true;
                else
                    isRouteVertical = false;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    routePort = "BBSV_Low";
                else
                    if (isRouteVertical)
                        routePort = "BBR_Low";
                    else
                        routePort = "BBRV_Low";

                Plane vertSec1RoutePlaneA = new Plane(), vertSec1RoutePlaneB = new Plane(), vertSec2RoutePlaneA = new Plane(), vertSec2RoutePlaneB = new Plane(), vertSec3RoutePlaneA = new Plane(), vertSec3RoutePlaneB = new Plane();
                Axis vertSec1RouteAxisA = new Axis(), vertSec1RouteAxisB = new Axis(), vertSec2RouteAxisA = new Axis(), vertSec2RouteAxisB = new Axis(), vertSec3RouteAxisA = new Axis(), vertSec3RouteAxisB = new Axis();
                Double vertSec1RoutePlOffset = 0, vertSecOrgOffset = 0, structSecOriginOffset = 0, horOffset1;
                string beginPort = string.Empty, endPort = string.Empty;

                if (horOffset < horDistBetStructPort)
                    horOffset1 = (horDistBetStructPort / 2) - horOffset;
                else
                    horOffset1 = horOffset - (horDistBetStructPort / 2);

                if (horDist1BetStructRoute < horDist2BetStructRoute)
                    structSecOriginOffset = boundingBoxWidth - ductWidth1 / 2;
                else
                    structSecOriginOffset = ductWidth2 / 2;
                if (Configuration == 1)
                {
                    vertSec1RoutePlaneA = Plane.ZX;
                    vertSec1RoutePlaneB = Plane.NegativeXY;
                    vertSec1RouteAxisA = Axis.X;
                    vertSec1RouteAxisB = Axis.NegativeX;

                    vertSec2RoutePlaneA = Plane.ZX;
                    vertSec2RoutePlaneB = Plane.XY;
                    vertSec2RouteAxisA = Axis.X;
                    vertSec2RouteAxisB = Axis.NegativeX;

                    vertSec3RoutePlaneA = Plane.XY;
                    vertSec3RoutePlaneB = Plane.NegativeZX;
                    vertSec3RouteAxisA = Axis.Y;
                    vertSec3RouteAxisB = Axis.NegativeX;
                    beginPort = "EndCap";
                    endPort = "BeginCap";
                    vertSec1RoutePlOffset = horOffset1;
                    vertSecOrgOffset = sectionWidth / 2;
                    if (SupportHelper.SupportingObjects.Count > 1)
                    {
                        componentDictionary[VERTSECTION1].SetPropertyValue(sectionWidth, "IJUAHgrOccOverLength", "BeginOverLength");
                        componentDictionary[VERTSECTION2].SetPropertyValue(sectionWidth, "IJUAHgrOccOverLength", "EndOverLength");
                    }
                }
                else if (Configuration == 2)
                {
                    vertSec1RoutePlaneA = Plane.ZX;
                    vertSec1RoutePlaneB = Plane.XY;
                    vertSec1RouteAxisA = Axis.X;
                    vertSec1RouteAxisB = Axis.NegativeX;

                    vertSec2RoutePlaneA = Plane.ZX;
                    vertSec2RoutePlaneB = Plane.NegativeXY;
                    vertSec2RouteAxisA = Axis.X;
                    vertSec2RouteAxisB = Axis.NegativeX;

                    vertSec3RoutePlaneA = Plane.XY;
                    vertSec3RoutePlaneB = Plane.ZX;
                    vertSec3RouteAxisA = Axis.Y;
                    vertSec3RouteAxisB = Axis.NegativeX;

                    beginPort = "BeginCap";
                    endPort = "EndCap";
                    vertSec1RoutePlOffset = -horOffset1;
                    vertSecOrgOffset = -sectionWidth / 2;
                    if (SupportHelper.SupportingObjects.Count > 1)
                    {
                        componentDictionary[VERTSECTION1].SetPropertyValue(sectionWidth, "IJUAHgrOccOverLength", "BeginOverLength");
                        componentDictionary[VERTSECTION2].SetPropertyValue(sectionWidth, "IJUAHgrOccOverLength", "EndOverLength");
                    }
                }
                //Create Joints
                //Add Joint between Horizontal Section 3 and Structure
                JointHelper.CreateRigidJoint(HORSECTION3, horSec3Port1, "-1", "Structure", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, distBetStructPort - offset, 0, structSecOriginOffset);
                //Add Joint between BeginCap and EndCap of Horizontal Section 3
                JointHelper.CreatePrismaticJoint(HORSECTION3, "BeginCap", HORSECTION3, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                //Add Joint between Horizontal Section 3 and Structure
                JointHelper.CreatePointOnPlaneJoint(HORSECTION3, horSec3Port2, "-1", "Struct_2", Plane.ZX);
                //Add Joint between Horizontal Section 3 and Horizontal Section 4
                JointHelper.CreateRigidJoint(HORSECTION3, "Neutral", HORSECTION4, "Neutral", Plane.ZX, Plane.ZX, Axis.X, Axis.NegativeX, 0, 0, boundingBoxWidth - sectionWidth);
                //Add joint between Vertical Section 1 and Horizontal Section 3
                JointHelper.CreateRigidJoint(VERTSECTION1, beginPort, HORSECTION3, "Neutral", vertSec1RoutePlaneA, vertSec1RoutePlaneB, vertSec1RouteAxisA, vertSec1RouteAxisB, vertSec1RoutePlOffset, -vertSecOrgOffset, -sectionWidth / 2);
                //Add joint between Vertical Section 2 and Horizontal Section 4
                JointHelper.CreateRigidJoint(VERTSECTION2, endPort, HORSECTION4, "Neutral", vertSec2RoutePlaneA, vertSec2RoutePlaneB, vertSec2RouteAxisA, vertSec2RouteAxisB, vertSec1RoutePlOffset, vertSecOrgOffset, -sectionWidth / 2);
                //Add Joint between BeginCap and EndCap of Vertical Section 1
                JointHelper.CreatePrismaticJoint(VERTSECTION1, "BeginCap", VERTSECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                //Add joint between Vertical Section 1 and Route
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION1, endPort, "-1", "BBR_Low", Plane.XY);
                //Add Joint between BeginCap and EndCap of Vertical Section 2
                JointHelper.CreatePrismaticJoint(VERTSECTION2, "BeginCap", VERTSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                //Add joint between Vertical Section 2 and Route
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION2, beginPort, "-1", "BBR_Low", Plane.XY);
                //Add joint between Vertical Section 1 and Horizontal Section 1
                JointHelper.CreateRigidJoint(VERTSECTION1, endPort, HORSECTION1, beginPort, vertSec3RoutePlaneA, vertSec3RoutePlaneB, vertSec3RouteAxisA, vertSec3RouteAxisB, 0, sectionDepth, 0);
                //Add joint between Horizontal Section 1 and Horizontal Section 2
                JointHelper.CreateRigidJoint(HORSECTION1, "Neutral", HORSECTION2, "Neutral", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -boundingBoxHeight - sectionDepth, 0);
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



