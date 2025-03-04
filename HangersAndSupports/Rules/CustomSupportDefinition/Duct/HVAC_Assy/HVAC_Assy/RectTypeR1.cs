//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectTypeR1.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectTypeR1
//   Author       :Vijaya
//   Creation Date:12.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  12.Jun.2013     Vijaya   CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//  31.Oct.2014     PVK      TR-CP-260301	Resolve coverity issues found in August 2014 report
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
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
    public class RectTypeR1 : CustomSupportDefinition
    {
        //Constants
        private const string HORSECTION1 = "HORSECTION1";
        private const string HORSECTION2 = "HORSECTION2";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";

        string sectionSize;
        string  sizeOfSection, stStandard;              
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
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RectTypeR1.cs", 72);
                            return null;
                        }
                    }
                    string strSteelStd;
                    value = GenericHelper.GetDataByRule("HgrHVACSSteelStandardName", supportObject, out strSteelStd);

                    if (strSteelStd == "JIS-2005")
                        stStandard = "Japan-2005";
                    else if (strSteelStd == "Russian")
                        stStandard = "Russia";
                    else if (strSteelStd == "GB")
                        stStandard = "CHINA-2006";
                    else if (strSteelStd == "ICHA-2000")
                        stStandard = "Chile-2000";
                    else if (strSteelStd == "BS5950-1:2000")
                        stStandard = "BS";
                    else if (strSteelStd == "AUST-OneSteel-05")
                        stStandard = "AUST-05";
                    else if (strSteelStd == "AISC-METRIC")
                        stStandard = "AISC-Metric";
                    else
                        stStandard = strSteelStd;

                    //Create the list of parts
                    parts.Add(new PartInfo(HORSECTION1, sizeOfSection + " " + stStandard));
                    parts.Add(new PartInfo(HORSECTION2, sizeOfSection + " " + stStandard));
                    parts.Add(new PartInfo(VERTSECTION1, sizeOfSection + " " + stStandard));
                    parts.Add(new PartInfo(VERTSECTION2, sizeOfSection + " " + stStandard));

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
                
                double boundingBoxWidth = boundingBox.Width,boundingBoxHeight = boundingBox.Height;
                
                string supportingType;
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    supportingType = "Slab";
                else
                {
                    if (SupportHelper.SupportingObjects.Count > 1)
                    {
                        supportingType = "steel";
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                            supportingType = "Slab";          //Two Slabs

                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Member))
                            supportingType = "Slab-Steel";    //Slab then Steel

                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                            supportingType = "Steel-Slab";    //Steel then Slab
                    }
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                        {
                            if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                                supportingType = "Steel";    //Steel                      
                            else
                                supportingType = "Slab";
                        }
                        else
                            supportingType = "Slab";
                    }
                }
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                //Get Section Structure dimensions
                BusinessObject horizontalSectionPart = componentDictionary[HORSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double sectionWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double sectionDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;

                //Get the Current Location in the Route Connection Cycle
                bool[] isOffsetApplied = new bool[2];
                isOffsetApplied = HVACAssemblyServices.GetIsLugEndOffsetApplied(this);

                string[] structPort = new string[2];
                structPort = HVACAssemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0];
                string rightStructPort = structPort[1];

                // Set Values of Part Occurance Attributes
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                //Set SectionSize attribute value on the support
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (sectionFromRule)
                {
                    CodelistItem codeList = metadataManager.GetCodelistInfo("HVACAssyLSecSize", "UDP").GetCodelistItem(sizeOfSection);
                    support.SetPropertyValue(codeList.Value, "IJOAHgrHVACAssyLSize", "LSectionSize");
                }
               
                // Miscallaneous
                //Get the route structure distance
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                double routeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                string structPort1, structPort2;

                structPort1 = "Structure";
                structPort2 = structPort1;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    componentDictionary[HORSECTION1].SetPropertyValue(boundingBoxWidth + 2 * sectionWidth, "IJUAHgrOccLength", "Length");
                    componentDictionary[HORSECTION2].SetPropertyValue(boundingBoxWidth + 2 * sectionWidth, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    Double horSec1Length = 0, horSec2Length = 0, distBetStruct = 0; 

                    if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && Math.Abs(routeStructAngle) > (Math.PI - 0.001)))
                       distBetStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Vertical);
                    else
                       distBetStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
                   
                    if (supportingType.ToUpper().Equals("SLAB"))
                         horSec1Length = boundingBoxWidth + 2 * sectionDepth;
                    else if (supportingType.ToUpper().Equals("STEEL"))
                    {
                        horSec1Length = boundingBoxWidth + 2 * sectionDepth;
                        if (SupportHelper.SupportingObjects.Count > 1)
                            structPort2 = "Struct_2";
                    }
                    else if (supportingType.Equals("STEEL-SLAB") || supportingType.Equals("SLAB-STEEL"))
                    {
                        horSec1Length = boundingBoxWidth + 2 * sectionDepth;
                        structPort2 = "Struct_2";
                    }
                    horSec2Length = horSec1Length;
                    componentDictionary[HORSECTION1].SetPropertyValue(horSec1Length, "IJUAHgrOccLength", "Length");
                    componentDictionary[HORSECTION2].SetPropertyValue(horSec2Length, "IJUAHgrOccLength", "Length");
                }

                Plane horSec1RoutePlaneA = new Plane(), horSec1RoutePlaneB = new Plane();
                Axis horSec1RouteAxisA = new Axis(), horSec1RouteAxisB = new Axis();               
                
                String horSecPort = string.Empty, routePort = string.Empty;
                Double horSec1RouteAxOffset = 0, horSec1RouteOrgOffset = 0, angle;
                
                angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                Boolean isRouteVertical;
                //when pipe is vertical
                if (routeAngle < Math.PI / 4)
                    isRouteVertical = true;
                else if (routeAngle > 3 * Math.PI / 4)
                    isRouteVertical = true;
                else
                    isRouteVertical = false;

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    routePort = "BBSV_Low";
                else
                    if (isRouteVertical)
                        routePort = "BBR_Low";
                    else
                        routePort = "BBRV_Low";
                
                if (Configuration == 1)
                {
                    horSec1RoutePlaneA = Plane.ZX;
                    horSec1RoutePlaneB = Plane.NegativeXY;
                    horSec1RouteAxisA = Axis.X;
                    horSec1RouteAxisB = Axis.X;
                    horSecPort = "BeginCap";
                    horSec1RouteAxOffset = sectionWidth;
                    horSec1RouteOrgOffset = -sectionDepth / 2;
                }
                else if (Configuration == 2)
                {
                    horSec1RoutePlaneA = Plane.ZX;
                    horSec1RoutePlaneB = Plane.NegativeXY;
                    horSec1RouteAxisA = Axis.X;
                    horSec1RouteAxisB = Axis.NegativeX;
                    horSecPort = "EndCap";
                    horSec1RouteAxOffset = -sectionWidth;
                    horSec1RouteOrgOffset = -sectionDepth / 2;
                }

                //Create Joints
                //Add joint between Horizontal Section 1 and Route
                JointHelper.CreateRigidJoint(HORSECTION1, horSecPort, "-1", routePort, horSec1RoutePlaneA, horSec1RoutePlaneB, horSec1RouteAxisA, horSec1RouteAxisB, 0, horSec1RouteAxOffset, horSec1RouteOrgOffset);

                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(HORSECTION1, "Neutral", HORSECTION2, "Neutral", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -boundingBoxHeight - sectionWidth, 0);

                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(VERTSECTION1, "BeginCap", HORSECTION1, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX, sectionDepth, sectionDepth, 0);

                //Add Joint between BeginCap and EndCap of Vertical Section 1
                JointHelper.CreatePrismaticJoint(VERTSECTION1, "BeginCap", VERTSECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add Joint between EndCap of Vertical Section1 and Structure
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION1, "EndCap", "-1", structPort2, Plane.XY);

                //Add joint between Horizontal Section 1 and Vertical Section 2
                JointHelper.CreateRigidJoint(VERTSECTION2, "EndCap", HORSECTION1, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, sectionDepth, -sectionDepth, 0);

                //Add Joint between BeginCap and EndCap of Vertical Section 2
                JointHelper.CreatePrismaticJoint(VERTSECTION2, "BeginCap", VERTSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add Joint between BeginCap of Vertical Section and Structure
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION2, "BeginCap", "-1", structPort1, Plane.XY);
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

