//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectTypeR2.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectTypeR2
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
    public class RectTypeR2 : CustomSupportDefinition
    {
        //Constants
        private const string HORSECTION1 = "HORSECTION1";
        private const string HORSECTION2 = "HORSECTION2";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";

        double horOverhang, verOverhang;
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
                    horOverhang = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyR2OH", "HorOverhang")).PropValue;
                    verOverhang = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyR2OH", "VerOverhang")).PropValue;

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
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RectTypeR2.cs", 76);
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

                Note noteDimenssion2 = CreateNote("Dim 2", componentDictionary[HORSECTION2], "BeginCap");
                noteDimenssion2.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem fabrication2 = noteDimenssion2.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                noteDimenssion2.SetPropertyValue(fabrication2, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                noteDimenssion2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                //Get Section Structure dimensions
                BusinessObject horizontalSectionPart = componentDictionary[HORSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double sectionWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double sectionDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;

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
                componentDictionary[HORSECTION1].SetPropertyValue(boundingBoxWidth + 2 * horOverhang, "IJUAHgrOccLength", "Length");
                componentDictionary[HORSECTION2].SetPropertyValue(boundingBoxWidth + 2 * horOverhang, "IJUAHgrOccLength", "Length");
                componentDictionary[VERTSECTION1].SetPropertyValue(boundingBoxHeight + 2 * verOverhang, "IJUAHgrOccLength", "Length");
                componentDictionary[VERTSECTION2].SetPropertyValue(boundingBoxHeight + 2 * verOverhang, "IJUAHgrOccLength", "Length");

                //set overhang value
                support.SetPropertyValue(horOverhang, "IJOAHgrHVACAssyR2OH", "HorOverhang");
                support.SetPropertyValue(verOverhang, "IJOAHgrHVACAssyR2OH", "VerOverhang");

                Plane horSec1RoutePlaneA = new Plane(), horSec1RoutePlaneB = new Plane(), vertSec1RoutePlaneA = new Plane(), vertSec1RoutePlaneB = new Plane();
                Axis horSec1RouteAxisA = new Axis(), horSec1RouteAxisB = new Axis(), vertSec1RouteAxisA = new Axis(), vertSec1RouteAxisB = new Axis();

                String routePort = string.Empty;
                Double horSec1RouteAxOffset = 0, horSec1RouteOrgOffset = 0, horSec1RoutePlOffset = 0, verSec1RouteAxOffset = 0, verSec1RouteOrgOffset = 0, verSec1RoutePlOffset = 0;

                if (Configuration == 1)
                {
                    horSec1RoutePlaneA = Plane.ZX;
                    horSec1RoutePlaneB = Plane.NegativeXY;
                    horSec1RouteAxisA = Axis.X;
                    horSec1RouteAxisB = Axis.X;

                    vertSec1RoutePlaneA = Plane.YZ;
                    vertSec1RoutePlaneB = Plane.ZX;
                    vertSec1RouteAxisA = Axis.Z;
                    vertSec1RouteAxisB = Axis.Z;
                    horSec1RoutePlOffset = -sectionDepth / 2;
                    horSec1RouteOrgOffset = 0;
                    horSec1RouteAxOffset = -boundingBoxWidth / 2;
                    verSec1RoutePlOffset = -sectionWidth / 2;
                    verSec1RouteOrgOffset = boundingBoxHeight / 2;
                    verSec1RouteAxOffset = -sectionDepth;
                }
                else if (Configuration == 2)
                {
                    horSec1RoutePlaneA = Plane.ZX;
                    horSec1RoutePlaneB = Plane.NegativeXY;
                    horSec1RouteAxisA = Axis.X;
                    horSec1RouteAxisB = Axis.NegativeX;

                    vertSec1RoutePlaneA = Plane.YZ;
                    vertSec1RoutePlaneB = Plane.ZX;
                    vertSec1RouteAxisA = Axis.Z;
                    vertSec1RouteAxisB = Axis.NegativeZ;
                    horSec1RouteAxOffset = -sectionWidth;
                    horSec1RouteOrgOffset = -sectionDepth / 2;
                }

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    routePort = "BBSR_Low";
                else
                    routePort = "BBR_Low";

                //Create Joints
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is placed only by Point.", "", "RectTypeR2.cs", 238);
                    return;
                }
                //Add joint between Horizontal Section 1 and Route
                JointHelper.CreateRigidJoint(HORSECTION1, "Neutral", "-1", "BBR_Low", horSec1RoutePlaneA, horSec1RoutePlaneB, horSec1RouteAxisA, horSec1RouteAxisB, horSec1RoutePlOffset, horSec1RouteAxOffset, 0);

                //Add joint between Horizontal Section 1 and Horizontal Section 2
                JointHelper.CreateRigidJoint(HORSECTION1, "Neutral", HORSECTION2, "Neutral", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -boundingBoxHeight - sectionWidth, 0);

                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(VERTSECTION1, "Neutral", "-1", "BBR_High", vertSec1RoutePlaneA, vertSec1RoutePlaneB, vertSec1RouteAxisA, vertSec1RouteAxisB, verSec1RoutePlOffset, verSec1RouteAxOffset, verSec1RouteOrgOffset);

                //Add joint between Horizontal Section 1 and Vertical Section 2
                JointHelper.CreateRigidJoint(VERTSECTION2, "Neutral", VERTSECTION1, "Neutral", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.Y, 0, -boundingBoxWidth - sectionWidth, 0);

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

