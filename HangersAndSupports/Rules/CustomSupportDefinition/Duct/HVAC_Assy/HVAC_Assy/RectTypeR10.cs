//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectTypeR10.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectTypeR10
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
using Ingr.SP3D.ReferenceData.Middle.Services;
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
    public class RectTypeR10 : CustomSupportDefinition
    {
        //Constants
        private const string HORSECTION1 = "HORSECTION1";
        private const string HORSECTION2 = "HORSECTION2";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";

        double overhangLeft, overhangRight;
        string sizeOfSection, steelStandard, sectionSize;
        bool sectionFromRule, overhangFromRule, value;

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
                 
                    overhangFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyOHRule", "OverhangFromRule")).PropValue;
                    overhangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyLOH", "OverhangLeft")).PropValue;
                    overhangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyROH", "OverhangRight")).PropValue;

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
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RectTypeR10.cs", 78);
                            return null;
                        }
                    }
                    string steelStd, sectionCode = string.Empty;
                     value = GenericHelper.GetDataByRule("HgrHVACSSteelStandardName", supportObject, out steelStd);

                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    ReadOnlyCollection<BusinessObject> classItems;
                    PartClass auxilaryTable = (PartClass)catalogBaseHelper.GetPartClass("HgrHVACStCorrespond");
                    classItems = auxilaryTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                    foreach (BusinessObject classItem in classItems)
                    {
                        if (((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrHVACStCorrespon", "SectionSize")).PropValue == sectionSize.ToString()) && ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrHVACStCorrespon", "StdName")).PropValue == steelStd))
                        {
                            sectionCode = (string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrHVACStCorrespon", "Size")).PropValue;
                            break;
                        }
                    }
                    if (overhangFromRule)
                    {
                        //Get LeftOverHang

                        PartClass auxilaryTable1 = (PartClass)catalogBaseHelper.GetPartClass("HVACAssy_OverhangDim");
                        classItems = auxilaryTable1.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                        foreach (BusinessObject classItem in classItems)
                        {
                            if (((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrHVACSrvOHDim", "SectionSize")).PropValue == sectionCode))
                            {
                                overhangLeft = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrHVACSrvOHDim", "LeftOverhang")).PropValue;
                                overhangRight = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrHVACSrvOHDim", "RightOverhang")).PropValue;
                                break;
                            }
                        }
                    }
                    else if (!overhangFromRule)
                    {
                        overhangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyLOH", "OverhangLeft")).PropValue;
                        overhangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyROH", "OverhangRight")).PropValue;
                    }

                    //GetSteel Standard               
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
              
                double boundingBoxWidth = boundingBox.Width,boundingBoxHeight = boundingBox.Height;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                Note noteDimenssion2 = CreateNote("Dim 2", componentDictionary[HORSECTION2], "BeginCap");
                noteDimenssion2.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem fabrication2 = noteDimenssion2.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                noteDimenssion2.SetPropertyValue(fabrication2, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                noteDimenssion2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

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

                //Set SectionSize attribute value on the support
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (sectionFromRule)
                {
                    CodelistItem codeList = metadataManager.GetCodelistInfo("HVACAssyLSecSize", "UDP").GetCodelistItem(sectionSize);
                    support.SetPropertyValue(codeList.Value, "IJOAHgrHVACAssyLSize", "LSectionSize");
                }
                // Miscallaneous
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                double routeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                string structPort1, structPort2;
                double horSec1Length, horSec2Length;


                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is placed only by Point.", "", "RectTypeR10.cs", 228);
                    return;
                }
                else
                {
                    horSec1Length = boundingBoxWidth + 2 * sectionWidth;
                    horSec2Length = horSec1Length;
                    componentDictionary[HORSECTION1].SetPropertyValue(horSec1Length, "IJUAHgrOccLength", "Length");
                    componentDictionary[HORSECTION2].SetPropertyValue(horSec2Length, "IJUAHgrOccLength", "Length");
                }

                componentDictionary[HORSECTION2].SetPropertyValue(overhangRight, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HORSECTION2].SetPropertyValue(overhangLeft, "IJUAHgrOccOverLength", "EndOverLength");

                string leftStructPort = string.Empty, rightStructPort = string.Empty;

                string[] structPortName = new string[2];
                if (SupportHelper.SupportingObjects.Count > 1)
                {
                    rightStructPort = "Structure";
                    leftStructPort = "Struct_2";

                    structPortName[0] = rightStructPort;
                    structPortName[1] = leftStructPort;

                    if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                    {
                        Double angle1;
                        angle1 = HVACAssemblyServices.GetVectorProjectioBetweenPorts(this, "Route", "Structure", PortAxisType.Y);

                        if (angle1 > 0)
                        {
                            structPortName[0] = leftStructPort;
                            structPortName[1] = rightStructPort;
                        }
                        else
                        {
                            structPortName[0] = rightStructPort;
                            structPortName[1] = leftStructPort;
                        }

                    }
                }
                    structPort1 = structPortName[0];
                    structPort2 = structPortName[1];

                    Double angle2;
                    angle2 = RefPortHelper.AngleBetweenPorts( "Route",PortAxisType.X, structPortName[0], PortAxisType.X,OrientationAlong.Direct);

                    String horSecPort = string.Empty, routePort = string.Empty;

                    Plane horSec1RoutePlaneA = new Plane(), horSec1RoutePlaneB = new Plane(), horSec2RoutePlaneA = new Plane(), horSec2RoutePlaneB = new Plane();
                    Axis horSec1RouteAxisA = new Axis(), horSec1RouteAxisB = new Axis(), horSec2RouteAxisA = new Axis(), horSec2RouteAxisB = new Axis();
                    Double horSec1RouteAxOffset = 0;

                    if (HgrCompareDoubleService.cmpdbl(Math.Round(angle2, 4), Math.Round(Math.PI, 4)) == false)
                    {
                        horSec2RoutePlaneA = Plane.YZ;
                        horSec2RoutePlaneB = Plane.XY;
                        horSec2RouteAxisA = Axis.Y;
                        horSec2RouteAxisB = Axis.NegativeX;
                    }
                    else
                    {
                        horSec2RoutePlaneA = Plane.YZ;
                        horSec2RoutePlaneB = Plane.XY;
                        horSec2RouteAxisA = Axis.Y;
                        horSec2RouteAxisB = Axis.X;
                    }

                    if (Configuration == 1)
                    {
                        horSec1RoutePlaneA = Plane.ZX;
                        horSec1RoutePlaneB = Plane.NegativeXY;
                        horSec1RouteAxisA = Axis.X;
                        horSec1RouteAxisB = Axis.X;

                        horSec1RouteAxOffset = sectionWidth;
                        horSecPort = "BeginCap";
                        componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                    }
                    else if (Configuration == 2)
                    {
                        horSec1RoutePlaneA = Plane.ZX;
                        horSec1RoutePlaneB = Plane.NegativeXY;
                        horSec1RouteAxisA = Axis.X;
                        horSec1RouteAxisB = Axis.NegativeX;

                        horSec1RouteAxOffset = -sectionWidth;
                        horSecPort = "EndCap";
                        componentDictionary[HORSECTION2].SetPropertyValue(Math.PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                    }


                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        routePort = "BBSR_Low";
                    else
                        routePort = "BBR_Low";

                //Create Joints
                //Add joint between Horizontal Section 1 and Route
                JointHelper.CreateRigidJoint(HORSECTION1, horSecPort, "-1", routePort, horSec1RoutePlaneA, horSec1RoutePlaneB, horSec1RouteAxisA, horSec1RouteAxisB, 0, horSec1RouteAxOffset, 0);

                //Joint between EndCap of top horizontal and Structure ports
                JointHelper.CreateRigidJoint(HORSECTION2, "BeginCap", "-1", structPort1, horSec2RoutePlaneA, horSec2RoutePlaneB, horSec2RouteAxisA, horSec2RouteAxisB, 0, 0, 0);

                // -- Joint between BeginCap and EndCap of top horizontal
                JointHelper.CreatePrismaticJoint(HORSECTION2, "BeginCap", HORSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                // Joint between BeginCap and Struct_2
                JointHelper.CreatePointOnAxisJoint(HORSECTION2, "EndCap", "-1", structPort2, Axis.Z);

                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(VERTSECTION1, "BeginCap", HORSECTION1, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX, sectionDepth, sectionDepth, 0);

                //--Add Joint between BeginCap and EndCap of Vertical Section 1
                JointHelper.CreatePrismaticJoint(VERTSECTION1, "BeginCap", VERTSECTION1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //--Add Joint between EndCap of Vertical Section 1 and End Cap of top horizontal
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION1, "EndCap", HORSECTION2, "EndCap", Plane.YZ);

                //Add joint between Horizontal Section 1 and Vertical Section 2
                JointHelper.CreateRigidJoint(VERTSECTION2, "EndCap", HORSECTION1, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, sectionDepth, -sectionDepth, 0);

                //--Add Joint between BeginCap and EndCap of Vertical Section 2
                JointHelper.CreatePrismaticJoint(VERTSECTION2, "BeginCap", VERTSECTION2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //--Add Joint between BeginCaps of Verical Section 1 and top horizontal
                JointHelper.CreatePointOnPlaneJoint(VERTSECTION2, "BeginCap", HORSECTION2, "BeginCap", Plane.YZ);
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



