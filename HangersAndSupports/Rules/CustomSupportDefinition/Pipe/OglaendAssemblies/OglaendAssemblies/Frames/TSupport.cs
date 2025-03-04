//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   TSupport.cs
//   OglaendAssemblies,Ingr.SP3D.Content.Support.Rules.TSupport
//   Author       : Chethan
//   Creation Date:  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   23-01-2015      Chethan  CR-CP-241198  Convert all Oglaend Content to .NET  
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{

    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class TSupport : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        #region Declarations
        //Steel Sections

        private const string HOR_SECTION = "HOR_SECTION";
        private const string VER_SECTION = "VER_SECTION";

        //Other parts

        private const string WELDED_STARTER = "WELDED_STARTER";
        private const string END_PROTECTION1 = "END_PROTECTION1";
        private const string END_PROTECTION2 = "END_PROTECTION2";
        private const string STRUCT_CONN = "STRUCT_CONN";

        //Steel Members - Stores Data related to steel member (Part Number, Depth, Width etc...)
        private OglaendAssemblyServices.HSSteelMember HorSection;
        private OglaendAssemblyServices.HSSteelMember VerSection;        

        int[] structureIndex = new int[2];
        Boolean isMember1, isMember2, isEndProtection1, isEndProtection2, isweldedstarter;

        #endregion Declarations

        #region CatalogParts

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                    string strRule = "";
                    BusinessObject ruleBO = null;
                    // Add the Steel Section for the Frame Support

                    string member1Part = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsOGHorSection", "HorSection")).PropValue;
                    isMember1 = OglaendAssemblyServices.AddPart(this, HOR_SECTION, member1Part, strRule, parts, ruleBO);

                    string member2Part = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsOGHorSection", "HorSection")).PropValue;
                    isMember2 = OglaendAssemblyServices.AddPart(this, VER_SECTION, member2Part, strRule, parts, ruleBO);

                    // Add the Plates
                    string WeldedStarter = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsOGWeldedStarter", "WeldedStarter")).PropValue;
                    isweldedstarter = OglaendAssemblyServices.AddPart(this, WELDED_STARTER, WeldedStarter, null, parts);

                    string EndProtection1 = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsOGEndProtection", "EndProtection")).PropValue;
                    isEndProtection1 = OglaendAssemblyServices.AddPart(this, END_PROTECTION1, EndProtection1, null, parts);

                    string EndProtection2 = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsOGEndProtection", "EndProtection")).PropValue;
                    isEndProtection2 = OglaendAssemblyServices.AddPart(this, END_PROTECTION2, EndProtection2, null, parts);

                    //Add the Structure Connection Object
                    OglaendAssemblyServices.AddPart(this, STRUCT_CONN,"Log_Conn_Part_1", null, parts);

                    // Return the collection of Catalog Parts
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

        #endregion CatalogParts

        #region Get Assembly Implied Parts
        //-----------------------------------------------------------------------------------
        //Get Assembly Implied Parts
        //-----------------------------------------------------------------------------------
        public override ReadOnlyCollection<PartInfo> ImpliedParts
        {
            get
            {
                try
                {
                    SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> impliedParts = new Collection<PartInfo>();
                    BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                    string supportpartnumber = support.SupportDefinition.PartNumber;
                    OglaendAssemblyServices.AddImpliedPartbyInterface(this, impliedParts, "IJUAhsFrImpPart", supportpartnumber, 10005);
                    ReadOnlyCollection<PartInfo> rImpliedParts = new ReadOnlyCollection<PartInfo>(impliedParts);
                    return rImpliedParts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Implied Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }

        #endregion Get Assembly Implied Parts

        #region Get Aseembly Joints
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                SP3D.Support.Middle.Support support = SupportHelper.Support;

                // Get interface for accessing items on the collection of Part Occurences
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

                // ==========================
                // Get Required Information about the Steel Parts
                // ==========================

                HorSection = OglaendAssemblyServices.GetSectionDataFromPartIndex(this, HOR_SECTION);
                VerSection = OglaendAssemblyServices.GetSectionDataFromPartIndex(this, VER_SECTION);

                // ==========================
                // Create the Frame Bounding Box
                // ==========================

                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;

                boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double boundingBoxWidth = boundingBox.Width;
                double boundingBoxHeight = boundingBox.Height;

                // ==========================
                // Determine the number of Supporting Objects
                // ==========================
                int supportedCount = 0;
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    // For Place-By-Reference Command No Supporting Objects are Selected
                    // However, we want to treat it like 1 was selected
                    supportedCount = 1;
                else
                    supportedCount = SupportHelper.SupportedObjects.Count;                

                Double Offset1 = 0, CantileverLength = 0, PortOrientationAngle = 0, horOffsetX = 0, horOffsetY = 0, Wlength1 = 0, StructureIntPosZ = 0, WeldFlangeWidth = 0, WeldThickness, WeldedStarterThickness = 0, temp = 0,
                        Routestructdist = 0, ShoeHeight = 0, OrientationAngle = 0, Offset = 0, HorAxisOffset = 0, HorOriginOffset = 0, HorPlaneOffset = 0, ProjectionY = 0, ProjectionZ = 0, Width = 0, temp1 = 0, MaximumHeight=0,
                         MaximumWidth = 0, CapacityDWPy = 0, PipingPy = 0, PipingPx = 0, BlastPy = 0, BlastPx = 0, MinimumWidth=0;

                Plane horSec1RoutePlaneA = new Plane(), horSec1RoutePlaneB = new Plane();
                Axis horSec1RouteAxisA = new Axis(), horSec1RouteAxisB = new Axis();

                String SupportNumber;

                Boolean AutoAdjust;

                Offset1 = (Double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsOGOffset", "Offset")).PropValue;
                SupportNumber = (String)((PropertyValueString)part.GetPropertyValue("IJUAhsSupportNumber", "SupportNumber")).PropValue;
                CantileverLength = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight", "Height")).PropValue;

                try
                {
                     AutoAdjust = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsLengthOverride", "LengthOverride")).PropValue;
                }
                catch
                {
                    AutoAdjust = false;
                }

                PortOrientationAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

                //Get and Set Part Attributes

                PropertyValueCodelist cardinalPoint6HorSection = (PropertyValueCodelist)componentDictionary[HOR_SECTION].GetPropertyValue("IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint7VerSection = (PropertyValueCodelist)componentDictionary[VER_SECTION].GetPropertyValue("IJOAhsSteelCPFlexPort", "CP7");
                PropertyValueCodelist cardinalPoint6VerSection = (PropertyValueCodelist)componentDictionary[VER_SECTION].GetPropertyValue("IJOAhsSteelCP", "CP6");

                PropertyValue oWlength1 = OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELDED_STARTER], "IJUAHgrOGLength", "Length", OglaendAssemblyServices.FromCatalogOrOccurrence.FromCatalog, false);
                Wlength1 = (double)(((PropertyValueDouble)oWlength1).PropValue);

                PropertyValue oStructureIntPosZ = OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELDED_STARTER], "IJUAHgrOGStructureIntPort", "StuctureIntPosZ", OglaendAssemblyServices.FromCatalogOrOccurrence.FromCatalog, false);
                StructureIntPosZ = (double)(((PropertyValueDouble)oStructureIntPosZ).PropValue);

                PropertyValue oWeldFlangeWidth = OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELDED_STARTER], "IJUAHgrOGFlangeWidth", "FlangeWidth", OglaendAssemblyServices.FromCatalogOrOccurrence.FromCatalog, false);
                WeldFlangeWidth = (double)(((PropertyValueDouble)oWeldFlangeWidth).PropValue);

                PropertyValue oThickness = OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELDED_STARTER], "IJUAHgrOGThickness", "Thickness", OglaendAssemblyServices.FromCatalogOrOccurrence.FromCatalog, false);
                WeldThickness = (double)(((PropertyValueDouble)oThickness).PropValue);

                if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                {
                    componentDictionary[VER_SECTION].SetPropertyValue(0.5, "IJUAHgrOccLength", "Length");
                    componentDictionary[VER_SECTION].SetPropertyValue(Math.PI/2, "IJOAhsFlexPort", "FlexPortRotZ");
                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint6VerSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                    componentDictionary[WELDED_STARTER].SetPropertyValue((double)0, "IJUAHgrOGStructureIntPort", "StuctureIntPosZ");

                    horOffsetY = 0;

                    horOffsetX = Wlength1/2 - StructureIntPosZ;
                }

                else if (SupportNumber == "MGT03")
                {
                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)5, "IJOAhsSteelCPFlexPort", "CP7");
                    componentDictionary[VER_SECTION].SetPropertyValue(0.5, "IJUAHgrOccLength", "Length");
                    componentDictionary[VER_SECTION].SetPropertyValue((double)0, "IJOAhsFlexPort", "FlexPortRotZ");
                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint6VerSection.PropValue = (int)5, "IJOAhsSteelCP", "CP6");

                    horOffsetY = 0.00175;
                    horOffsetX = Wlength1 / 2;
                    WeldedStarterThickness = WeldThickness;
                }
                else
                {
                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)5, "IJOAhsSteelCPFlexPort", "CP7");
                    componentDictionary[VER_SECTION].SetPropertyValue(0.5, "IJUAHgrOccLength", "Length");
                    componentDictionary[VER_SECTION].SetPropertyValue((double)0, "IJOAhsFlexPort", "FlexPortRotZ");
                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint6VerSection.PropValue = (int)5, "IJOAhsSteelCP", "CP6");

                    horOffsetY = 0.00175;
                    horOffsetX = Wlength1 / 2;
                    WeldedStarterThickness = WeldThickness;

                }

                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                if (supportingType == "Steel")
                {

                    if ((PortOrientationAngle > Math.PI / 4 & PortOrientationAngle < 3 * Math.PI / 4) || (PortOrientationAngle > 5 * Math.PI / 4 & PortOrientationAngle < 7 * Math.PI / 7))
                    {
                        if (HgrCompareDoubleService.cmpdbl(CantileverLength,0) == true)
                        {                            
                            CantileverLength = RefPortHelper.DistanceBetweenPorts("BBSR_Low", "Structure", PortAxisType.Z)+HorSection.depth;
                            support.SetPropertyValue(Math.Abs(CantileverLength), "IJUAhsOGDummyHeight", "DummyHeight");
                            support.SetPropertyValue(CantileverLength, "IJUAhsOGHeight", "Height");
                        }
                        else if (AutoAdjust == true)
                        {
                            temp = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGDummyHeight", "DummyHeight")).PropValue;                            
                            Routestructdist = Math.Abs(RefPortHelper.DistanceBetweenPorts("BBSR_Low", "Structure", PortAxisType.Z));                            
                            CantileverLength = RefPortHelper.DistanceBetweenPorts("BBSR_Low", "Structure", PortAxisType.Z)+HorSection.depth;
                            support.SetPropertyValue(Math.Abs(CantileverLength), "IJUAhsOGDummyHeight", "DummyHeight");

                            if (HgrCompareDoubleService.cmpdbl(temp, Routestructdist) == true)
                            {
                                CantileverLength = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight", "Height")).PropValue;
                            }
                            else
                            {                                
                                CantileverLength = RefPortHelper.DistanceBetweenPorts("BBSR_Low", "Structure", PortAxisType.Z) + ShoeHeight;
                                support.SetPropertyValue(CantileverLength, "IJUAhsOGHeight", "Height");
                            }
                        }
                        else
                        {
                            temp = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGDummyHeight", "DummyHeight")).PropValue;                            
                            Routestructdist = Math.Abs(RefPortHelper.DistanceBetweenPorts("BBSR_Low", "Structure", PortAxisType.Z));
                            support.SetPropertyValue(Routestructdist, "IJUAhsOGDummyHeight", "DummyHeight");
                            CantileverLength = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight", "Height")).PropValue;
                        }
                        if (HgrCompareDoubleService.cmpdbl(Offset1,0)==true)
                        {
                            if (supportedCount == 2)
                            {                                
                                support.SetPropertyValue((RefPortHelper.DistanceBetweenPorts("Route", "Route_2", PortAxisType.Y)) / 2, "IJOAhsOGOffset", "Offset");
                            }                           
                            else
                            {
                                support.SetPropertyValue(boundingBoxWidth / 2, "IJOAhsOGOffset", "Offset");
                            }
                        }

                        Offset1 = (Double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsOGOffset", "Offset")).PropValue;

                        #region sup config
                        if (supportedCount == 1)
                        {
                            if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                            {
                                if (Configuration == 1)
                                {
                                    OrientationAngle = Math.PI / 2;
                                    Offset = Offset1 + ((WeldFlangeWidth - WeldedStarterThickness) / 2) - WeldedStarterThickness;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = -VerSection.depth / 2;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 2)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = Offset1 + ((WeldFlangeWidth - WeldedStarterThickness) / 2) - WeldedStarterThickness;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = 0;
                                    HorOriginOffset = -VerSection.depth / 2;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 3)
                                {
                                    OrientationAngle = 3 * Math.PI / 2;
                                    Offset = -Offset1 - ((WeldFlangeWidth - WeldedStarterThickness) / 2) + WeldedStarterThickness;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = -VerSection.depth / 2;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 4)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1 - ((WeldFlangeWidth - WeldedStarterThickness) / 2) + WeldedStarterThickness;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = 0;
                                    HorOriginOffset = -VerSection.depth / 2;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 5)
                                {
                                    OrientationAngle = Math.PI / 2;
                                    Offset = Offset1 + ((WeldFlangeWidth - WeldedStarterThickness) / 2) - WeldedStarterThickness;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = -VerSection.depth / 2;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 6)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = Offset1 + ((WeldFlangeWidth - WeldedStarterThickness) / 2) - WeldedStarterThickness;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = 0;
                                    HorOriginOffset = -VerSection.depth / 2;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 7)
                                {
                                    OrientationAngle = 3 * Math.PI / 2;
                                    Offset = -Offset1 - ((WeldFlangeWidth - WeldedStarterThickness) / 2) + WeldedStarterThickness;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = -VerSection.depth / 2;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 8)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1 - ((WeldFlangeWidth - WeldedStarterThickness) / 2) + WeldedStarterThickness;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = 0;
                                    HorOriginOffset = -VerSection.depth / 2;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                            }
                            else if (SupportNumber == "MGT03")
                            {
                                if (Configuration == 1)
                                {
                                    OrientationAngle = 0;
                                    Offset = Offset1 + HorSection.depth / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 2)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = Offset1 + HorSection.depth / 2 - horOffsetY;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 3)
                                {
                                    OrientationAngle = 0;
                                    Offset = Offset1 + HorSection.depth / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 4)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = Offset1 + HorSection.depth / 2 + horOffsetY;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 5)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1 - HorSection.depth / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 6)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = -Offset1 - HorSection.depth / 2 - horOffsetY;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 7)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1 - HorSection.depth / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 8)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = -Offset1 - HorSection.depth / 2 + horOffsetY;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                            }
                            else
                            {
                                if (Configuration == 1)
                                {
                                    OrientationAngle = Math.PI / 2;
                                    Offset = Offset1 + HorSection.depth / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 2)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = Offset1 + HorSection.depth / 2 - WeldedStarterThickness / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 3)
                                {
                                    OrientationAngle = 3 * Math.PI / 2;
                                    Offset = -Offset1 - HorSection.depth / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 4)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1 - HorSection.depth / 2 + WeldedStarterThickness / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 5)
                                {
                                    OrientationAngle = Math.PI / 2;
                                    Offset = Offset1 + HorSection.depth / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 6)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = Offset1 + HorSection.depth / 2 - WeldedStarterThickness / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 7)
                                {
                                    OrientationAngle = 3 * Math.PI / 2;
                                    Offset = -Offset1 - HorSection.depth / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 8)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1 - HorSection.depth / 2 + WeldedStarterThickness / 2;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                            }
                        }
                        else
                        {
                            if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                            {
                                if (Configuration == 1)
                                {
                                    OrientationAngle = Math.PI / 2;
                                    Offset = -Offset1;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = -VerSection.depth / 2;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 2)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = -Offset1;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = 0;
                                    HorOriginOffset = -VerSection.depth / 2;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 3)
                                {
                                    OrientationAngle = 3 * Math.PI / 2;
                                    Offset = -Offset1;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = -VerSection.depth / 2;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 4)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = 0;
                                    HorOriginOffset = -VerSection.depth / 2;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 5)
                                {
                                    OrientationAngle = Math.PI / 2;
                                    Offset = -Offset1;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = -VerSection.depth / 2;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 6)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = -Offset1;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = 0;
                                    HorOriginOffset = -VerSection.depth / 2;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 7)
                                {
                                    OrientationAngle = 3 * Math.PI / 2;
                                    Offset = -Offset1;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = -VerSection.depth / 2;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 8)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1;
                                    componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                    HorAxisOffset = 0;
                                    HorOriginOffset = -VerSection.depth / 2;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                            }
                            else if (SupportNumber == "MGT03")
                            {
                                if (Configuration == 1)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 2)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 3)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 4)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 5)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 6)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 7)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 8)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                            }
                            else
                            {
                                if (Configuration == 1)
                                {
                                    OrientationAngle = Math.PI / 2;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 2)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 3)
                                {
                                    OrientationAngle = 3 * Math.PI / 2;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 4)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 5)
                                {
                                    OrientationAngle = Math.PI / 2;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 6)
                                {
                                    OrientationAngle = Math.PI;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 7)
                                {
                                    OrientationAngle = 3 * Math.PI / 2;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.X;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                                else if (Configuration == 8)
                                {
                                    OrientationAngle = 0;
                                    Offset = -Offset1;
                                    HorAxisOffset = 0;
                                    HorOriginOffset = 0;
                                    HorPlaneOffset = -VerSection.depth / 2;
                                    horSec1RoutePlaneA = Plane.ZX;
                                    horSec1RoutePlaneB = Plane.XY;
                                    horSec1RouteAxisA = Axis.Z;
                                    horSec1RouteAxisB = Axis.NegativeX;
                                }
                            }
                        }  
                        #endregion change sup
                        if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                        {
                            JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, Offset);
                            JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, OrientationAngle));
                            JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                        }
                        else
                        {
                            JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0,Offset);
                            JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, OrientationAngle));
                            JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -horOffsetX, 0, -horOffsetY);
                        }
                    }                    
                    else
                    {                        
                        ProjectionY = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortAxisType.Y);                        
                        ProjectionZ = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortAxisType.Z);

                        if (HgrCompareDoubleService.cmpdbl(CantileverLength, 0) == true)
                        {

                            CantileverLength = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z) + HorSection.depth;
                            support.SetPropertyValue(Math.Abs(CantileverLength), "IJUAhsOGDummyHeight", "DummyHeight");
                            support.SetPropertyValue(CantileverLength, "IJUAhsOGHeight", "Height");
                        }
                        else if (AutoAdjust == true)
                        {
                            temp = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGDummyHeight", "DummyHeight")).PropValue;                            
                            Routestructdist = Math.Abs(RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z));                            
                            CantileverLength = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z) + HorSection.depth;
                            support.SetPropertyValue(Math.Abs(CantileverLength), "IJUAhsOGDummyHeight", "DummyHeight");

                            if (HgrCompareDoubleService.cmpdbl(temp, Routestructdist) == true)
                            {
                                CantileverLength = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight", "Height")).PropValue;
                            }
                            else
                            {                                
                                CantileverLength = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z) + ShoeHeight;
                                support.SetPropertyValue(CantileverLength, "IJUAhsOGHeight", "Height");
                            }
                        }
                        else
                        {
                            temp = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGDummyHeight", "DummyHeight")).PropValue;                            
                            Routestructdist = Math.Abs(RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z));
                            support.SetPropertyValue(Routestructdist, "IJUAhsOGDummyHeight", "DummyHeight");
                            CantileverLength = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight", "Height")).PropValue;
                        }

                        //Get normal vector along the Surface

                        Vector structFaceNormal = new Vector();
                        Vector structZ = new Vector();
                        Matrix4X4 hangerPort = new Matrix4X4();
                        IPlane iplane = (IPlane)support.SupportingFaces[0];

                        structFaceNormal = new Vector(iplane.Normal.X, iplane.Normal.Y, iplane.Normal.Z);

                        hangerPort = new Matrix4X4();
                        hangerPort = RefPortHelper.PortLCS("Structure");
                        structZ = new Vector(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);

                        if (HgrCompareDoubleService.cmpdbl(Offset1, 0) == true)
                        {
                            Offset1 = 0;
                        }
                        else
                        {
                            Offset1 = (Double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsOGOffset", "Offset")).PropValue;
                        }
                        #region Sup config
                        if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                        {
                            if (Configuration == 1)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 2)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 3)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 4)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 5)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 6)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 7)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 8)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                        }
                        else if (SupportNumber == "MGT03")
                        {
                            if (Configuration == 1)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 2)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 3)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 4)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 5)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 6)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 7)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 8)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                        }
                        else
                        {
                            if (Configuration == 1)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 2)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 3)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 4)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 5)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 6)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 7)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 8)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                        }                        
                        #endregion sup config

                        if (Math.Abs(ProjectionZ) > Math.Abs(ProjectionY))
                        {
                            if (OglaendAssemblyServices.AngleBetweenVectors(structFaceNormal, structZ) < Math.PI / 2)
                            {
                                if (ProjectionY > 0)
                                {
                                    if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                                    }
                                    else
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -horOffsetX, 0, -horOffsetY);
                                    }                                    
                                }
                                else
                                {
                                    if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0,-Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                                    }
                                    else
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0,-Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -horOffsetX, 0, -horOffsetY);
                                    }                                    
                                }
                            }
                            else
                            {
                                if (ProjectionY > 0)
                                {
                                    if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0,Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                                    }
                                    else
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -horOffsetX, 0, -horOffsetY);
                                    }
                                }
                                else
                                {
                                    if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0,-Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                                    }
                                    else
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -horOffsetX, 0, -horOffsetY);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (ProjectionY > 0)
                            {
                                if (OglaendAssemblyServices.AngleBetweenVectors(structFaceNormal, structZ) < Math.PI / 2)
                                {
                                    if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                                    }
                                    else
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -horOffsetX, 0, -horOffsetY);
                                    }
                                }
                                else
                                {
                                    if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                                    }
                                    else
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0,Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -horOffsetX, 0, -horOffsetY);
                                    }
                                }
                            }
                            else
                            {
                                if (OglaendAssemblyServices.AngleBetweenVectors(structFaceNormal, structZ) < Math.PI / 2)
                                {
                                    if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0,-Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                                    }
                                    else
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0,-Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -horOffsetX, 0, -horOffsetY);
                                    }
                                }
                                else
                                {
                                    if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0,-Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                                    }
                                    else
                                    {
                                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0,-Offset,0);
                                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, Math.PI / 2 + OrientationAngle));
                                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -horOffsetX, 0, -horOffsetY);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    //wall, Salb and Shape
                    if (HgrCompareDoubleService.cmpdbl(CantileverLength, 0) == true)
                    {                        
                        CantileverLength = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z) + HorSection.depth;
                        support.SetPropertyValue(Math.Abs(CantileverLength), "IJUAhsOGDummyHeight", "DummyHeight");
                        support.SetPropertyValue(CantileverLength, "IJUAhsOGHeight", "Height");
                    }
                    else if (AutoAdjust == true)
                    {
                        temp = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGDummyHeight", "DummyHeight")).PropValue;                        
                        Routestructdist = Math.Abs(RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z));                        
                        CantileverLength = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z) + HorSection.depth;
                        support.SetPropertyValue(Math.Abs(CantileverLength), "IJUAhsOGDummyHeight", "DummyHeight");
                        if (HgrCompareDoubleService.cmpdbl(temp, Routestructdist) == true)
                        {
                            CantileverLength = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight", "Height")).PropValue;
                        }
                        else
                        {                            
                            CantileverLength = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z) + ShoeHeight;
                            support.SetPropertyValue(CantileverLength,"IJUAhsOGHeight", "Height");
                        }
                    }
                    else
                    {
                        temp = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGDummyHeight", "DummyHeight")).PropValue;                        
                        Routestructdist = Math.Abs(RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z));
                        support.SetPropertyValue(Routestructdist, "IJUAhsOGDummyHeight", "DummyHeight");
                        CantileverLength = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight", "Height")).PropValue;
                    }
                    if (HgrCompareDoubleService.cmpdbl(Offset1, 0) == true)
                    {
                        if (supportedCount == 2)
                        {
                            support.SetPropertyValue(RefPortHelper.DistanceBetweenPorts("Route", "Route_2", PortAxisType.Y) / 2, "IJOAhsOGOffset", "Offset");
                        }
                        else 
                        {
                            
                            support.SetPropertyValue(boundingBoxWidth / 2, "IJOAhsOGOffset", "Offset");
                        }
                    }
                    Offset1 = (Double)((PropertyValueDouble)support.GetPropertyValue("IJOAhsOGOffset", "Offset")).PropValue;
                    #region change config
                    if (supportedCount == 1)
                    {
                        if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                        {
                            if (Configuration == 1)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1 + ((WeldFlangeWidth - WeldedStarterThickness) / 2) - WeldedStarterThickness;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 2)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1 + ((WeldFlangeWidth - WeldedStarterThickness) / 2) - WeldedStarterThickness;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 3)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = -Offset1 - ((WeldFlangeWidth - WeldedStarterThickness) / 2) + WeldedStarterThickness;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 4)
                            {
                                OrientationAngle = 0;
                                Offset = -Offset1 - ((WeldFlangeWidth - WeldedStarterThickness) / 2) + WeldedStarterThickness;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 5)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1 + ((WeldFlangeWidth - WeldedStarterThickness) / 2) - WeldedStarterThickness;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 6)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1 + ((WeldFlangeWidth - WeldedStarterThickness) / 2) - WeldedStarterThickness;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 7)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = -Offset1 - ((WeldFlangeWidth - WeldedStarterThickness) / 2) + WeldedStarterThickness;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 8)
                            {
                                OrientationAngle = 0;
                                Offset = -Offset1 - ((WeldFlangeWidth - WeldedStarterThickness) / 2) + WeldedStarterThickness;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                        }
                        else if (SupportNumber == "MGT03")
                        {
                            if (Configuration == 1)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1 + HorSection.depth / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 2)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1 + HorSection.depth / 2 - horOffsetY;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 3)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1 + HorSection.depth / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 4)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1 + HorSection.depth / 2 + horOffsetY;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 5)
                            {
                                OrientationAngle = 0;
                                Offset = -Offset1 - HorSection.depth / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 6)
                            {
                                OrientationAngle = Math.PI;
                                Offset = -Offset1 - HorSection.depth / 2 - horOffsetY;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 7)
                            {
                                OrientationAngle = 0;
                                Offset = -Offset1 - HorSection.depth / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 8)
                            {
                                OrientationAngle = Math.PI;
                                Offset = -Offset1 - HorSection.depth / 2 + horOffsetY;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                        }
                        else
                        {
                            if (Configuration == 1)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1 + HorSection.depth / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 2)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1 + HorSection.depth / 2 - WeldedStarterThickness / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 3)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = -Offset1 - HorSection.depth / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 4)
                            {
                                OrientationAngle = 0;
                                Offset = -Offset1 - HorSection.depth / 2 + WeldedStarterThickness / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 5)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1 + HorSection.depth / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 6)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1 + HorSection.depth / 2 - WeldedStarterThickness / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 7)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = -Offset1 - HorSection.depth / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 8)
                            {
                                OrientationAngle = 0;
                                Offset = -Offset1 - HorSection.depth / 2 + WeldedStarterThickness / 2;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                        }
                    }
                    else
                    {
                        if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                        {
                            if (Configuration == 1)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 2)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 3)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 4)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 5)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 6)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 7)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)2, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = -VerSection.depth / 2;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 8)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                componentDictionary[HOR_SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                                componentDictionary[VER_SECTION].SetPropertyValue(cardinalPoint7VerSection.PropValue = (int)4, "IJOAhsSteelCPFlexPort", "CP7");
                                HorAxisOffset = 0;
                                HorOriginOffset = -VerSection.depth / 2;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                        }
                        else if (SupportNumber == "MGT03")
                        {
                            if (Configuration == 1)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 2)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 3)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 4)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 5)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 6)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 7)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 8)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                        }
                        else
                        {
                            if (Configuration == 1)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 2)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 3)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 4)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 5)
                            {
                                OrientationAngle = Math.PI / 2;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 6)
                            {
                                OrientationAngle = Math.PI;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 7)
                            {
                                OrientationAngle = 3 * Math.PI / 2;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.X;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                            else if (Configuration == 8)
                            {
                                OrientationAngle = 0;
                                Offset = Offset1;
                                HorAxisOffset = 0;
                                HorOriginOffset = 0;
                                HorPlaneOffset = -VerSection.depth / 2;
                                horSec1RoutePlaneA = Plane.ZX;
                                horSec1RoutePlaneB = Plane.XY;
                                horSec1RouteAxisA = Axis.Z;
                                horSec1RouteAxisB = Axis.NegativeX;
                            }
                        }
                    }
                    #endregion config
                    if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                    {
                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0,Offset,0);
                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, OrientationAngle - Math.PI/2));
                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, Offset,0);
                        JointHelper.CreateAngularRigidJoint(WELDED_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, OrientationAngle - Math.PI/2));
                        JointHelper.CreateRigidJoint(VER_SECTION, "BeginFlex", "WELDED_STARTER", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -horOffsetX, 0, -horOffsetY);
                    }                    
                }

                JointHelper.CreateRigidJoint(HOR_SECTION, "Neutral", VER_SECTION, "EndFlex", horSec1RoutePlaneA, horSec1RoutePlaneB, horSec1RouteAxisA, horSec1RouteAxisB, HorPlaneOffset, HorAxisOffset, HorOriginOffset);

                if (isEndProtection1)
                {
                    JointHelper.CreateRigidJoint(HOR_SECTION,"EndFace",END_PROTECTION1,"Port1",Plane.XY,Plane.ZX,Axis.X,Axis.X,0,0,0);
                }
                if (isEndProtection2)
                {
                    JointHelper.CreateRigidJoint(HOR_SECTION, "BeginFace", END_PROTECTION2, "Port2", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, 0);
                }

                //Adjust the Width of T Support

                Width = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGWidth", "Width")).PropValue;

                if (HgrCompareDoubleService.cmpdbl(Width, 0) == true)
                {
                    support.SetPropertyValue(boundingBoxWidth, "IJUAhsOGDummyWidth", "DummyWidth");

                    if (supportedCount == 1)
                    {
                        support.SetPropertyValue(3 * boundingBoxWidth, "IJUAhsOGWidth", "Width");
                    }
                    else if (supportedCount == 2)
                    {
                        support.SetPropertyValue(Math.Abs(boundingBoxWidth), "IJUAhsOGWidth", "Width");
                    }
                    else
                    {
                        support.SetPropertyValue(boundingBoxWidth, "IJUAhsOGWidth", "Width");
                    }
                }
                else
                {
                    temp1 = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGDummyWidth", "DummyWidth")).PropValue;

                    support.SetPropertyValue(boundingBoxWidth, "IJUAhsOGDummyWidth", "DummyWidth");

                    
                    if (HgrCompareDoubleService.cmpdbl(temp1, boundingBoxWidth) == true)
                    {
                        Width = (Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGWidth", "Width")).PropValue;
                    }
                    else
                    {
                        if (supportedCount == 1)
                        {
                            support.SetPropertyValue(3 * boundingBoxWidth, "IJUAhsOGWidth", "Width");
                        }
                        else if (supportedCount == 2)
                        {
                            support.SetPropertyValue(Math.Abs(boundingBoxWidth), "IJUAhsOGWidth", "Width");
                        }
                        else
                        {
                            support.SetPropertyValue(boundingBoxWidth, "IJUAhsOGWidth", "Width");
                        }
                    }
                }                

                //Check for Maximum Length and Query For Loads
                    
                Width=(Double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGWidth","Width")).PropValue;                

                string svalue;
                bool isEqual;

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass auxTable = (PartClass)cataloghelper.GetPartClass("HS_OGAssy_MaxValues");
                ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in classItems)
                {
                    svalue = ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGMaxValue", "SupportNumber")).PropValue);
                    isEqual = String.Equals(SupportNumber, ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGMaxValue", "SupportNumber")).PropValue), StringComparison.Ordinal);
                    if (isEqual == true)
                    {
                        MaximumHeight = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGMaxValue", "MaximumHeight")).PropValue);
                        MaximumWidth = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGMaxValue", "MaximumWidth")).PropValue);
                        break;

                    }

                }

                if (CantileverLength > MaximumHeight)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinMaxHeight,"Total Height of T Support is greater than the Maximum allowable Height.Hence,resetting the height to maximum height."), "", "TSupport", 346);
                    CantileverLength = MaximumHeight;
                    support.SetPropertyValue(CantileverLength, "IJUAhsOGHeight", "Height");
                    try
                    {
                        part.SetPropertyValue((double)0, "IJOAhsOGShoeHeight", "ShoeHeight");
                    }
                    catch
                    {
                        ShoeHeight = 0;
                    }
                }


                PartClass auxTable1 = (PartClass)cataloghelper.GetPartClass("HS_OGAssy_MinValues");
                ReadOnlyCollection<BusinessObject> classItems1 = auxTable1.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in classItems1)
                {
                    svalue = ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGMinValue", "SupportNumber")).PropValue);
                    isEqual = String.Equals(SupportNumber, ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGMinValue", "SupportNumber")).PropValue), StringComparison.Ordinal);
                    if (isEqual == true)
                    {                            
                        MinimumWidth = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGMinValue", "MinimumWidth")).PropValue);
                        break;
                    }

                }

                if (CantileverLength > MaximumHeight)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinMaxHeight, "Height Exceeds Maximum Allowable Height. Resetting it to Maximum Height"), "", "TSupport", 346);
                    CantileverLength = MaximumHeight;
                    support.SetPropertyValue(Width, "IJUAhsOGWidth", "Width");
                }                

                if (Width > MaximumWidth)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMaxWidth, "Width is less than minimum Allowable Width. Resetting it to Maximum Width"), "", "TSupport", 346);
                    Width = MaximumWidth;
                    support.SetPropertyValue(Width, "IJUAhsOGWidth", "Width");
                }

                PartClass LoadTable = (PartClass)cataloghelper.GetPartClass("HS_OGAssy_LoadTable");
                classItems = LoadTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in classItems)
                {
                    svalue = ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGLoadTable", "SupportNumber")).PropValue);
                    isEqual = String.Equals(SupportNumber, ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGLoadTable", "SupportNumber")).PropValue), StringComparison.Ordinal);
                    if (isEqual == true)
                    {
                        CapacityDWPy = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "CapacityDWPy")).PropValue);
                        PipingPy = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "PipingPy")).PropValue);
                        PipingPx = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "PipingPx")).PropValue);
                        BlastPy = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "BlastPy")).PropValue);
                        BlastPx = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "BlastPx")).PropValue);
                        break;

                    }

                }

                //set part attributes

                componentDictionary[HOR_SECTION].SetPropertyValue(Width, "IJUAHgrOccLength", "Length");
                componentDictionary[VER_SECTION].SetPropertyValue(CantileverLength - horOffsetX, "IJUAHgrOccLength", "Length");
                if (SupportNumber == "MGT01" || SupportNumber == "MGT02" || SupportNumber == "MGT04")
                {
                    componentDictionary[VER_SECTION].SetPropertyValue(HorSection.depth, "IJUAHgrOccOverLength", "EndOverLength");
                }

                //Set Support Properties
                if (support.SupportsInterface("IJUAhsOGMaxValue"))
                    support.SetPropertyValue(MaximumHeight, "IJUAhsOGMaxValue", "MaxHeight");
                if (support.SupportsInterface("IJUAhsOGMaxValue"))
                    support.SetPropertyValue(MaximumWidth, "IJUAhsOGMaxValue", "MaxWidth");
                if (support.SupportsInterface("IJUAhsOGMinValue"))
                    support.SetPropertyValue(MinimumWidth, "IJUAhsOGMinValue", "MinWidth");
                if (support.SupportsInterface("IJUAhsOGCapacityDWPy"))
                    support.SetPropertyValue(CapacityDWPy, "IJUAhsOGCapacityDWPy", "CapacityDWPy");
                if (support.SupportsInterface("IJUAhsOGPipingPy"))
                    support.SetPropertyValue(PipingPy, "IJUAhsOGPipingPy", "PipingPy");
                if (support.SupportsInterface("IJUAhsOGPipingPx"))
                    support.SetPropertyValue(PipingPx, "IJUAhsOGPipingPx", "PipingPx");
                if (support.SupportsInterface("IJUAhsOGBlastPy"))
                    support.SetPropertyValue(BlastPy, "IJUAhsOGBlastPy", "BlastPy");
                if (support.SupportsInterface("IJUAhsOGBlastPx"))
                    support.SetPropertyValue(BlastPx, "IJUAhsOGBlastPx", "BlastPx");                    
            }            
                    
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion Get Aseembly Joints
        
        #region Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 8;
            }
        }
        #endregion
        
        #region Get Route Connections
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
                    int numberOfRoutes = SupportHelper.SupportedObjects.Count;
                    for (int routeIndex = 1; routeIndex <= numberOfRoutes; routeIndex++)
                    {
                        routeConnections.Add(new ConnectionInfo(HOR_SECTION, routeIndex)); // partindex, routeindex     
                    }

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
        #endregion
        
        #region Get Struct Connections
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
                    if (isweldedstarter)
                        structConnections.Add(new ConnectionInfo(WELDED_STARTER, 1)); // partindex, routeindex
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
        #endregion

        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomDesciption = "";
            try
            {
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject catalogpart = SupportOrComponent.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

                string supportNumber = (string)((PropertyValueString)catalogpart.GetPropertyValue("IJUAhsSupportNumber", "SupportNumber")).PropValue;
                bomDesciption = supportNumber;

                return bomDesciption;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion

    }
}
