//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectHorTypeC.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectHorTypeC
//   Author       :Vijaya
//   Creation Date:11.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  11.Jun.2013     Vijaya   CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//  27.Apr.2015     PVK      TR-CP-253033 Elevation CP not shown by default
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
    public class RectHorTypeC : CustomSupportDefinition
    {
        //Constants
        private const string HORSECTION1 = "HORSECTION1";
        private const string HORSECTION2 = "HORSECTION2";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";
        string BRACE = "BRACE";
        string CONNECTION = "CONNECTION";
        string BRACEPLATE = "BRACEPLATE";
        string CANTIPLATE = "CANTIPLATE";
        string[] boltPartKeys = new string[8];

        double plateThickness, overhang, braceAngle;
        string sectionSize, boltSize, boltBOMDesc, gasketBOMDesc, sizeOfSection = string.Empty, sizeOfBolt, angleOption;
        int boltBegin, boltEnd;
        bool sectionFromRule, boltsizeFromRule, showBrace, includeCantiPlate, includeBracePlate, showBolts, value;
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
                    sectionFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionFromRule")).PropValue;

                    PropertyValueCodelist sectionCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionSize");
                    sectionSize = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionCodeList.PropValue).DisplayName;

                    includeCantiPlate = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyCantiPlate", "IncludeCantiPlate")).PropValue;
                    includeBracePlate = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBrPlate", "IncludeBrPlate")).PropValue;
                    plateThickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyPlateThk", "PlateThick")).PropValue;
                    overhang = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyOH", "Overhang")).PropValue;
                    showBrace = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBrace", "ShowBrace")).PropValue;
                    braceAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyBrace", "BraceAngle")).PropValue;
                    showBolts = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "ShowBolts")).PropValue;
                    boltsizeFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSizeFromRule")).PropValue;

                    PropertyValueCodelist boltCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSize");
                    boltSize = boltCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltCodeList.PropValue).DisplayName;

                    boltBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACBoltsDesc", "BoltDesc")).PropValue;
                    gasketBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACGasketDesc", "GasketDesc")).PropValue;
                    PropertyValueCodelist angleCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyAngleOpt", "AngleOpt");
                    angleOption = angleCodeList.PropValue.ToString();

                    BusinessObject supportObject = support;
                    //Get the Section Size from Rule
                    if (sectionFromRule)
                        value = GenericHelper.GetDataByRule("HVACHgrSectionSize", supportObject, out sizeOfSection);
                    else if (!sectionFromRule)
                    {
                        sizeOfSection = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSize).DisplayName;

                        if (sizeOfSection.ToUpper().Equals("NONE") || sizeOfSection.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RectHorTypeC.cs", 95);
                            return null;
                        }
                    }

                    //Get the Bolt Size from Rule
                    if (boltsizeFromRule)
                        value = GenericHelper.GetDataByRule("HVACHgrBoltSize", supportObject, out sizeOfBolt);
                    else if (!boltsizeFromRule)
                    {
                        sizeOfBolt = boltCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltSize).DisplayName;

                        if (sizeOfBolt.ToUpper().Equals("NONE") || sizeOfBolt.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Bolt size is not available.", "", "RectHorTypeC.cs", 109);
                            return null;
                        }
                    }
                    //Create the list of parts
                    parts.Add(new PartInfo(HORSECTION1, sizeOfSection));
                    parts.Add(new PartInfo(HORSECTION2, sizeOfSection));
                    parts.Add(new PartInfo(VERTSECTION1, sizeOfSection));
                    parts.Add(new PartInfo(VERTSECTION2, sizeOfSection));

                    if (showBrace)
                    {
                        parts.Add(new PartInfo(BRACE, "HgrHVACGenericBrace_L1_1"));
                        parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                        if (includeBracePlate)
                            parts.Add(new PartInfo(BRACEPLATE, "Util_Plate_Metric_1"));
                    }
                    if (includeCantiPlate)
                        parts.Add(new PartInfo(CANTIPLATE, "Util_Plate_Metric_1"));

                    if (showBolts)
                    {
                        int partIndex = 0;
                        boltBegin = parts.Count + 1;
                        boltEnd = boltBegin + 1;
                        for (int index = boltBegin; index <= boltEnd; index++)
                        {
                            boltPartKeys[partIndex] = "Bolt" + (partIndex + 1).ToString();
                            parts.Add(new PartInfo(boltPartKeys[partIndex], "Util_Fixed_Cyl_Metric_1"));
                            partIndex++;
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
                //Auto Dimensioning of Supports
                if (includeCantiPlate)
                {
                    Note noteDimenssion1 = CreateNote("Dim 1", componentDictionary[CANTIPLATE], "TopStructure");
                    noteDimenssion1.SetPropertyValue("", "IJGeneralNote", "Text");
                    CodelistItem fabrication1 = noteDimenssion1.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    noteDimenssion1.SetPropertyValue(fabrication1, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                    noteDimenssion1.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                }
                Note noteDimenssion2 = CreateNote("Dim 2", componentDictionary[HORSECTION1], "EndCap");
                noteDimenssion2.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem fabrication2 = noteDimenssion2.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                noteDimenssion2.SetPropertyValue(fabrication2, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                noteDimenssion2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                if (showBrace)
                {
                    Note noteDimenssion3 = CreateNote("Dim 3", componentDictionary[BRACE], "BeginCap");
                    noteDimenssion2.SetPropertyValue("", "IJGeneralNote", "Text");
                    CodelistItem fabrication3 = noteDimenssion3.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    noteDimenssion2.SetPropertyValue(fabrication3, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                    noteDimenssion2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                }

                String gasketNote = string.Empty, boltsDesc = string.Empty;
                if (string.IsNullOrEmpty(gasketBOMDesc))
                    gasketNote = "To line with 3mm thick Gasket all around for all contacted surface between Duct & Support";
                else
                    gasketNote = gasketBOMDesc;
                //Gasket Note
                bool excludeNote;
                if (support.SupportsInterface("IJUAhsExcludeNotes"))
                    excludeNote = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;
                else
                    excludeNote = false;
                if (!excludeNote)
                {
                    Note gasketnote = CreateNote("Note 1");
                    noteDimenssion2.SetPropertyValue(gasketNote, "IJGeneralNote", "Text");
                    CodelistItem fabrication = noteDimenssion2.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    noteDimenssion2.SetPropertyValue(fabrication, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                    noteDimenssion2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                }
                else
                    DeleteNoteIfExists("Note 1");
                //Get Section Structure dimensions
                BusinessObject horizontalSectionPart = componentDictionary[HORSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double sectionWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double sectionDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                double sectionThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                //Get the Current Location in the Route Connection Cycle
                //Determine whether Route is horizontal
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);

                if ((Math.Abs(routeStructAngle) < (Math.PI / 2 + 0.001) && Math.Abs(routeStructAngle) > (Math.PI / 2 - 0.001)))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is valid only when the Route is horizontal and the supporting Strcuture is besides the Route.", "", "RectHorTypeC.cs", 240);
                    return;
                }

                //Set Values of Part Occurance Attributes
                PropertyValueCodelist beginMiterCodelist = (PropertyValueCodelist)componentDictionary[HORSECTION1].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMiterCodelist.PropValue == -1)
                    beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist endMiterCodelist = (PropertyValueCodelist)componentDictionary[HORSECTION1].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMiterCodelist.PropValue == -1)
                    endMiterCodelist.PropValue = 1;

                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HORSECTION1].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[HORSECTION1].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[HORSECTION1].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");
                componentDictionary[HORSECTION1].SetPropertyValue(0.001, "IJUAHgrOccLength", "Length");

                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HORSECTION2].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[HORSECTION2].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[HORSECTION2].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");
                componentDictionary[HORSECTION2].SetPropertyValue(0.001, "IJUAHgrOccLength", "Length");

                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERTSECTION1].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERTSECTION1].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.001, "IJUAHgrOccLength", "Length");

                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERTSECTION2].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERTSECTION2].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.001, "IJUAHgrOccLength", "Length");

                if (string.IsNullOrEmpty(boltBOMDesc))
                    boltsDesc = sizeOfBolt + " " + "Bolts";
                else
                    boltsDesc = boltBOMDesc;

                if (showBolts)
                {
                    for (int boltIndex = boltBegin; boltIndex <= boltEnd; boltIndex++)
                    {
                        componentDictionary[boltPartKeys[boltIndex - boltBegin]].SetPropertyValue(sectionDepth / 2, "IJOAHgrUtilMetricL", "L");
                        componentDictionary[boltPartKeys[boltIndex - boltBegin]].SetPropertyValue(sectionWidth / 4, "IJOAHgrUtilMetricRadius", "Radius");
                        componentDictionary[boltPartKeys[boltIndex - boltBegin]].SetPropertyValue(boltsDesc, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                    }
                }
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                //Set SectionSize attribute value on the support
                if (sectionFromRule)
                {
                    CodelistItem codeList = metadataManager.GetCodelistInfo("HVACAssyHgrSecSize", "UDP").GetCodelistItem(sizeOfSection);
                    support.SetPropertyValue(codeList.Value, "IJOAHgrHVACAssySection", "SectionSize");
                }
                //Set BoltSize attribute value on the support
                if (boltsizeFromRule)
                {
                    CodelistItem codeList = metadataManager.GetCodelistInfo("HVACAssyHgrBoltSize", "UDP").GetCodelistItem(sizeOfBolt);
                    support.SetPropertyValue(codeList.Value, "IJOAHgrHVACAssyBolts", "BoltSize");
                }
                Double horSec1Length, horSec2Length, vertSec1Length, vertSec2Length, routeHighStructDist, braceAngleValue = 0, braceLength1 = 0;

                horSec1Length = boundingBoxHeight + 2 * sectionWidth;
                vertSec1Length = boundingBoxWidth + sectionDepth;
                horSec2Length = horSec1Length;
                vertSec2Length = vertSec1Length;
                componentDictionary[HORSECTION1].SetPropertyValue(horSec1Length, "IJUAHgrOccLength", "Length");
                componentDictionary[HORSECTION2].SetPropertyValue(horSec2Length, "IJUAHgrOccLength", "Length");
                componentDictionary[VERTSECTION1].SetPropertyValue(vertSec1Length, "IJUAHgrOccLength", "Length");
                componentDictionary[VERTSECTION2].SetPropertyValue(vertSec2Length, "IJUAHgrOccLength", "Length");

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    routeHighStructDist = RefPortHelper.DistanceBetweenPorts("Structure", "BBSR_High", PortDistanceType.Horizontal);
                else
                    routeHighStructDist = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_High", PortDistanceType.Horizontal);

                componentDictionary[HORSECTION2].SetPropertyValue(overhang, "IJUAHgrOccOverLength", "EndOverLength");

                if (includeCantiPlate)
                    componentDictionary[HORSECTION2].SetPropertyValue(routeHighStructDist - sectionWidth - plateThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                else
                    componentDictionary[HORSECTION2].SetPropertyValue(routeHighStructDist - sectionWidth, "IJUAHgrOccOverLength", "BeginOverLength");

                Matrix4X4 routePortOrientation = new Matrix4X4(), structPortOrientation = new Matrix4X4();

                if (angleOption.Equals("1"))
                {
                    if (SupportHelper.SupportingObjects.Count == 1)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Second structural member is not available. Change the Angle option to User Defined", "", "RectHorTypeC.cs", 336);
                        return;
                    }
                    else if (SupportHelper.SupportingObjects.Count > 1)
                    {
                        if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                        {
                            //Get the Ref Port Origin 
                            routePortOrientation = RefPortHelper.PortLCS("BBR_High");
                            structPortOrientation = RefPortHelper.PortLCS("Struct_2");
                        }
                        Vector tempVector = new Vector(), tempVector1 = new Vector(), normalVector = new Vector();
                        tempVector1 = new Vector(routePortOrientation.ZAxis.X, routePortOrientation.ZAxis.Y, routePortOrientation.ZAxis.Z);
                        tempVector1.Length = -(boundingBoxHeight + sectionWidth + overhang);
                        Position point = routePortOrientation.Origin.Offset(tempVector1);
                        tempVector = new Vector(point.X - structPortOrientation.Origin.X, point.Y - structPortOrientation.Origin.Y, point.Z - structPortOrientation.Origin.Z);
                        normalVector = tempVector.Cross(tempVector1);
                        double angle = tempVector.Angle(tempVector1, normalVector);
                        braceLength1 = tempVector.Length;
                        braceAngleValue = angle;
                        support.SetPropertyValue(braceAngleValue, "IJOAHgrHVACAssyBrace", "BraceAngle");
                    }
                }
                else if (angleOption.Equals("2"))
                {
                    braceAngleValue = braceAngle;
                    support.SetPropertyValue(braceAngleValue, "IJOAHgrHVACAssyBrace", "BraceAngle");
                }
                double brPadOffset = 0;
                if (showBrace)
                {
                    Double horSecLength = 0, braceLength = 0;
                    if (angleOption.Equals("2")) //User Defined
                    {
                        horSecLength = horSec2Length + overhang + routeHighStructDist - sectionWidth;
                        braceLength = (horSecLength - sectionThickness) / Math.Cos(braceAngleValue);
                    }
                    else if (angleOption.Equals("1"))    //Adjust to Supporting Steel
                        braceLength = braceLength1;

                    brPadOffset = (braceLength) * Math.Sin(braceAngleValue);                   
                  
                    string braceBOM = "Generic Brace " + sizeOfSection + " Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(support.GetPropertyValue("IJOAHgrHVACAssyOH", "Overhang").PropertyInfo.UOMType, braceLength, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.DISTANCE_INCH_SYMBOL);

                    componentDictionary[BRACE].SetPropertyValue(braceLength, "IJOAHgrHVACGenBrace", "L");
                    componentDictionary[BRACE].SetPropertyValue(braceAngleValue, "IJOAHgrHVACGenBrace", "Angle");
                    PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[BRACE].GetPropertyValue("IJOAHgrHVACGenBrace", "BraceOrient");
                    braceCodelist.PropValue = 2;
                    componentDictionary[BRACE].SetPropertyValue(braceCodelist.PropValue, "IJOAHgrHVACGenBrace", "BraceOrient");
                    componentDictionary[BRACE].SetPropertyValue(sectionWidth, "IJOAHgrHVACGenBrace", "W");
                    componentDictionary[BRACE].SetPropertyValue(sectionDepth, "IJOAHgrHVACGenBrace", "Depth");
                    componentDictionary[BRACE].SetPropertyValue(sectionThickness, "IJOAHgrHVACGenBrace", "FlangeT");
                    componentDictionary[BRACE].SetPropertyValue(braceAngleValue - Math.PI / 2, "IJOAHgrHVACGenBrace", "CutBackAngle1");
                    componentDictionary[BRACE].SetPropertyValue(Math.PI / 2 - braceAngleValue, "IJOAHgrHVACGenBrace", "CutBackAngle2");
                    componentDictionary[BRACE].SetPropertyValue(braceBOM, "IJOAHgrHVACGenBrace", "InputBomDesc");
                }

                Plane vertSec1RoutePlaneA = new Plane(), vertSec1RoutePlaneB = new Plane(), vertSec2RoutePlaneA = new Plane(), vertSec2RoutePlaneB = new Plane();
                Axis vertSec1RouteAxisA = new Axis(), vertSec1RouteAxisB = new Axis(), vertSec2RouteAxisA = new Axis(), vertSec2RouteAxisB = new Axis();
                String routePort1 = string.Empty, routePort2 = string.Empty, horSec1Port = string.Empty, vertSecPort = string.Empty;
                Double vertSec1RoutePlOffset = 0, vertSec2RoutePlOffset = 0;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    routePort1 = "BBSR_Low";
                    routePort2 = "BBSR_High";
                }
                else
                {
                    routePort1 = "BBR_Low";
                    routePort2 = "BBR_High";
                }
                if (Configuration == 1)
                {
                    vertSec1RoutePlaneA = Plane.XY;
                    vertSec1RoutePlaneB = Plane.NegativeZX;
                    vertSec1RouteAxisA = Axis.X;
                    vertSec1RouteAxisB = Axis.NegativeX;

                    vertSec2RoutePlaneA = Plane.XY;
                    vertSec2RoutePlaneB = Plane.ZX;
                    vertSec2RouteAxisA = Axis.X;
                    vertSec2RouteAxisB = Axis.NegativeX;


                    vertSec1RoutePlOffset = 0;
                    vertSec2RoutePlOffset = -sectionDepth;
                    horSec1Port = "BeginCap";
                    vertSecPort = "EndCap";
                }
                else if (Configuration == 2)
                {
                    vertSec1RoutePlaneA = Plane.XY;
                    vertSec1RoutePlaneB = Plane.ZX;
                    vertSec1RouteAxisA = Axis.X;
                    vertSec1RouteAxisB = Axis.X;

                    vertSec2RoutePlaneA = Plane.XY;
                    vertSec2RoutePlaneB = Plane.NegativeZX;
                    vertSec2RouteAxisA = Axis.X;
                    vertSec2RouteAxisB = Axis.X;


                    vertSec1RoutePlOffset = sectionDepth;
                    vertSec2RoutePlOffset = 0;
                    horSec1Port = "BeginCap";
                    vertSecPort = "BeginCap";
                }

                //Create Joints
                if (showBrace)
                {
                    //Add a Joint between horizontal section and the point that we want the angle to rotate around.
                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", HORSECTION2, "EndCap", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0, 0);

                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", BRACE, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.X, 0, overhang, 0);
                    if (includeBracePlate)
                    {
                        componentDictionary[BRACEPLATE].SetPropertyValue(3 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                        componentDictionary[BRACEPLATE].SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                        componentDictionary[BRACEPLATE].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");
                        if (angleOption.Equals("2"))
                            //Add Joint Between the Plate and the Horizontal Section              
                            JointHelper.CreateRigidJoint(BRACEPLATE, "TopStructure", HORSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, routeHighStructDist - sectionWidth, -brPadOffset - sectionWidth / 2, sectionDepth / 2);

                        else if (angleOption.Equals("1"))
                            //Add Joint Between the Plate and the Horizontal Section
                            JointHelper.CreateRigidJoint(BRACEPLATE, "TopStructure", "-1", "Struct_2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, sectionWidth / 2, -sectionWidth - sectionWidth / 2);
                    }
                }
                if (includeCantiPlate)
                {
                    componentDictionary[CANTIPLATE].SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[CANTIPLATE].SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[CANTIPLATE].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");
                    //Add Joint Between the Left plate and the Vertical Section 2
                    JointHelper.CreateRigidJoint(CANTIPLATE, "TopStructure", HORSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, routeHighStructDist - sectionWidth, -sectionWidth / 2, sectionDepth / 2);
                }
                //Add bolts
                if (showBolts)
                {
                    JointHelper.CreateRigidJoint(boltPartKeys[0], "StartOther", HORSECTION1, "EndCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, -sectionWidth / 2);

                    JointHelper.CreateRigidJoint(boltPartKeys[1], "StartOther", HORSECTION1, "BeginCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, sectionWidth / 2);
                }

                //Add joint between Vertical Section 1 and Route
                JointHelper.CreateRigidJoint(VERTSECTION1, vertSecPort, "-1", routePort1, vertSec1RoutePlaneA, vertSec1RoutePlaneB, vertSec1RouteAxisA, vertSec1RouteAxisB, vertSec1RoutePlOffset, 0, 0);

                //Add joint between  Vertical Section 2 and Route
                JointHelper.CreateRigidJoint(VERTSECTION2, vertSecPort, "-1", routePort2, vertSec2RoutePlaneA, vertSec2RoutePlaneB, vertSec2RouteAxisA, vertSec2RouteAxisB, vertSec2RoutePlOffset, 0, 0);

                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(VERTSECTION1, "BeginCap", HORSECTION1, horSec1Port, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX, sectionDepth, sectionDepth, 0);

                //Add joint between Horizontal Section 1 and Horizontal Section 2
                JointHelper.CreateRigidJoint(HORSECTION1, "Neutral", HORSECTION2, "Neutral", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -boundingBoxWidth - sectionWidth, -sectionWidth);
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

                    structConnections.Add(new ConnectionInfo(HORSECTION1, 1)); // partindex, routeindex
                    int structCount = SupportHelper.SupportingObjects.Count;
                    if (showBrace)
                    {
                        //Create an ARRAY to hold Structural Connection information for the Beam Clamp          
                        if (structCount > 1)
                            structConnections.Add(new ConnectionInfo(BRACE, 2)); // partindex, routeindex            
                    }
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

