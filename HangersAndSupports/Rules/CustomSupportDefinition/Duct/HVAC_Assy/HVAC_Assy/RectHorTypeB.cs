//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectHorTypeB.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectHorTypeB
//   Author       :Vijaya
//   Creation Date:10.Jun.2013
//   Description: CR-CP-224486 Convert HS_HVAC_Assy to C# .Net

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  10.Jun.2013     Vijaya   CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//  27.Apr.2015     PVK      TR-CP-253033 Elevation CP not shown by default
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
    public class RectHorTypeB : CustomSupportDefinition
    {
        //Constants
        private const string HORSECTION1 = "HORSECTION1";
        private const string HORSECTION2 = "HORSECTION2";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";
        string[] boltPartKeys = new string[8];
        string LEFTPLATE = "LEFTPLATE";
        string RIGHTPLATE = "RIGHTPLATE";


        double plateThickness;
        string sectionSize, boltSize, boltBOMDesc, gasketBOMDesc, sizeOfSection = string.Empty, sizeOfBolt;
        int boltBegin=0, boltEnd=0;
        bool includeLeftPlate, includeRightPlate, showBolts, boltsizeFromRule, sectionFromRule, value;
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

                    includeLeftPlate = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyLPlate", "IncludeLeftPlate")).PropValue;
                    includeRightPlate = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyRPlate", "IncludeRightPlate")).PropValue;
                    plateThickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyPlateThk", "PlateThick")).PropValue;
                    showBolts = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "ShowBolts")).PropValue;

                    boltsizeFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSizeFromRule")).PropValue;
                    PropertyValueCodelist boltCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSize");

                    boltSize = boltCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltCodeList.PropValue).DisplayName;
                    boltBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACBoltsDesc", "BoltDesc")).PropValue;
                    gasketBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACGasketDesc", "GasketDesc")).PropValue;

                    //Get the Section Size from Rule
                    if (sectionFromRule)
                        value = GenericHelper.GetDataByRule("HVACHgrSectionSize", support, out sizeOfSection);                    
                    else if (!sectionFromRule)
                    {
                        sizeOfSection = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSize).DisplayName;

                        if (sizeOfSection.ToUpper().Equals("NONE") || sizeOfSection.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RectHorTypeB.cs", 88);
                            return null;
                        }
                    }

                    //Get the Bolt Size from Rule
                    if (boltsizeFromRule)
                        value = GenericHelper.GetDataByRule("HVACHgrBoltSize", support, out sizeOfBolt);                   
                   else if (!boltsizeFromRule)
                    {
                        sizeOfBolt = boltCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltSize).DisplayName;

                        if (sizeOfBolt.ToUpper().Equals("NONE") || sizeOfBolt.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Bolt size is not available.", "", "RectHorTypeB.cs", 102);
                            return null;
                        }
                    }
                    //Create the list of parts
                    parts.Add(new PartInfo(HORSECTION1, sizeOfSection));
                    parts.Add(new PartInfo(HORSECTION2, sizeOfSection));
                    parts.Add(new PartInfo(VERTSECTION1, sizeOfSection));
                    parts.Add(new PartInfo(VERTSECTION2, sizeOfSection));
                   
                    if (includeLeftPlate)
                        parts.Add(new PartInfo(LEFTPLATE, "Util_Plate_Metric_1"));
                    if (includeRightPlate)
                        parts.Add(new PartInfo(RIGHTPLATE, "Util_Plate_Metric_1"));
                    if (showBolts)
                    {
                       int  partIndex = 0;
                        boltBegin = parts.Count + 1;
                        boltEnd = boltBegin + 3;
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
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                   boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
               else
                   boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height;

                string supportingType, gasketNote = string.Empty, boltsDesc = string.Empty;
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
                //Auto Dimensioning of Supports
                if (includeRightPlate)
                { 
                    Note noteDimenssion1 = CreateNote("Dim 1", componentDictionary[RIGHTPLATE], "TopStructure");
                    noteDimenssion1.SetPropertyValue("", "IJGeneralNote", "Text");
                    CodelistItem fabrication1 = noteDimenssion1.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    noteDimenssion1.SetPropertyValue(fabrication1, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                    noteDimenssion1.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                }
                Note noteDimenssion2 = CreateNote("Dim 2", componentDictionary[HORSECTION2], "BeginCap");
                noteDimenssion2.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem fabrication2 = noteDimenssion2.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                noteDimenssion2.SetPropertyValue(fabrication2, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                noteDimenssion2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");                
              
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
                    Note gasketnote = CreateNote("Dim 2", componentDictionary[HORSECTION2], "BeginCap");
                    noteDimenssion2.SetPropertyValue(gasketNote, "IJGeneralNote", "Text");
                    CodelistItem fabrication = noteDimenssion2.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    noteDimenssion2.SetPropertyValue(fabrication, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                    noteDimenssion2.SetPropertyValue(true,"IJGeneralNote","Dimensioned");
                }
                else
                    DeleteNoteIfExists("Dim 2");
                
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
                string leftStructPort = structPort[0],rightStructPort = structPort[1];
               
                //Set Values of Part Occurance Attributes
                PropertyValueCodelist beginMiterCodelist = (PropertyValueCodelist)componentDictionary[HORSECTION1].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMiterCodelist.PropValue == -1)
                    beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist endMiterCodelist = (PropertyValueCodelist)componentDictionary[HORSECTION1].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMiterCodelist.PropValue == -1)
                    endMiterCodelist.PropValue = 1;

                // Set Values of Part Occurance Attributes
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

                // Miscallaneous
                //Get the route structure distance
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                double routeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                string structPort1, structPort2, horSecPort = string.Empty, routePort = string.Empty;
                double overhangLeft = 0, overhangRight = 0, horSec1RouteAxOffset = 0, horSec1RouteOrgOffset = 0;
                if (Configuration == 1)
                {
                    structPort1 = leftStructPort;
                    structPort2 = rightStructPort;
                }
                else
                {
                    structPort1 = rightStructPort;
                    structPort2 = leftStructPort;
                }

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    double routeStructDist = 0;
                    if ((Math.Abs(routeAngle) < (0 + 0.001) && Math.Abs(routeAngle) > (0 - 0.001)) || (Math.Abs(routeAngle) < (Math.PI + 0.001) && Math.Abs(routeAngle) > (Math.PI - 0.001)))
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is valid only when the Route is horizontal. For vertical route, use Vertical Duct Support.", "", "RectHorTypeB.cs", 330);
                        return;
                    }
                    else
                    {
                        if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && Math.Abs(routeStructAngle) > (Math.PI - 0.001)))
                            routeStructDist = RefPortHelper.DistanceBetweenPorts("BBR_High", "Structure", PortDistanceType.Horizontal);
                        else
                            routeStructDist = RefPortHelper.DistanceBetweenPorts("BBR_High", "Structure", PortDistanceType.Vertical);
                    }

                    overhangLeft = 0;
                    overhangRight = 0;
                    componentDictionary[VERTSECTION1].SetPropertyValue(routeStructDist + boundingBoxHeight + sectionDepth, "IJUAHgrOccLength", "Length");
                    componentDictionary[VERTSECTION2].SetPropertyValue(routeStructDist + boundingBoxHeight + sectionDepth, "IJUAHgrOccLength", "Length");
                    componentDictionary[HORSECTION1].SetPropertyValue(boundingBoxWidth + 2 * sectionWidth, "IJUAHgrOccLength", "Length");
                    componentDictionary[HORSECTION2].SetPropertyValue(boundingBoxWidth + 2 * sectionWidth, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    Double routeStruct1VertDist = 0, routeStruct2VertDist = 0, routeLowStruct1HorDist = 0, routeLowStruct2HorDist = 0, routeHighStruct1HorDist = 0, routeHighStruct2HorDist = 0, horSec1Length = 0, horSec2Length = 0, distBetStruct = 0;

                    if ((Math.Abs(routeAngle) < (0 + 0.001) && Math.Abs(routeAngle) > (0 - 0.001)) || (Math.Abs(routeAngle) < (Math.PI + 0.001) && Math.Abs(routeAngle) > (Math.PI - 0.001)))
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is valid only when the Route is horizontal. For vertical route, use Vertical Duct Support.", "", "RectHorTypeB.cs", 354);
                        return;
                    }
                    else
                        if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && Math.Abs(routeStructAngle) > (Math.PI - 0.001)))
                        {
                            routeStruct1VertDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", structPort1, PortDistanceType.Horizontal);
                            routeStruct2VertDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", structPort2, PortDistanceType.Horizontal);
                            routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Vertical);
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Vertical);
                            routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Vertical);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Vertical);
                            distBetStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Vertical);
                        }
                        else
                        {
                            routeStruct1VertDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", structPort1, PortDistanceType.Vertical);
                            routeStruct2VertDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", structPort2, PortDistanceType.Vertical);
                            routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Horizontal);
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Horizontal);
                            routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Horizontal);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Horizontal);
                            distBetStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
                        }

                    if (supportingType.ToUpper().Equals("SLAB"))
                    {
                        horSec1Length = boundingBoxWidth + 2 * sectionDepth;
                        overhangLeft = 0;
                        overhangRight = 0;
                    }
                    else if (supportingType.ToUpper().Equals("STEEL"))
                    {
                        horSec1Length = distBetStruct + sectionWidth;
                        if (routeLowStruct2HorDist < routeHighStruct2HorDist)
                            overhangLeft = routeLowStruct2HorDist - sectionWidth / 2;
                        else
                            overhangLeft = routeHighStruct2HorDist - sectionWidth / 2;
                        if (routeLowStruct1HorDist < routeHighStruct1HorDist)
                            overhangRight = routeLowStruct1HorDist - sectionWidth / 2;
                        else
                            overhangRight = routeHighStruct1HorDist - sectionWidth / 2;
                    }
                    else if (supportingType.Equals("STEEL-SLAB") || supportingType.Equals("SLAB-STEEL"))
                    {
                        if (routeLowStruct2HorDist < routeHighStruct2HorDist)
                            overhangLeft = routeLowStruct2HorDist - sectionWidth / 2;
                        else
                            overhangLeft = routeHighStruct2HorDist - sectionWidth / 2;
                        if (routeLowStruct1HorDist < routeHighStruct1HorDist)
                            overhangRight = routeLowStruct1HorDist - sectionWidth / 2;
                        else
                            overhangRight = routeHighStruct1HorDist - sectionWidth / 2;

                        horSec1Length = boundingBoxWidth + sectionWidth + overhangLeft + overhangRight;
                    }
                    horSec2Length = horSec1Length;
                    componentDictionary[HORSECTION1].SetPropertyValue(horSec1Length, "IJUAHgrOccLength", "Length");
                    componentDictionary[HORSECTION2].SetPropertyValue(horSec2Length, "IJUAHgrOccLength", "Length");
                    componentDictionary[VERTSECTION1].SetPropertyValue(routeStruct1VertDist + sectionDepth, "IJUAHgrOccLength", "Length");
                    componentDictionary[VERTSECTION2].SetPropertyValue(routeStruct2VertDist + sectionDepth, "IJUAHgrOccLength", "Length");
                }

                Plane horSec1RoutePlaneA = new Plane(), horSec1RoutePlaneB = new Plane();
                Axis horSec1RouteAxisA = new Axis(), horSec1RouteAxisB = new Axis(); 

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    routePort = "BBSR_Low";
                else
                    routePort = "BBR_Low";
                if (Configuration == 1)
                { 
                    horSec1RoutePlaneA = Plane.ZX;
                    horSec1RoutePlaneB = Plane.NegativeXY;
                    horSec1RouteAxisA = Axis.X;
                    horSec1RouteAxisB = Axis.X;
                    horSecPort = "BeginCap";
                    horSec1RouteAxOffset = sectionWidth + overhangLeft;
                    horSec1RouteOrgOffset = -sectionDepth / 2;
                }
                else if (Configuration == 2)
                {
                    horSec1RoutePlaneA = Plane.ZX;
                    horSec1RoutePlaneB = Plane.NegativeXY;
                    horSec1RouteAxisA = Axis.X;
                    horSec1RouteAxisB = Axis.NegativeX;
                    horSecPort = "EndCap";
                    horSec1RouteAxOffset = -sectionWidth - overhangLeft;
                    horSec1RouteOrgOffset = -sectionDepth / 2;
                }

                //Create Joints
                //Joints when plates are included
                if (!includeLeftPlate && includeRightPlate)
                {
                    componentDictionary[VERTSECTION1].SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[RIGHTPLATE].SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[RIGHTPLATE].SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[RIGHTPLATE].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    //Add Joint Between the Right plate and the Vertical Section 1
                    JointHelper.CreateRigidJoint(RIGHTPLATE, "BotStructure", VERTSECTION1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);
                }
                else if (includeLeftPlate && !includeRightPlate)
                {
                    componentDictionary[VERTSECTION2].SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                    componentDictionary[LEFTPLATE].SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[LEFTPLATE].SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[LEFTPLATE].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");
                    //Add Joint Between the Left plate and the Vertical Section 2
                    JointHelper.CreateRigidJoint(LEFTPLATE, "TopStructure", VERTSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);
                }
                else if (includeLeftPlate && includeRightPlate)
                {
                    componentDictionary[VERTSECTION1].SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[RIGHTPLATE].SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[RIGHTPLATE].SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[RIGHTPLATE].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    componentDictionary[VERTSECTION2].SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                    componentDictionary[LEFTPLATE].SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[LEFTPLATE].SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[LEFTPLATE].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");
                    //Add Joint Between the Right plate and the Vertical Section 1
                    JointHelper.CreateRigidJoint(LEFTPLATE, "BotStructure", VERTSECTION1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);
                    //Add Joint Between the Left plate and the Vertical Section 2
                    JointHelper.CreateRigidJoint(RIGHTPLATE, "TopStructure", VERTSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);
                }                

                //Add bolts
                if (showBolts)
                {
                    JointHelper.CreateRigidJoint(boltPartKeys[0], "StartOther", HORSECTION1, "EndCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, -sectionWidth / 2);
                    JointHelper.CreateRigidJoint(boltPartKeys[1], "StartOther", HORSECTION1, "BeginCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, sectionWidth / 2);
                    JointHelper.CreateRigidJoint(boltPartKeys[2], "StartOther", HORSECTION2, "EndCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, -sectionWidth / 2);
                    JointHelper.CreateRigidJoint(boltPartKeys[3], "StartOther", HORSECTION2, "BeginCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, sectionWidth / 2);
                }
                
                //Add joint between Horizontal Section 1 and Route
                JointHelper.CreateRigidJoint(HORSECTION1, horSecPort, "-1", routePort, horSec1RoutePlaneA, horSec1RoutePlaneB, horSec1RouteAxisA, horSec1RouteAxisB, 0, horSec1RouteAxOffset, horSec1RouteOrgOffset);

                //Add joint between Horizontal Section 1 and Horizontal Section 2
                JointHelper.CreateRigidJoint(HORSECTION1, "Neutral", HORSECTION2, "Neutral", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -boundingBoxHeight - sectionWidth, 0);

                //Add joint between Horizontal Section 1 and Vertical Section 1
                 JointHelper.CreateRigidJoint(VERTSECTION1, "BeginCap", HORSECTION1, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX, sectionDepth, sectionDepth, 0);

                //Add joint between Horizontal Section 1 and Vertical Section 2
                JointHelper.CreateRigidJoint(VERTSECTION2, "EndCap", HORSECTION1, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, sectionDepth, -sectionDepth, 0);
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

