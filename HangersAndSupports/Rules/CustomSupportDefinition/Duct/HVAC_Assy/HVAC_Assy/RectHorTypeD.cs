//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectHorTypeD.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectHorTypeD
//   Author       :Vijaya
//   Creation Date:12.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  12.Jun.2013     Vijaya   CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
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
    public class RectHorTypeD : CustomSupportDefinition
    {
        //Constants
        private const string FLATBAR = "FLATBAR";
        private const string HORSECTION = "HORSECTION";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";
        string[] boltPartKeys = new string[8];
        string LEFTPLATE = "LEFTPLATE";
        string RIGHTPLATE = "RIGHTPLATE";


        double  plateThickness, flatBarThickness, flatBarWidth;
        string sectionSize, boltBOMDesc, gasketBOMDesc, sizeOfSection=string.Empty, sizeOfBolt, flatBarDim, boltSize;
        int  boltBegin, boltEnd;
        bool sectionFromRule, boltsizeFromRule, showBolts, includeLeftPlate, includeRightPlate, value;
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
                    flatBarDim = (string)((PropertyValueString)support.GetPropertyValue("IJOAHgrHVACAssyFlatBarDim", "FlatBarDim")).PropValue;
                    showBolts = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "ShowBolts")).PropValue;
                    boltsizeFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSizeFromRule")).PropValue;

                    PropertyValueCodelist boltCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSize");
                    boltSize = boltCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltCodeList.PropValue).DisplayName;  
                 
                    boltBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACBoltsDesc", "BoltDesc")).PropValue;
                    gasketBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACGasketDesc", "GasketDesc")).PropValue;

                    string[] position = flatBarDim.Split('x');
                    flatBarWidth = double.Parse(position[0]) / 1000;
                    flatBarThickness = double.Parse(position[1]) / 1000;

                    BusinessObject supportObject = support;
                    //Get the Section Size from Rule
                    if (sectionFromRule)
                             value = GenericHelper.GetDataByRule("HVACHgrSectionSize", supportObject, out sizeOfSection);                      
                     else if (!sectionFromRule)
                    {
                        sizeOfSection = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSize).DisplayName;

                        if (sizeOfSection.ToUpper().Equals("NONE") || sizeOfSection.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RectHorTypeD.cs", 94);
                            return null;
                        }   
                    }
                    //Get the Bolt Size from Rule
                    if (boltsizeFromRule)
                        value = GenericHelper.GetDataByRule("HVACHgrBoltSize", supportObject, out sizeOfBolt);                       
                    else if (!boltsizeFromRule)
                    {
                        string boltSizeName = boltCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltSize).DisplayName;

                        if (boltSizeName.ToUpper().Equals("NONE") || boltSizeName.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Bolt size is not available.", "", "RectHorTypeD.cs", 107);
                            return null;
                        }
                    }
                    //Create the list of parts
                    parts.Add(new PartInfo(FLATBAR, "Util_Plate_Metric_1"));
                    parts.Add(new PartInfo(HORSECTION, sizeOfSection));
                    parts.Add(new PartInfo(VERTSECTION1, "HgrHVACGenericBrace_L2_1"));
                    parts.Add(new PartInfo(VERTSECTION2, "HgrHVACGenericBrace_L2_1"));
                   
                    if (includeLeftPlate)
                     parts.Add(new PartInfo(LEFTPLATE, "Util_Plate_Metric_1"));
                  
                    if (includeRightPlate)
                      parts.Add(new PartInfo(RIGHTPLATE, "Util_Plate_Metric_1"));                    

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
              
                double boundingBoxWidth = boundingBox.Width,boundingBoxHeight = boundingBox.Height;           

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                //Get Section Structure dimensions
                BusinessObject horizontalSectionPart = componentDictionary[HORSECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double sectionWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double sectionDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                double sectionThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                //Get the Current Location in the Route Connection Cycle
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
                bool[] isOffsetApplied = new bool[2];
                isOffsetApplied = HVACAssemblyServices.GetIsLugEndOffsetApplied(this);

                string[] structPort = new string[2];
                structPort = HVACAssemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0];
                string rightStructPort = structPort[1];

                //Auto Dimensioning of Supports
                if (includeLeftPlate)
                {
                    Note noteDimenssion1 = CreateNote("Dim 1", componentDictionary[LEFTPLATE], "TopStructure");
                    noteDimenssion1.SetPropertyValue("", "IJGeneralNote", "Text");
                    CodelistItem fabrication1 = noteDimenssion1.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    noteDimenssion1.SetPropertyValue(fabrication1, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                    noteDimenssion1.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                }
                Note noteDimenssion2 = CreateNote("Dim 2", componentDictionary[HORSECTION], "BeginCap");
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
                    Note gasketnote = CreateNote("Dim 2", componentDictionary[HORSECTION], "BeginCap");
                    noteDimenssion2.SetPropertyValue(gasketNote, "IJGeneralNote", "Text");
                    CodelistItem fabrication = noteDimenssion2.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    noteDimenssion2.SetPropertyValue(fabrication, "IJGeneralNote", "Purpose"); //value 3 means fabrication
                    noteDimenssion2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                }
                else
                    DeleteNoteIfExists("Dim 2");

                //Set Values of Part Occurance Attributes
                PropertyValueCodelist beginMiterCodelist = (PropertyValueCodelist)componentDictionary[HORSECTION].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMiterCodelist.PropValue == -1)
                    beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist endMiterCodelist = (PropertyValueCodelist)componentDictionary[HORSECTION].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMiterCodelist.PropValue == -1)
                    endMiterCodelist.PropValue = 1;

                // Set Values of Part Occurance Attributes
                componentDictionary[HORSECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[HORSECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[HORSECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[HORSECTION].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[HORSECTION].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");
                componentDictionary[HORSECTION].SetPropertyValue(0.001, "IJUAHgrOccLength", "Length");

                Double leftPlateThickness, rightPlateThickness;

                if (includeLeftPlate)
                    leftPlateThickness = plateThickness;
                else
                    leftPlateThickness = 0;

                if (includeRightPlate)
                    rightPlateThickness = plateThickness;
                else
                    rightPlateThickness = 0;

                componentDictionary[VERTSECTION1].SetPropertyValue(sectionWidth, "IJOAHgrHVACGenBrace", "W");
                componentDictionary[VERTSECTION1].SetPropertyValue(sectionDepth, "IJOAHgrHVACGenBrace", "Depth");
                componentDictionary[VERTSECTION1].SetPropertyValue(sectionThickness, "IJOAHgrHVACGenBrace", "WebT");
                componentDictionary[VERTSECTION1].SetPropertyValue(sectionThickness, "IJOAHgrHVACGenBrace", "FlangeT");
                componentDictionary[VERTSECTION1].SetPropertyValue(Math.PI / 4, "IJOAHgrHVACGenBrace", "CutBackAngle1");
                componentDictionary[VERTSECTION1].SetPropertyValue(Math.PI / 2, "IJOAHgrHVACGenBrace", "CutBackAngle2");

                componentDictionary[VERTSECTION2].SetPropertyValue(sectionWidth, "IJOAHgrHVACGenBrace", "W");
                componentDictionary[VERTSECTION2].SetPropertyValue(sectionDepth, "IJOAHgrHVACGenBrace", "Depth");
                componentDictionary[VERTSECTION2].SetPropertyValue(sectionThickness, "IJOAHgrHVACGenBrace", "WebT");
                componentDictionary[VERTSECTION2].SetPropertyValue(sectionThickness, "IJOAHgrHVACGenBrace", "FlangeT");
                componentDictionary[VERTSECTION2].SetPropertyValue(Math.PI / 2, "IJOAHgrHVACGenBrace", "CutBackAngle1");
                componentDictionary[VERTSECTION2].SetPropertyValue(Math.PI / 4, "IJOAHgrHVACGenBrace", "CutBackAngle2");

                componentDictionary[FLATBAR].SetPropertyValue(flatBarWidth, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[FLATBAR].SetPropertyValue(flatBarThickness, "IJOAHgrUtilMetricT", "T");
              
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
                //Set SectionSize attribute value on the support
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
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
                string structPort1, structPort2, vertSec1BOM, vertSec2BOM;
                double overhangLeft = 0, overhangRight = 0, vertSec1Length = 0, vertSec2Length = 0;
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
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is valid only when the Route is horizontal. For vertical route, use Vertical Duct Support.", "", "RectHorTypeD.cs", 348);
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
                    vertSec1Length = routeStructDist + boundingBoxHeight + flatBarThickness - rightPlateThickness;
                    vertSec2Length = routeStructDist + boundingBoxHeight + flatBarThickness - leftPlateThickness;

                    componentDictionary[VERTSECTION1].SetPropertyValue(vertSec1Length, "IJOAHgrHVACGenBrace", "L");
                    componentDictionary[VERTSECTION2].SetPropertyValue(vertSec2Length, "IJOAHgrHVACGenBrace", "L");
                    componentDictionary[HORSECTION].SetPropertyValue(boundingBoxWidth + 2 * sectionWidth, "IJUAHgrOccLength", "Length");
                    componentDictionary[FLATBAR].SetPropertyValue(boundingBoxWidth, "IJOAHgrUtilMetricWidth", "Width");

                }
                else
                {
                    Double routeStruct1VertDist = 0, routeStruct2VertDist = 0, routeLowStruct1HorDist = 0, routeLowStruct2HorDist = 0, routeHighStruct1HorDist = 0, routeHighStruct2HorDist = 0, horSec1Length = 0, horSec2Length = 0, distBetStruct = 0;

                    if ((Math.Abs(routeAngle) < (0 + 0.001) && Math.Abs(routeAngle) > (0 - 0.001)) || (Math.Abs(routeAngle) < (Math.PI + 0.001) && Math.Abs(routeAngle) > (Math.PI - 0.001)))
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is valid only when the Route is horizontal. For vertical route, use Vertical Duct Support.", "", "RectHorTypeD.cs", 377);
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
                        horSec2Length = boundingBoxWidth;
                    }
                    else if (supportingType.ToUpper().Equals("STEEL"))
                    {
                        horSec1Length = distBetStruct + sectionWidth;
                        horSec2Length = distBetStruct - sectionWidth;
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
                        horSec2Length = boundingBoxWidth + overhangLeft + overhangRight - sectionWidth;
                    }
                    vertSec1Length = routeStruct1VertDist;
                    vertSec2Length = routeStruct2VertDist;
                    componentDictionary[HORSECTION].SetPropertyValue(horSec1Length, "IJUAHgrOccLength", "Length");
                    componentDictionary[FLATBAR].SetPropertyValue(horSec2Length, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[VERTSECTION1].SetPropertyValue(routeStruct1VertDist + sectionDepth, "IJOAHgrHVACGenBrace", "L");
                    componentDictionary[VERTSECTION2].SetPropertyValue(routeStruct2VertDist + sectionDepth, "IJOAHgrHVACGenBrace", "L");
                }              

                vertSec1BOM = "Generic Brace " + sizeOfSection + " Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, vertSec1Length, UnitName.DISTANCE_MILLIMETER);

                vertSec2BOM = "Generic Brace " + sizeOfSection + " Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, vertSec2Length, UnitName.DISTANCE_MILLIMETER);

                componentDictionary[VERTSECTION1].SetPropertyValue(vertSec1BOM, "IJOAHgrHVACGenBrace", "InputBomDesc");
                componentDictionary[VERTSECTION2].SetPropertyValue(vertSec2BOM, "IJOAHgrHVACGenBrace", "InputBomDesc");

                Plane horSec1RoutePlaneA = new Plane(), horSec1RoutePlaneB = new Plane();
                Axis horSec1RouteAxisA = new Axis(), horSec1RouteAxisB = new Axis();                
              
                String horSecPort = string.Empty, routePort = string.Empty;
                Double horSec1RouteAxOffset = 0, horSec1RouteOrgOffset = 0;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    routePort = "BBSR_High";
                else
                    routePort = "BBR_High";
                if (Configuration == 1)
                {                           
                    horSec1RoutePlaneA = Plane.ZX;
                    horSec1RoutePlaneB = Plane.XY;
                    horSec1RouteAxisA = Axis.X;
                    horSec1RouteAxisB = Axis.X;                  

                    horSecPort = "BeginCap";
                    horSec1RouteAxOffset = sectionWidth + overhangRight;
                    horSec1RouteOrgOffset = -sectionDepth / 2;
                }
                else if (Configuration == 2)
                {
                    
                    horSec1RoutePlaneA = Plane.ZX;
                    horSec1RoutePlaneB = Plane.XY;
                    horSec1RouteAxisA = Axis.X;
                    horSec1RouteAxisB = Axis.NegativeX;
               
                    horSecPort = "EndCap";
                    horSec1RouteAxOffset = -sectionWidth - overhangRight;
                    horSec1RouteOrgOffset = -sectionDepth / 2;
                }

                //Create Joints
                //Joints when plates are included
                if (!includeLeftPlate && includeRightPlate)
                {
                    componentDictionary[RIGHTPLATE].SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[RIGHTPLATE].SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[RIGHTPLATE].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    //Add Joint Between the Right plate and the Vertical Section 1
                    JointHelper.CreateRigidJoint(RIGHTPLATE, "BotStructure", VERTSECTION1, "EndStructure", Plane.YZ, Plane.XY, Axis.Y, Axis.Y, -sectionWidth / 2, 0, sectionDepth / 2);
                }
                else if (includeLeftPlate && !includeRightPlate)
                {
                    componentDictionary[LEFTPLATE].SetPropertyValue(sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[LEFTPLATE].SetPropertyValue(sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[LEFTPLATE].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");
                    //Add Joint Between the Left plate and the Vertical Section 2
                    JointHelper.CreateRigidJoint(LEFTPLATE, "TopStructure", VERTSECTION2, "StartStructure", Plane.YZ, Plane.XY, Axis.Y, Axis.Y, -sectionWidth / 2, 0, sectionDepth / 2);

                }
                else if (includeLeftPlate && includeRightPlate)
                {
                    componentDictionary[RIGHTPLATE].SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[RIGHTPLATE].SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[RIGHTPLATE].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    componentDictionary[LEFTPLATE].SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    componentDictionary[LEFTPLATE].SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    componentDictionary[LEFTPLATE].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    //Add Joint Between the Right plate and the Vertical Section 1
                    JointHelper.CreateRigidJoint(LEFTPLATE, "BotStructure", VERTSECTION1, "EndCap", Plane.YZ, Plane.XY, Axis.Y, Axis.Y, -sectionWidth / 2, 0, -sectionDepth / 2);
                    //Add Joint Between the Left plate and the Vertical Section 2
                    JointHelper.CreateRigidJoint(RIGHTPLATE, "TopStructure", VERTSECTION2, "BeginCap", Plane.YZ, Plane.XY, Axis.Y, Axis.Y, -sectionWidth / 2, 0, -sectionDepth / 2);
                }

                //Add bolts
                if (showBolts)
                {
                    JointHelper.CreateRigidJoint(boltPartKeys[0], "StartOther", HORSECTION, "EndCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, -sectionWidth / 2);

                    JointHelper.CreateRigidJoint(boltPartKeys[1], "StartOther", HORSECTION, "BeginCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, sectionWidth / 2);                  
                }


                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(FLATBAR,  "BotStructure", HORSECTION, "Neutral",Plane.ZX, Plane.NegativeXY, Axis.X ,Axis.X, 0, boundingBoxHeight + sectionDepth / 2, sectionDepth / 2 + sectionWidth / 2);
                
                //Add joint between Horizontal Section 2 and Route
                JointHelper.CreateRigidJoint(HORSECTION, horSecPort, "-1", routePort, horSec1RoutePlaneA, horSec1RoutePlaneB,horSec1RouteAxisA,horSec1RouteAxisB, 0, horSec1RouteAxOffset, horSec1RouteOrgOffset);

                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(VERTSECTION1, "StartStructure", HORSECTION, "EndCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeX, boundingBoxHeight + flatBarThickness, sectionDepth,0);

                //Add joint between Horizontal Section 2 and Vertical Section 2
                JointHelper.CreateRigidJoint(VERTSECTION2, "EndStructure", HORSECTION, "BeginCap", Plane.YZ, Plane.NegativeZX, Axis.Y, Axis.NegativeX, -boundingBoxHeight - flatBarThickness, sectionDepth, 0);

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

                    routeConnections.Add(new ConnectionInfo(FLATBAR, 1)); // partindex, routeindex

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

