//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RectVert.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectVert
//   Author       : Manikanth
//   Creation Date:13-06-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   13-06-2013     manikanth  CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//  27.Apr.2015     PVK      TR-CP-253033 Elevation CP not shown by default
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

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

    public class RectVert : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RectVert"
        //----------------------------------------------------------------------------------

        private const string HORSECTION1 = "HORSECTION1";
        private const string HORSECTION2 = "HORSECTION2";
        private const string HORSECTION3 = "HORSECTION3";
        private const string HORSECTION4 = "HORSECTION4";

        private const string HORZ_PARALLEL = "HORZ_PARALLEL";
        private const string HORZ_PERP = "HORZ_PERP";
        string leftPlate = "LEFTPLATE";
        string rightPlate = "RIGHTPLATE";
        string leftBrace = "LEFTBRACE";
        string rightBrace = "RIGHTBRACE";
        string leftBrPlate = "LEFTBRPLATE";
        string rightBrPlate = "RIGHTBRPLATE";
        string leftConnectionObject = "LEFTCONNECTIONOBJECT";
        string rightConnectionObject = "RIGHTCONNECTIONOBJECT";
        string[] boltPartKeys = new string[8];

        double braceAngle, plateThickness;
        string sectionSize, sizeOfSection = string.Empty;
        string boltSize, sizeOfBolt;
        int boltBegin = 0, boltEnd = 0;
        bool showBrace, sectionFromRule, includeLCantiPlate, includeLBrPlate, includeRCantiPlate, includeRBrPlate, showBolts, boltSizeFromRule, value;
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
                    showBrace = (Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBrace", "ShowBrace")).PropValue;
                    braceAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyBrace", "BraceAngle")).PropValue;
                    sectionFromRule = (Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionFromRule")).PropValue;
                    PropertyValueCodelist sectionCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionSize");
                    sectionSize = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionCodeList.PropValue).DisplayName;
                    includeLCantiPlate = (Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyLCantiPlate", "IncludeLCantiPlate")).PropValue;
                    includeLBrPlate = (Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyLBrPlate", "IncludeLBrPlate")).PropValue;
                    includeRCantiPlate = (Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyRCantiPlate", "IncludeRCantiPlate")).PropValue;
                    includeRBrPlate = (Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyRBrPlate", "IncludeRBrPlate")).PropValue;
                    plateThickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyPlateThk", "PlateThick")).PropValue;
                    showBolts = (Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "ShowBolts")).PropValue;
                    boltSizeFromRule = (Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSizeFromRule")).PropValue;
                    PropertyValueCodelist boltCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSize");
                    boltSize = boltCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltCodeList.PropValue).DisplayName;

                    //Get the Section Size
                    if (sectionFromRule)
                        value = GenericHelper.GetDataByRule("HVACHgrSectionSize", (BusinessObject)support, out sizeOfSection);
                    else if (!sectionFromRule)
                    {
                        PropertyValueCodelist sectionSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionSize");
                        sizeOfSection = sectionSizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodeList.PropValue).DisplayName;

                        if (sizeOfSection.ToUpper().Equals("NONE") || sizeOfSection.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RectVert.cs", 101);
                            return null;
                        }
                    }
                    //Get the Bolt Size from Rule
                    if (boltSizeFromRule)
                        value = GenericHelper.GetDataByRule("HVACHgrBoltSize", (BusinessObject)support, out sizeOfBolt);
                    else if (!boltSizeFromRule)
                    {
                        PropertyValueCodelist BoltSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSize");
                        sizeOfBolt = BoltSizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(BoltSizeCodeList.PropValue).DisplayName;

                        if (sizeOfBolt.ToUpper().Equals("NONE") || sizeOfBolt.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Bolt size is not available.", "", "RectVert.cs", 115);
                            return null;
                        }
                    }
                    //Create the list of parts
                    parts.Add(new PartInfo(HORSECTION1, sizeOfSection));
                    parts.Add(new PartInfo(HORSECTION2, sizeOfSection));
                    parts.Add(new PartInfo(HORSECTION3, sizeOfSection));
                    parts.Add(new PartInfo(HORSECTION4, sizeOfSection));

                    int partIndex = parts.Count;
                    if (showBrace)
                    {
                        parts.Add(new PartInfo(leftBrace, "HgrHVACGenericBrace_L1_1"));          //sSectionSize
                        parts.Add(new PartInfo(rightBrace, "HgrHVACGenericBrace_L1_1"));         //sSectionSize   
                        parts.Add(new PartInfo(leftConnectionObject, "Log_Conn_Part_1"));        //Rotational Connection Object                     
                        parts.Add(new PartInfo(rightConnectionObject, "Log_Conn_Part_1"));       //Rotational Connection Object

                        if (includeLBrPlate)
                            parts.Add(new PartInfo(leftBrPlate, "Util_Plate_Metric_1"));
                        if (includeRBrPlate)

                            parts.Add(new PartInfo(rightBrPlate, "Util_Plate_Metric_1"));
                    }
                    if (includeLCantiPlate)
                        parts.Add(new PartInfo(leftPlate, "Util_Plate_Metric_1"));
                    if (includeRCantiPlate)
                        parts.Add(new PartInfo(rightPlate, "Util_Plate_Metric_1"));
                    if (showBolts)
                    {
                        partIndex = 0;
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
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                BoundingBoxHelper.CreateStandardBoundingBoxes(true);
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                //=== ========== == ==== ===========
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========

                //==========================
                //1. Load standard bounding box definition
                //==========================
                BoundingBox boundingBox;

                //=========================
                //2. Get bounding box boundary objects dimension information
                //=========================
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                //====== ======
                //3. retrieve dimension of the bounding box
                //====== ======
                // Get route box geometry
                //
                //  ____________________
                // |                    |
                // |  ROUTE BOX BOUND   | BBXHeight
                // |____________________|
                //        BBXWidth
                //
                double boundingBoxWidth = boundingBox.Width;
                double boundingBoxHeight = boundingBox.Height;

                BusinessObject horizontalSectionPart = (componentDictionary[HORSECTION1]).GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                if (includeLCantiPlate)
                {
                    Note note1 = CreateNote("Dim1", (componentDictionary[leftPlate]), "TopStructure");
                    note1.SetPropertyValue("", "IJGeneralNote", "Text");
                    PropertyValueCodelist note1PropertyValueCL = (PropertyValueCodelist)note1.GetPropertyValue("IJGeneralNote", "Purpose");
                    CodelistItem codeList1 = note1PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                    note1.SetPropertyValue(codeList1, "IJGeneralNote", "Purpose");
                    note1.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                }
                Note note2 = CreateNote("Dim2", (componentDictionary[HORSECTION3]), "EndCap");
                note2.SetPropertyValue("", "IJGeneralNote", "Text");
                PropertyValueCodelist note2PropertyValueCL = (PropertyValueCodelist)note2.GetPropertyValue("IJGeneralNote", "Purpose");
                CodelistItem codeList2 = note2PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                note2.SetPropertyValue(codeList2, "IJGeneralNote", "Purpose");
                note2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                if (showBrace)
                {
                    Note note3 = CreateNote("Dim3", (componentDictionary[leftBrace]), "BeginCap");
                    note3.SetPropertyValue("", "IJGeneralNote", "Text");
                    PropertyValueCodelist note3PropertyValueCL = (PropertyValueCodelist)note3.GetPropertyValue("IJGeneralNote", "Purpose");
                    CodelistItem codeList3 = note3PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                    note3.SetPropertyValue(codeList3, "IJGeneralNote", "Purpose");
                    note3.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                    Note note4 = CreateNote("Dim4", (componentDictionary[rightBrace]), "EndCap");
                    note4.SetPropertyValue("", "IJGeneralNote", "Text");
                    PropertyValueCodelist note4PropertyValueCL = (PropertyValueCodelist)note4.GetPropertyValue("IJGeneralNote", "Purpose");
                    CodelistItem codeList4 = note2PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                    note4.SetPropertyValue(codeList2, "IJGeneralNote", "Purpose");
                    note4.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                }
                string gasketNote;
                string gasketBomDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACGasketDesc", "GasketDesc")).PropValue;
                if (gasketBomDesc == null)
                    gasketNote = "To line with 3mm thick Gasket all around for all contacted surface between Duct & Support";
                else
                    gasketNote = gasketBomDesc;

                bool excludeNote;
                if (support.SupportsInterface("IJUAhsExcludeNotes"))
                    excludeNote = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;
                else
                    excludeNote = false;

                if (!excludeNote)
                {
                    Note note5 = CreateNote("Note 1");
                    PropertyValueCodelist note5PropertyValueCL = (PropertyValueCodelist)note5.GetPropertyValue("IJGeneralNote", "Purpose");
                    CodelistItem codeList5 = note5PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                    note5.SetPropertyValue(gasketNote, "IJGeneralNote", "Text");
                }
                else
                    DeleteNoteIfExists("Note 1");
                //Width, Depth, Flange, or Web as 2nd argument
                double sectionDepth = crosssection.Depth;
                double sectionWidth = crosssection.Width;
                double steelThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                Boolean[] bIsOffsetApplied = HVACAssemblyServices.GetIsLugEndOffsetApplied(this);
                string[] idxStructPort = HVACAssemblyServices.GetIndexedStructPortName(this, bIsOffsetApplied);
                string leftStructPort = idxStructPort[0];
                string rightStructPort = idxStructPort[1];

                //Determine whether Route is vertical and steel is besides the route
                double routeAngle = RefPortHelper.AngleBetweenPorts("route", PortAxisType.X, OrientationAlong.Global_Z);

                if (Math.Abs(routeAngle) < (Math.PI / 2 + 0.001) && Math.Abs(routeAngle) > (Math.PI / 2 - 0.001))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is valid only when the Route is vertical and the supporting Structure is besides the Route..", "", "RectVert.cs", 266);
                    return;
                }

                PropertyValueCodelist hor1beginMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION1]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (hor1beginMiterCodelist.PropValue == -1)
                    hor1beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor1endMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION1]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (hor1endMiterCodelist.PropValue == -1)
                    hor1endMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor2beginMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION2]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (hor2beginMiterCodelist.PropValue == -1)
                    hor2beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor2endMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION2]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (hor2endMiterCodelist.PropValue == -1)
                    hor2endMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor3beginMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION3]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (hor3beginMiterCodelist.PropValue == -1)
                    hor3beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor3endMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION3]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (hor3endMiterCodelist.PropValue == -1)
                    hor3endMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor4beginMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION4]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (hor4beginMiterCodelist.PropValue == -1)
                    hor4beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor4endMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION4]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (hor4endMiterCodelist.PropValue == -1)
                    hor4endMiterCodelist.PropValue = 1;
                //  =================================
                //' Set Values of Part Occurance Attributes
                //'======================================

                (componentDictionary[HORSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HORSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[HORSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[HORSECTION1]).SetPropertyValue(hor1beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[HORSECTION1]).SetPropertyValue(hor1endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                (componentDictionary[HORSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HORSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[HORSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[HORSECTION2]).SetPropertyValue(hor2beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[HORSECTION2]).SetPropertyValue(hor2endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                (componentDictionary[HORSECTION3]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HORSECTION3]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[HORSECTION3]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[HORSECTION3]).SetPropertyValue(hor3beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[HORSECTION3]).SetPropertyValue(hor3endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                (componentDictionary[HORSECTION4]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HORSECTION4]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[HORSECTION4]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[HORSECTION4]).SetPropertyValue(hor4beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[HORSECTION4]).SetPropertyValue(hor4endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                string boltDesc;
                string boltBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACBoltsDesc", "BoltDesc")).PropValue;
                if (boltBOMDesc == null)
                    boltDesc = boltSize + " " + "Bolts";
                else
                    boltDesc = boltBOMDesc;

                if (showBolts)
                {
                    for (int boltIndex = boltBegin; boltIndex <= boltEnd; boltIndex++)
                    {
                        componentDictionary[boltPartKeys[boltIndex - boltBegin]].SetPropertyValue(sectionDepth / 2, "IJOAHgrUtilMetricL", "L");
                        componentDictionary[boltPartKeys[boltIndex - boltBegin]].SetPropertyValue(sectionWidth / 4, "IJOAHgrUtilMetricRadius", "Radius");
                        componentDictionary[boltPartKeys[boltIndex - boltBegin]].SetPropertyValue(boltDesc, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
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
                if (boltSizeFromRule)
                {
                    CodelistItem codeList = metadataManager.GetCodelistInfo("HVACAssyHgrBoltSize", "UDP").GetCodelistItem(sizeOfBolt);
                    support.SetPropertyValue(codeList.Value, "IJOAHgrHVACAssyBolts", "BoltSize");
                }
                double hor1SecLength = boundingBoxWidth + 2 * sectionWidth;
                double hor2SecLength = hor1SecLength;
                double hor3SecLength = boundingBoxHeight;
                double hor4SecLength = hor3SecLength;

                (componentDictionary[HORSECTION1]).SetPropertyValue(hor1SecLength, "IJUAHgrOccLength", "Length");
                (componentDictionary[HORSECTION2]).SetPropertyValue(hor2SecLength, "IJUAHgrOccLength", "Length");
                (componentDictionary[HORSECTION3]).SetPropertyValue(hor3SecLength, "IJUAHgrOccLength", "Length");
                (componentDictionary[HORSECTION4]).SetPropertyValue(hor4SecLength, "IJUAHgrOccLength", "Length");

                double routeStructAngle = RefPortHelper.AngleBetweenPorts("BBR_High", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                double routeHightStructDistance;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    routeHightStructDistance = RefPortHelper.DistanceBetweenPorts("Structure", "BBSR_High", PortDistanceType.Horizontal);
                else
                    if (leftStructPort != rightStructPort)
                        routeHightStructDistance = RefPortHelper.DistanceBetweenPorts(leftStructPort, "BBR_High", PortDistanceType.Horizontal_Perpendicular);
                    else
                        routeHightStructDistance = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_High", PortDistanceType.Horizontal_Perpendicular);

                (componentDictionary[HORSECTION3]).SetPropertyValue(routeHightStructDistance, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HORSECTION3]).SetPropertyValue(sectionDepth, "IJUAHgrOccOverLength", "BeginOverLength");

                (componentDictionary[HORSECTION4]).SetPropertyValue(routeHightStructDistance, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[HORSECTION4]).SetPropertyValue(sectionDepth, "IJUAHgrOccOverLength", "EndOverLength");

                double horizontalSectionLength, braceLength = 0, brPadOffset = 0;
                string braceBOM;
                if (showBrace)
                {
                    horizontalSectionLength = hor3SecLength + routeHightStructDistance;
                    braceLength = horizontalSectionLength / Math.Cos(braceAngle);
                    brPadOffset = braceLength * Math.Sin(braceAngle) - sectionDepth;
                   
                    braceBOM = "Generic Brace " + sizeOfSection + " Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, braceLength, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.DISTANCE_INCH_SYMBOL);

                    (componentDictionary[leftBrace]).SetPropertyValue(braceLength, "IJOAHgrHVACGenBrace", "L");
                    (componentDictionary[leftBrace]).SetPropertyValue(Math.PI / 2 - braceAngle, "IJOAHgrHVACGenBrace", "Angle");
                    PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[leftBrace].GetPropertyValue("IJOAHgrHVACGenBrace", "BraceOrient");
                    braceCodelist.PropValue = 2;
                    (componentDictionary[leftBrace]).SetPropertyValue(braceCodelist.PropValue, "IJOAHgrHVACGenBrace", "BraceOrient");
                    (componentDictionary[leftBrace]).SetPropertyValue(sectionWidth, "IJOAHgrHVACGenBrace", "W");
                    (componentDictionary[leftBrace]).SetPropertyValue(sectionDepth, "IJOAHgrHVACGenBrace", "Depth");
                    (componentDictionary[leftBrace]).SetPropertyValue(steelThickness, "IJOAHgrHVACGenBrace", "FlangeT");
                    (componentDictionary[leftBrace]).SetPropertyValue(-braceAngle + Math.PI / 2, "IJOAHgrHVACGenBrace", "CutBackAngle1");
                    (componentDictionary[leftBrace]).SetPropertyValue(braceAngle, "IJOAHgrHVACGenBrace", "CutBackAngle2");
                    (componentDictionary[leftBrace]).SetPropertyValue(braceBOM, "IJOAHgrHVACGenBrace", "InputBomDesc");


                    (componentDictionary[rightBrace]).SetPropertyValue(braceLength, "IJOAHgrHVACGenBrace", "L");
                    (componentDictionary[rightBrace]).SetPropertyValue(Math.PI / 2 - braceAngle, "IJOAHgrHVACGenBrace", "Angle");
                    PropertyValueCodelist braceCodelist1 = (PropertyValueCodelist)componentDictionary[rightBrace].GetPropertyValue("IJOAHgrHVACGenBrace", "BraceOrient");
                    braceCodelist1.PropValue = 2;
                    (componentDictionary[rightBrace]).SetPropertyValue(braceCodelist1.PropValue, "IJOAHgrHVACGenBrace", "BraceOrient");
                    (componentDictionary[rightBrace]).SetPropertyValue(sectionWidth, "IJOAHgrHVACGenBrace", "W");
                    (componentDictionary[rightBrace]).SetPropertyValue(sectionDepth, "IJOAHgrHVACGenBrace", "Depth");
                    (componentDictionary[rightBrace]).SetPropertyValue(steelThickness, "IJOAHgrHVACGenBrace", "FlangeT");
                    (componentDictionary[rightBrace]).SetPropertyValue(braceAngle, "IJOAHgrHVACGenBrace", "CutBackAngle1");
                    (componentDictionary[rightBrace]).SetPropertyValue(Math.PI / 2 - braceAngle, "IJOAHgrHVACGenBrace", "CutBackAngle2");
                    (componentDictionary[rightBrace]).SetPropertyValue(braceBOM, "IJOAHgrHVACGenBrace", "InputBomDesc");
                }

                string routePort1 = string.Empty, routePort2 = string.Empty, horSecPort = string.Empty;
                double horSecRouteAxOffset = 0;

                Plane planeSec1Route1, planesec1Route2, planeSec2Route1, planeSec2Route2;
                planeSec1Route1 = planesec1Route2 = planeSec2Route1 = planeSec2Route2 = new Plane();
                Axis axisSec1Route1, axissec1Route2, axisSec2Route1, axisSec2Route2;
                axisSec1Route1 = axissec1Route2 = axisSec2Route1 = axisSec2Route2 = new Axis();
                if (Configuration == 1)
                {
                    planeSec1Route1 = Plane.ZX;
                    planesec1Route2 = Plane.NegativeXY;
                    axisSec1Route1 = Axis.X;
                    axissec1Route2 = Axis.X;
                    planeSec2Route1 = Plane.ZX;
                    planeSec2Route2 = Plane.XY;
                    axisSec2Route1 = Axis.X;
                    axisSec2Route2 = Axis.X;

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
                    horSecRouteAxOffset = sectionWidth;
                    horSecPort = "BeginCap";
                }
                else if (Configuration == 2)
                {
                    planeSec1Route1 = Plane.ZX;
                    planesec1Route2 = Plane.NegativeXY;
                    axisSec1Route1 = Axis.X;
                    axissec1Route2 = Axis.NegativeX;
                    planeSec2Route1 = Plane.ZX;
                    planeSec2Route2 = Plane.XY;
                    axisSec2Route1 = Axis.X;
                    axisSec2Route2 = Axis.NegativeX;

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
                    horSecRouteAxOffset = -sectionWidth;
                    horSecPort = "EndCap";
                }

                if (showBrace)
                {
                    //Add a Joint between horizontal section and the point that we want the angle to rotate around.
                    JointHelper.CreateRigidJoint(leftConnectionObject, "Connection", HORSECTION3, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X, 0, -steelThickness, -2 * steelThickness);

                    JointHelper.CreateRigidJoint(leftConnectionObject, "Connection", leftBrace, "EndCap", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0, 0);

                    //Add a Joint between horizontal section and the point that we want the angle to rotate around.
                    JointHelper.CreateRigidJoint(rightConnectionObject, "Connection", HORSECTION4, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, 0, -steelThickness, 2 * steelThickness);

                    JointHelper.CreateRigidJoint(rightConnectionObject, "Connection", rightBrace, "BeginCap", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0, 0);

                    if (includeLBrPlate)
                    {
                        (componentDictionary[leftBrace]).SetPropertyValue(braceLength - (plateThickness / Math.Cos(braceAngle)), "IJOAHgrHVACGenBrace", "L");
                        (componentDictionary[leftBrPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                        (componentDictionary[leftBrPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                        (componentDictionary[leftBrPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                        //Add Joint Between the Plate and the Horizontal Section
                        JointHelper.CreateRigidJoint(leftBrPlate, "BotStructure", HORSECTION3, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, -routeHightStructDistance, -sectionDepth / 2, -brPadOffset - 2 * sectionWidth / 3);
                    }
                    if (includeRBrPlate)
                    {
                        (componentDictionary[rightBrace]).SetPropertyValue(braceLength - (plateThickness / Math.Cos(braceAngle)), "IJOAHgrHVACGenBrace", "L");
                        (componentDictionary[rightBrPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                        (componentDictionary[rightBrPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                        (componentDictionary[rightBrPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                        //Add Joint Between the Plate and the Horizontal Section
                        JointHelper.CreateRigidJoint(rightBrPlate, "TopStructure", HORSECTION4, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, routeHightStructDistance, -sectionDepth / 2, -brPadOffset - 2 * sectionWidth / 3);

                    }
                }
                if (includeLCantiPlate && !includeRCantiPlate)
                {
                    (componentDictionary[HORSECTION3]).SetPropertyValue(routeHightStructDistance - plateThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[leftPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(leftPlate, "BotStructure", HORSECTION3, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, -routeHightStructDistance, -sectionWidth / 2, sectionDepth / 2);
                }
                if (!includeLCantiPlate && includeRCantiPlate)
                {
                    (componentDictionary[HORSECTION4]).SetPropertyValue(routeHightStructDistance - plateThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[rightPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(rightPlate, "TopStructure", HORSECTION4, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, routeHightStructDistance, -sectionWidth / 2, sectionDepth / 2);
                }
                if (includeLCantiPlate && includeRCantiPlate)
                {
                    (componentDictionary[HORSECTION3]).SetPropertyValue(routeHightStructDistance - plateThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[leftPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(leftPlate, "BotStructure", HORSECTION3, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, -routeHightStructDistance, -sectionWidth / 2, sectionDepth / 2);

                    (componentDictionary[HORSECTION4]).SetPropertyValue(routeHightStructDistance - plateThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[rightPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(rightPlate, "TopStructure", HORSECTION4, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, routeHightStructDistance, -sectionWidth / 2, sectionDepth / 2);
                }
                if (showBolts)
                {
                    JointHelper.CreateRigidJoint(boltPartKeys[0], "StartOther", HORSECTION1, "EndCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, -sectionWidth / 2);
                    JointHelper.CreateRigidJoint(boltPartKeys[1], "StartOther", HORSECTION1, "BeginCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, sectionWidth / 2);
                    JointHelper.CreateRigidJoint(boltPartKeys[2], "StartOther", HORSECTION2, "BeginCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, sectionWidth / 2);
                    JointHelper.CreateRigidJoint(boltPartKeys[3], "StartOther", HORSECTION3, "BeginCap", Plane.YZ, Plane.ZX, Axis.Y, Axis.NegativeZ, -sectionDepth / 2, sectionDepth / 4, sectionWidth / 2);
                }

                //Add joint between Horizontal Section 1 and Route
                JointHelper.CreateRigidJoint(HORSECTION1, horSecPort, "-1", routePort1, planeSec1Route1, planesec1Route2, axisSec1Route1, axissec1Route2, 0, horSecRouteAxOffset, -sectionDepth / 2);

                //Add joint between Horizontal Section 2 and Route
                JointHelper.CreateRigidJoint(HORSECTION2, horSecPort, "-1", routePort2, planeSec2Route1, planeSec2Route2, axisSec2Route1, axisSec2Route2, 0, horSecRouteAxOffset, -sectionDepth / 2);

                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(HORSECTION3, "BeginCap", HORSECTION1, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX, sectionDepth, 0, 0);

                //Add joint between Horizontal Section 2 and Vertical Section 2
                JointHelper.CreateRigidJoint(HORSECTION4, "EndCap", HORSECTION1, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, sectionDepth, 0, 0);
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

                    structConnections.Add(new ConnectionInfo(HORSECTION3, 1)); // partindex, routeindex

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
    }
}
