//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RndHorTypeB.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RndHorTypeB
//   Author       :Manikanth
//   Creation Date:13-06-2014
//   Description:


//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   13-06-2013     manikanth  CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//  22-Jan-2015     PVK        TR-CP-264951  Resolve coverity issues found in November 2014 report
//  27.Apr.2015     PVK        TR-CP-253033 Elevation CP not shown by default
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
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
    public class RndHorTypeB : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RndHorTypeB"
        //----------------------------------------------------------------------------------

        //Constants
        //Define your Constants here
        private const string HORSECTION = "HORSECTION";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";

        string[] routePartKeys = new string[8];
        string[] boltPartKeys = new string[8];
        double[] ductRadius = new double[6];

        string leftPlate = "LEFTPLATE", rightPlate = "RIGHTPLATE", sectionSize, flatBarDim, sizeOfSection = string.Empty, sizeOfBolt, boltSize;
        double flatBarThickness, dxClampEnd, overHangLeft, overHangRight, plateThickness, clampLegLength, flatBarWidth;
        int boltBegin, boltEnd, dxClamp = 0, dxClampBegin = 0, numRoutes, dxRoute = 0;
        bool sectionFromRule, boltSizeFromRule, showBolts, includeLeftPlate, includeRigthPlate, value;
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
                    includeRigthPlate = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyRPlate", "IncludeRightPlate")).PropValue;
                    plateThickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyPlateThk", "PlateThick")).PropValue;
                    overHangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyOHLeft", "OverhangLeft")).PropValue;
                    overHangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyOHRight", "OverhangRight")).PropValue;
                    flatBarDim = (string)((PropertyValueString)support.GetPropertyValue("IJOAHgrHVACAssyFlatBarDim", "FlatBarDim")).PropValue;
                    clampLegLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyClamp", "LegLength")).PropValue;
                    showBolts = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "ShowBolts")).PropValue;
                    boltSizeFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSizeFromRule")).PropValue;
                    PropertyValueCodelist boltCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSize");
                    boltSize = boltCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltCodeList.PropValue).DisplayName;

                    string[] position = flatBarDim.Split('x');
                    flatBarWidth = double.Parse(position[0]) / 1000;
                    flatBarThickness = double.Parse(position[1]) / 1000;

                    //Get the Section Size from Rule

                    if (sectionFromRule)
                        value = GenericHelper.GetDataByRule("HVACHgrSectionSize", (BusinessObject)support, out sizeOfSection);
                    else if (!sectionFromRule)
                    {
                        PropertyValueCodelist sectionSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionSize");
                        sizeOfSection = sectionSizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodeList.PropValue).DisplayName;

                        if (sizeOfSection.ToUpper().Equals("NONE") || sizeOfSection.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RndHorTypeB.cs", 97);
                            return null;
                        }
                    }
                    numRoutes = SupportHelper.SupportedObjects.Count;
                    for (dxRoute = 1; dxRoute <= numRoutes; dxRoute++)
                    {
                        DuctObjectInfo ductInfo = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(dxRoute);
                        ductRadius[dxRoute] = ductInfo.OutsideDiameter / 2;
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
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Bolt size is not available.", "", "RndHorTypeB.cs", 117);
                            return null;
                        }
                    }

                    int partCount = 0;
                    parts.Add(new PartInfo(HORSECTION, sizeOfSection));
                    parts.Add(new PartInfo(VERTSECTION1, sizeOfSection));
                    parts.Add(new PartInfo(VERTSECTION2, sizeOfSection));

                    dxClampEnd = dxClampBegin - 1 + numRoutes;

                    for (dxClamp = dxClampBegin; dxClamp <= dxClampEnd; dxClamp++)
                    {
                        routePartKeys[dxClamp] = "Clamp" + (dxClamp).ToString();
                        parts.Add(new PartInfo(routePartKeys[dxClamp], "HgrHVACClamp1_1"));
                    }
                    if (includeLeftPlate)
                        parts.Add(new PartInfo(leftPlate, "Util_Plate_Metric_1"));
                    if (includeRigthPlate)
                        parts.Add(new PartInfo(rightPlate, "Util_Plate_Metric_1"));
                    if (showBolts)
                    {
                        partCount = 0;
                        boltBegin = parts.Count + 1;
                        boltEnd = Convert.ToInt32(boltBegin + numRoutes * 2 - 1);
                        for (int index = boltBegin; index <= boltEnd; index++)
                        {
                            boltPartKeys[partCount] = "Bolt" + (partCount + 1).ToString();
                            parts.Add(new PartInfo(boltPartKeys[partCount], "Util_Fixed_Cyl_Metric_1"));
                            partCount++;
                        }
                    }
                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in GetCatalogParts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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

                BusinessObject horizontalSectionPart = (componentDictionary[HORSECTION]).GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                Note note1 = CreateNote("Dim 1", (componentDictionary[VERTSECTION2]), "EndCap");
                note1.SetPropertyValue("", "IJGeneralNote", "Text");
                PropertyValueCodelist note1PropertyValueCL = (PropertyValueCodelist)note1.GetPropertyValue("IJGeneralNote", "Purpose");
                CodelistItem codeList1 = note1PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                note1.SetPropertyValue(codeList1, "IJGeneralNote", "Purpose");

                Note note2 = CreateNote("Dim 2", (componentDictionary[routePartKeys[dxClampBegin]]), "StartOther");
                note2.SetPropertyValue("", "IJGeneralNote", "Text");
                PropertyValueCodelist note2PropertyValueCL = (PropertyValueCodelist)note2.GetPropertyValue("IJGeneralNote", "Purpose");
                CodelistItem codeList2 = note2PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                note2.SetPropertyValue(codeList1, "IJGeneralNote", "Purpose");

                double sectionDepth = crosssection.Depth;
                double sectionWidth = crosssection.Width;
                double steelThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                string gasketNote;
                string gasketBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACGasketDesc", "GasketDesc")).PropValue;
                if (gasketBOMDesc == null)
                    gasketNote = "To line with 3mm thick Gasket all around for all contacted surface between Duct & Support";    //Check for "ExcludeNotes" attribute (for migrated DB)
                else
                    gasketNote = gasketBOMDesc;

                bool excludeNote;
                if (support.SupportsInterface("IJUAhsExcludeNotes"))
                    excludeNote = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;
                else
                    excludeNote = false;

                if (!excludeNote)
                {
                    Note note3 = CreateNote("Note 1");
                    PropertyValueCodelist note5PropertyValueCL = (PropertyValueCodelist)note3.GetPropertyValue("IJGeneralNote", "Purpose");
                    CodelistItem codeList5 = note5PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                    note3.SetPropertyValue("", "IJGeneralNote", "Text");
                }
                else
                    DeleteNoteIfExists("Note 1");
                //=======================================
                //Do Something if more than one Structure
                //=======================================
                //get structure count
                Boolean[] bIsOffsetApplied = HVACAssemblyServices.GetIsLugEndOffsetApplied(this);
                string[] idxStructPort = HVACAssemblyServices.GetIndexedStructPortName(this, bIsOffsetApplied);
                string leftStructPort = idxStructPort[0];
                string rightStructPort = idxStructPort[1];

                PropertyValueCodelist hor1beginMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (hor1beginMiterCodelist.PropValue == -1)
                    hor1beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor1endMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (hor1endMiterCodelist.PropValue == -1)
                    hor1endMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor2beginMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION1]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (hor2beginMiterCodelist.PropValue == -1)
                    hor2beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor2endMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION1]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (hor2endMiterCodelist.PropValue == -1)
                    hor2endMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor3beginMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION2]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (hor3beginMiterCodelist.PropValue == -1)
                    hor3beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor3endMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION2]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (hor3endMiterCodelist.PropValue == -1)
                    hor3endMiterCodelist.PropValue = 1;
                
                //========================================
                // Set Values of Part Occurance Attributes
                //========================================
                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[HORSECTION]).SetPropertyValue(hor1beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[HORSECTION]).SetPropertyValue(hor1endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                (componentDictionary[VERTSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(hor2beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(hor2endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                (componentDictionary[VERTSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(hor3beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(hor3endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                string boltDesc;
                string boltBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACBoltsDesc", "BoltDesc")).PropValue;
                if (string.IsNullOrEmpty(boltBOMDesc))
                    boltDesc = boltSize + " " + "Bolts";
                else
                    boltDesc = boltBOMDesc;

                if (showBolts)
                {
                    for (int boltIndex = boltBegin; boltIndex <= boltEnd; boltIndex++)
                    {
                        componentDictionary[boltPartKeys[boltIndex - boltBegin]].SetPropertyValue(2 * (flatBarThickness + steelThickness), "IJOAHgrUtilMetricL", "L");
                        componentDictionary[boltPartKeys[boltIndex - boltBegin]].SetPropertyValue(clampLegLength / 8, "IJOAHgrUtilMetricRadius", "Radius");
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
                numRoutes = SupportHelper.SupportedObjects.Count;
                for (dxClamp = dxClampBegin; dxClamp <= dxClampEnd; dxClamp++)
                {
                    if (dxClamp == dxClampBegin)
                        dxRoute = 1;
                    else
                        dxRoute = dxClamp - dxClampBegin + 1;
                    string clampBomDesc = "Clamp for " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, 2 * ductRadius[dxRoute], UnitName.DISTANCE_MILLIMETER) + " Duct outside diameter";

                    (componentDictionary[routePartKeys[dxClamp]]).SetPropertyValue(clampLegLength, "IJUAHgrHVACClamp", "LegLength");
                    (componentDictionary[routePartKeys[dxClamp]]).SetPropertyValue(flatBarWidth, "IJUAHgrHVACClamp", "Width");
                    (componentDictionary[routePartKeys[dxClamp]]).SetPropertyValue(flatBarThickness, "IJUAHgrHVACClampThk", "Thickness");
                    (componentDictionary[routePartKeys[dxClamp]]).SetPropertyValue(ductRadius[dxRoute], "IJUAHgrOccLength", "Length");
                    (componentDictionary[routePartKeys[dxClamp]]).SetPropertyValue(ductRadius[dxRoute], "IJUAHgrHVACClamp", "Radius");
                    (componentDictionary[routePartKeys[dxClamp]]).SetPropertyValue(clampBomDesc, "IJOAHgrHVACBomDesc", "InputBomDesc1");
                }
                //Get the route structure distance
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);

                double routeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                double routeStructDistance = 0;

                string structPort1, structPort2;
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
                double distBetStruct = 0, routeStruct1Dist = 0, routeStruct2Dist = 0, horSecLength = 0, vertSec1AxisOffset = 0, vertSec2AxisOffset = 0, routeLowStruct1HorDist, routeLowStruct2HorDist, routeHighStruct1HorDist, routeHighStruct2HorDist;

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
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if ((Math.Abs(routeAngle) < (0 + 0.001) && Math.Abs(routeAngle) > (0 - 0.001)) || (Math.Abs(routeAngle) < (Math.PI + 0.001) && Math.Abs(routeAngle) > (Math.PI - 0.001)))
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is valid only when the Route is horizontal. For vertical route, use Vertical Duct Support.", "", "RectHorTypeB.cs", 375);
                        return;
                    }
                    else
                    {
                        if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && (Math.Abs(routeStructAngle) > (Math.PI - 0.001))))
                            routeStructDistance = RefPortHelper.DistanceBetweenPorts("BBSR_High", "Structure", PortDistanceType.Horizontal);
                        else
                            routeStructDistance = RefPortHelper.DistanceBetweenPorts("BBSR_High", "Structure", PortDistanceType.Vertical);
                    }
                    (componentDictionary[VERTSECTION1]).SetPropertyValue(routeStructDistance, "IJUAHgrOccLength", "Length");
                    (componentDictionary[VERTSECTION2]).SetPropertyValue(routeStructDistance, "IJUAHgrOccLength", "Length");
                    (componentDictionary[HORSECTION]).SetPropertyValue(boundingBoxWidth, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    if ((Math.Abs(routeAngle) < (0 + 0.001) && Math.Abs(routeAngle) > (0 - 0.001)) || (Math.Abs(routeAngle) < (Math.PI + 0.001) && Math.Abs(routeAngle) > (Math.PI - 0.001)))
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is valid only when the Route is horizontal. For vertical route, use Vertical Duct Support.", "", "RectHorTypeB.cs", 393);
                        return;
                    }
                    else
                    {
                        if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && (Math.Abs(routeStructAngle) > (Math.PI - 0.001))))
                        {
                            routeStruct1Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", structPort1, PortDistanceType.Horizontal);
                            routeStruct2Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", structPort2, PortDistanceType.Horizontal);
                            routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Vertical);
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Vertical);
                            routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Vertical);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Vertical);
                            distBetStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Vertical);
                        }
                        else
                        {
                            routeStruct1Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", structPort1, PortDistanceType.Vertical);
                            routeStruct2Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", structPort2, PortDistanceType.Vertical);
                            routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Horizontal);
                            routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Horizontal);
                            routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Horizontal);
                            routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Horizontal);
                            distBetStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
                        }

                        if (supportingType.ToUpper().Equals("SLAB"))
                            horSecLength = boundingBoxWidth;
                        else if (supportingType.ToUpper().Equals("STEEL"))
                        {
                            horSecLength = boundingBoxWidth;
                            if (routeLowStruct2HorDist < routeHighStruct2HorDist)
                                overHangLeft = routeLowStruct2HorDist - sectionWidth;
                            else
                                overHangLeft = routeHighStruct2HorDist - sectionWidth;
                            if (routeLowStruct1HorDist < routeHighStruct1HorDist)
                                overHangRight = routeLowStruct1HorDist - sectionWidth;
                            else
                                overHangRight = routeHighStruct1HorDist - sectionWidth;
                        }
                        else if (supportingType.ToUpper().Equals("STEEL-SLAB") || supportingType.ToUpper().Equals("SLAB-STEEL"))
                        {
                            if (routeLowStruct2HorDist < routeHighStruct2HorDist)
                                overHangLeft = routeLowStruct2HorDist - sectionWidth;
                            else
                                overHangLeft = routeHighStruct2HorDist - sectionWidth;
                            if (routeLowStruct1HorDist < routeHighStruct1HorDist)
                                overHangRight = routeLowStruct1HorDist - sectionWidth;
                            else
                                overHangRight = routeHighStruct1HorDist - sectionWidth;

                            horSecLength = boundingBoxWidth;
                        }
                        (componentDictionary[VERTSECTION1]).SetPropertyValue(routeStruct1Dist, "IJUAHgrOccLength", "Length");
                        (componentDictionary[VERTSECTION2]).SetPropertyValue(routeStruct2Dist, "IJUAHgrOccLength", "Length");
                        (componentDictionary[HORSECTION]).SetPropertyValue(horSecLength, "IJUAHgrOccLength", "Length");
                    }
                }
                Plane plane1 = new Plane(); Plane plane2 = new Plane();
                Axis axis1 = new Axis(); Axis axis2 = new Axis();
                string horSecPort = string.Empty, routePort = string.Empty;

                if (Configuration == 1)
                {
                    plane1 = Plane.YZ;
                    plane2 = Plane.XY;
                    axis1 = Axis.Y;
                    axis2 = Axis.NegativeX;
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        routePort = "BBSR_High";
                    else
                        routePort = "BBR_High";
                    horSecPort = "BeginCap";
                    vertSec1AxisOffset = -overHangRight;
                    vertSec2AxisOffset = overHangLeft;
                    (componentDictionary[HORSECTION]).SetPropertyValue(overHangRight, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[HORSECTION]).SetPropertyValue(overHangLeft, "IJUAHgrOccOverLength", "BeginOverLength");
                }
                else if (Configuration == 2)
                {
                    plane1 = Plane.YZ;
                    plane2 = Plane.XY;
                    axis1 = Axis.Y;
                    axis2 = Axis.X;
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        routePort = "BBSR_High";
                    else
                        routePort = "BBR_High";

                    horSecPort = "EndCap";
                    vertSec1AxisOffset = -overHangLeft;
                    vertSec2AxisOffset = overHangRight;

                    (componentDictionary[HORSECTION]).SetPropertyValue(overHangLeft, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[HORSECTION]).SetPropertyValue(overHangRight, "IJUAHgrOccOverLength", "BeginOverLength");
                }
                string refPortName;
                int bolts = 0;
                for (dxClamp = dxClampBegin; dxClamp <= dxClampEnd; dxClamp++)
                {
                    if (dxClamp == dxClampBegin)
                        refPortName = "Route";
                    else
                        refPortName = "Route_" + (dxClamp - dxClampBegin + 1).ToString();

                    //Add Joint Between the UBolt and Route
                    JointHelper.CreateRigidJoint(routePartKeys[dxClamp], "Route", "-1", refPortName, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    if (showBolts)
                    {
                        //Add Joint Between the UBolt and Route
                        JointHelper.CreateRigidJoint(routePartKeys[dxClamp], "StartOther", boltPartKeys[bolts], "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -(flatBarThickness + steelThickness), 0, 0);
                        bolts++;
                        JointHelper.CreateRigidJoint(routePartKeys[dxClamp], "EndOther", boltPartKeys[bolts], "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -(flatBarThickness + steelThickness), 0, 0);
                        bolts++;
                    }
                }
                if (numRoutes == 1)
                    //Add joint between Horizontal Section 2 and Route
                    JointHelper.CreateRigidJoint(HORSECTION, "Neutral", "-1", "Route", plane1, plane2, axis1, axis2, -ductRadius[dxClamp] - sectionDepth / 2, 0, 0);
                else
                    //Add joint between Horizontal Section 2 and Route
                    JointHelper.CreateRigidJoint(HORSECTION, horSecPort, "-1", routePort, plane1, plane2, axis1, axis2, 0, 0, sectionDepth / 2);

                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(VERTSECTION1, "BeginCap", HORSECTION, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, 0, vertSec1AxisOffset, 0);

                //Add joint between Horizontal Section 2 and Vertical Section 2
                JointHelper.CreateRigidJoint(VERTSECTION2, "EndCap", HORSECTION, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0, -vertSec2AxisOffset, 0);

                if (!includeLeftPlate && includeRigthPlate)
                {
                    (componentDictionary[VERTSECTION1]).SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[rightPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    //Add Joint Between the Left and the Vertical Section 1
                    JointHelper.CreateRigidJoint(rightPlate, "BotStructure", VERTSECTION1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);
                }
                else if (includeLeftPlate && !includeRigthPlate)
                {
                    (componentDictionary[VERTSECTION2]).SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[leftPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(leftPlate, "TopStructure", VERTSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);
                }
                else if (includeLeftPlate && includeRigthPlate)
                {
                    (componentDictionary[VERTSECTION1]).SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[rightPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    (componentDictionary[VERTSECTION2]).SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[leftPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(leftPlate, "BotStructure", VERTSECTION1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(rightPlate, "TopStructure", VERTSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
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

                    routeConnections.Add(new ConnectionInfo(HORSECTION, 1)); // partindex, routeindex

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




