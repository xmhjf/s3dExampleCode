//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RndVert.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RndVert
//   Author       : Rajeswari
//   Creation Date: 17-June-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 17-June-2013  Rajeswari  CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//  27.Apr.2015     PVK     TR-CP-253033 Elevation CP not shown by default
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
    public class RndVert : CustomSupportDefinition
    {
        private const string HORSECTION = "HORSECTION";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";

        private const string HORZ_PARALLEL = "HORZ_PARALLEL";
        private const string HORZ_PERP = "HORZ_PERP";
        string[] idxClamp = new string[6];
        string[] boltPartKeys = new string[8];
        
        double[] ductRadius = new double[6];

        double plateThickness, overhangLeft, overhangRight, braceAngle, clampLegLength, flatBarThickness, flatBarWidth;
        string sectionSize, flatBarDim, boltSize, sizeOfSection = string.Empty;
        string leftBrace, rightBrace, leftConnectionObject, rightConnectionObject, leftPlate, rightPlate, leftBrPlate, rightBrPlate, sizeOfBolt;
        int iNumRoutes, idxClampBegin = 4, idxClampEnd, idxRoute, boltSizeValue, boltsBegin, boltsEnd, dxClamp;
        bool sectionFromRule, includeLCantiPlate, includeLBrPlate, includeRCantiPlate, includeRBrPlate, showBrace, showBolts, boltSizeFromRule, value;

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

                    // Get the attributes from assembly
                    sectionFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionFromRule")).PropValue;
                    PropertyValueCodelist sectionCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionSize");
                    sectionSize = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionCodeList.PropValue).DisplayName;
                    includeLCantiPlate = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyLCantiPlate", "IncludeLCantiPlate")).PropValue;
                    includeLBrPlate = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyLBrPlate", "IncludeLBrPlate")).PropValue;
                    includeRCantiPlate = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyRCantiPlate", "IncludeRCantiPlate")).PropValue;
                    includeRBrPlate = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyRBrPlate", "IncludeRBrPlate")).PropValue;
                    plateThickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyPlateThk", "PlateThick")).PropValue;
                    showBrace = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBrace", "ShowBrace")).PropValue;
                    braceAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyBrace", "BraceAngle")).PropValue;
                    flatBarDim = (string)((PropertyValueString)support.GetPropertyValue("IJOAHgrHVACAssyFlatBarDim", "FlatBarDim")).PropValue;
                    clampLegLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyClamp", "LegLength")).PropValue;
                    showBolts = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "ShowBolts")).PropValue;
                    boltSizeFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSizeFromRule")).PropValue;
                    boltSizeValue = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSize")).PropValue;
                    overhangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyOHLeft", "OverhangLeft")).PropValue;
                    overhangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyOHRight", "OverhangRight")).PropValue;
                    PropertyValueCodelist boltCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSize");
                    boltSize = boltCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltCodeList.PropValue).DisplayName;

                    string[] position = flatBarDim.Split('x');
                    flatBarWidth = double.Parse(position[0]) / 1000;
                    flatBarThickness = double.Parse(position[1]) / 1000;
                    if (sectionFromRule)
                        value = GenericHelper.GetDataByRule("HVACHgrSectionSize", (BusinessObject)support, out sizeOfSection);
                    else if (!sectionFromRule)
                    {
                        PropertyValueCodelist sectionSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionSize");
                        sizeOfSection = sectionSizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodeList.PropValue).DisplayName;

                        if (sizeOfSection.ToUpper().Equals("NONE") || sizeOfSection.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RndVert.cs", 96);
                            return null;
                        }
                    }
                    iNumRoutes = SupportHelper.SupportedObjects.Count;
                    for (idxRoute = 1; idxRoute <= iNumRoutes; idxRoute++)
                    {
                        DuctObjectInfo ductInfo = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(idxRoute);
                        ductRadius[idxRoute] = ductInfo.OutsideDiameter / 2;
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
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Bolt size is not available.", "", "RndVert.cs", 116);
                            return null;
                        }
                    }
                    //Create the list of parts
                    parts.Add(new PartInfo(HORSECTION, sizeOfSection));
                    parts.Add(new PartInfo(VERTSECTION1, sizeOfSection));
                    parts.Add(new PartInfo(VERTSECTION2, sizeOfSection));

                    int partIndex;
                    idxClampEnd = idxClampBegin - 1 + iNumRoutes;

                    for (dxClamp = idxClampBegin; dxClamp <= idxClampEnd; dxClamp++)
                    {
                        idxClamp[dxClamp] = "Clamp" + (dxClamp).ToString();
                        parts.Add(new PartInfo(idxClamp[dxClamp], "HgrHVACClamp1_1"));
                    }
                    if (showBrace)
                    {
                        partIndex = parts.Count;
                        leftBrace = (partIndex + 1).ToString();
                        rightBrace = (partIndex + 2).ToString();
                        leftConnectionObject = (partIndex + 3).ToString();
                        rightConnectionObject = (partIndex + 4).ToString();
                        parts.Add(new PartInfo(leftBrace, "HgrHVACGenericBrace_L1_1")); // sSectionSize
                        parts.Add(new PartInfo(rightBrace, "HgrHVACGenericBrace_L1_1")); // sSectionSize
                        parts.Add(new PartInfo(leftConnectionObject, "Log_Conn_Part_1")); // Rotational Connection Object
                        parts.Add(new PartInfo(rightConnectionObject, "Log_Conn_Part_1")); // Rotational Connection Object
                        if (includeLBrPlate)
                        {
                            partIndex = parts.Count;
                            leftBrPlate = (partIndex + 1).ToString();
                            parts.Add(new PartInfo(leftBrPlate, "Util_Plate_Metric_1"));
                        }
                        if (includeRBrPlate)
                        {
                            partIndex = parts.Count;
                            rightBrPlate = (partIndex + 1).ToString();
                            parts.Add(new PartInfo(rightBrPlate, "Util_Plate_Metric_1"));
                        }
                    }
                    if (includeLCantiPlate)
                    {
                        partIndex = parts.Count;
                        leftPlate = (partIndex + 1).ToString();
                        parts.Add(new PartInfo(leftPlate, "Util_Plate_Metric_1"));
                    }
                    if (includeRCantiPlate)
                    {
                        partIndex = parts.Count;
                        rightPlate = (partIndex + 1).ToString();
                        parts.Add(new PartInfo(rightPlate, "Util_Plate_Metric_1"));
                    }
                    if (showBolts)
                    {
                        partIndex = 0;
                        boltsBegin = parts.Count + 1;
                        boltsEnd = Convert.ToInt32(boltsBegin + iNumRoutes * 2 - 1);
                        for (int index = boltsBegin; index <= boltsEnd; index++)
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

                // Get Section Structure dimensions
                BusinessObject horizontalSectionPart = (componentDictionary[HORSECTION]).GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //double sectionDepth = crosssection.Depth;
                //double sectionWidth = crosssection.Width;
                double sectionWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double sectionDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                double sectionThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                // Gasket Note
                string gasketNote = string.Empty;
                string gasketBomDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACGasketDesc", "GasketDesc")).PropValue;
                if (string.IsNullOrEmpty(gasketNote))
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
                    Note note = CreateNote("Note 1");
                    PropertyValueCodelist notePropertyValueCL = (PropertyValueCodelist)note.GetPropertyValue("IJGeneralNote", "Purpose");
                    CodelistItem codeList = notePropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                    note.SetPropertyValue(gasketNote, "IJGeneralNote", "Text");
                    note.SetPropertyValue(codeList, "IJGeneralNote", "Purpose");
                }
                else
                    DeleteNoteIfExists("Note 1");

                // get structure count
                Boolean[] bIsOffsetApplied = HVACAssemblyServices.GetIsLugEndOffsetApplied(this);
                string[] idxStructPort = HVACAssemblyServices.GetIndexedStructPortName(this, bIsOffsetApplied);
                string leftStructPort = idxStructPort[0];
                string rightStructPort = idxStructPort[1];

                // Check to see what they are connecting to
                double ductAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);

                if (Math.Abs(ductAngle) < (Math.PI / 2 + 0.001) && Math.Abs(ductAngle) > (Math.PI / 2 - 0.001))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is valid only when the Route is vertical and the supporting Structure is besides the Route.", "", "RndVert.cs", 280);
                    return;
                }

                double distRouteStruct = 0;
                if (ductAngle < (0 + 0.0001) && ductAngle > (0 - 0.0001) || ductAngle < (Math.PI + 0.0001) && ductAngle > (Math.PI - 0.0001))
                    distRouteStruct = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                else if (ductAngle < (Math.PI / 2 + 0.0001) && ductAngle > (Math.PI / 2 - 0.0001))
                    distRouteStruct = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                // ========================================
                // Set Values of Part Occurance Attributes
                // ========================================
                PropertyValueCodelist horizontalbeginMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (horizontalbeginMiterCodelist.PropValue == -1)
                    horizontalbeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist horizontalendMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (horizontalendMiterCodelist.PropValue == -1)
                    horizontalendMiterCodelist.PropValue = 1;
                PropertyValueCodelist verticalSection1beginMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION1]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (verticalSection1beginMiterCodelist.PropValue == -1)
                    verticalSection1beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist verticalSection1endMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION1]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (verticalSection1endMiterCodelist.PropValue == -1)
                    verticalSection1endMiterCodelist.PropValue = 1;
                PropertyValueCodelist verticalSection2beginMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION2]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (verticalSection2beginMiterCodelist.PropValue == -1)
                    verticalSection2beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist verticalSection2endMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION2]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (verticalSection2endMiterCodelist.PropValue == -1)
                    verticalSection2endMiterCodelist.PropValue = 1;

                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[HORSECTION]).SetPropertyValue(horizontalbeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[HORSECTION]).SetPropertyValue(horizontalendMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                (componentDictionary[VERTSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(verticalSection1beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(verticalSection1endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                (componentDictionary[VERTSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(verticalSection2beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(verticalSection2endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                string boltDesc = string.Empty;
                string boltBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACBoltsDesc", "BoltDesc")).PropValue;
                if (string.IsNullOrEmpty(boltBOMDesc))
                    boltDesc = boltSize + " " + "Bolts";
                else
                    boltDesc = boltBOMDesc;

                if (showBolts)
                {
                    for (int boltIndex = boltsBegin; boltIndex <= boltsEnd; boltIndex++)
                    {
                        componentDictionary[boltPartKeys[boltIndex - boltsBegin]].SetPropertyValue(sectionDepth / 2, "IJOAHgrUtilMetricL", "L");
                        componentDictionary[boltPartKeys[boltIndex - boltsBegin]].SetPropertyValue(sectionWidth / 4, "IJOAHgrUtilMetricRadius", "Radius");
                        componentDictionary[boltPartKeys[boltIndex - boltsBegin]].SetPropertyValue(boltDesc, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
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
                iNumRoutes = SupportHelper.SupportedObjects.Count;
                for (dxClamp = idxClampBegin; dxClamp <= idxClampEnd; dxClamp++)
                {
                    if (dxClamp == idxClampBegin)
                        idxRoute = 1;
                    else
                        idxRoute = dxClamp - idxClampBegin + 1;
                    (componentDictionary[idxClamp[dxClamp]]).SetPropertyValue(clampLegLength, "IJUAHgrHVACClamp", "LegLength");
                    (componentDictionary[idxClamp[dxClamp]]).SetPropertyValue(flatBarWidth, "IJUAHgrHVACClamp", "Width");
                    (componentDictionary[idxClamp[dxClamp]]).SetPropertyValue(flatBarThickness, "IJUAHgrHVACClampThk", "Thickness");
                    (componentDictionary[idxClamp[dxClamp]]).SetPropertyValue(ductRadius[idxRoute], "IJUAHgrOccLength", "Length");
                    (componentDictionary[idxClamp[dxClamp]]).SetPropertyValue(ductRadius[idxRoute], "IJUAHgrHVACClamp", "Radius");
                }
                // Get the route structure distance
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);

                string structPort1, structPort2;
                double routeStruct1Dist, routeStruct2Dist, horizontalSectionLength;
                horizontalSectionLength = boundingBoxWidth + 2 * sectionWidth;
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
                double vertSec1AxisOffset = 0, vertSec2AxisOffset = 0, routeLowStruct1HorDist = 0, routeLowStruct2HorDist = 0, routeHighStruct1HorDist = 0, routeHighStruct2HorDist = 0, distBetStruct = 0;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if ((Math.Abs(ductAngle) < (0 + 0.001) && Math.Abs(ductAngle) > (0 - 0.001)) || (Math.Abs(ductAngle) < (Math.PI + 0.001) && Math.Abs(ductAngle) > (Math.PI - 0.001)))
                        routeStruct1Dist = RefPortHelper.DistanceBetweenPorts("BBSR_High", "Structure", PortDistanceType.Horizontal_Perpendicular);
                    else
                    {
                        if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && Math.Abs(routeStructAngle) > (Math.PI - 0.001)))
                            routeStruct1Dist = RefPortHelper.DistanceBetweenPorts("BBSR_High", "Structure", PortDistanceType.Horizontal);
                        else
                            routeStruct1Dist = RefPortHelper.DistanceBetweenPorts("BBSR_High", "Structure", PortDistanceType.Vertical);
                    }
                    routeStruct2Dist = routeStruct1Dist;
                    (componentDictionary[VERTSECTION1]).SetPropertyValue(routeStruct1Dist, "IJUAHgrOccLength", "Length");
                    (componentDictionary[VERTSECTION2]).SetPropertyValue(routeStruct2Dist, "IJUAHgrOccLength", "Length");
                    (componentDictionary[HORSECTION]).SetPropertyValue(horizontalSectionLength, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    if ((Math.Abs(ductAngle) < (0 + 0.001) && Math.Abs(ductAngle) > (0 - 0.001)) || (Math.Abs(ductAngle) < (Math.PI + 0.001) && Math.Abs(ductAngle) > (Math.PI - 0.001)))
                    {
                        routeStruct1Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", structPort1, PortDistanceType.Horizontal_Perpendicular);
                        routeStruct2Dist = RefPortHelper.DistanceBetweenPorts("BBR_High", structPort2, PortDistanceType.Horizontal_Perpendicular);
                        routeLowStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", rightStructPort, PortDistanceType.Horizontal_Parallel);
                        routeLowStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_Low", leftStructPort, PortDistanceType.Horizontal_Parallel);
                        routeHighStruct1HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", rightStructPort, PortDistanceType.Horizontal_Parallel);
                        routeHighStruct2HorDist = RefPortHelper.DistanceBetweenPorts("BBR_High", leftStructPort, PortDistanceType.Horizontal_Parallel);
                        distBetStruct = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortDistanceType.Horizontal);
                    }
                    else
                    {
                        if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && Math.Abs(routeStructAngle) > (Math.PI - 0.001)))
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
                    }
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
                    if (supportingType.ToUpper().Equals("SLAB"))
                        horizontalSectionLength = boundingBoxWidth;
                    else if (supportingType.ToUpper().Equals("STEEL"))
                    {
                        horizontalSectionLength = boundingBoxWidth;
                        if (routeLowStruct2HorDist < routeHighStruct2HorDist)
                            overhangRight = routeLowStruct2HorDist;
                        else
                            overhangRight = routeHighStruct2HorDist;
                        if (routeLowStruct1HorDist < routeHighStruct1HorDist)
                            overhangLeft = routeLowStruct1HorDist;
                        else
                            overhangLeft = routeHighStruct1HorDist;
                    }
                    else if (supportingType.ToUpper().Equals("STEEL-SLAB") || supportingType.ToUpper().Equals("SLAB-STEEL"))
                    {
                        if (routeLowStruct2HorDist < routeHighStruct2HorDist)
                            overhangRight = routeLowStruct2HorDist;
                        else
                            overhangRight = routeHighStruct2HorDist;
                        if (routeLowStruct1HorDist < routeHighStruct1HorDist)
                            overhangLeft = routeLowStruct1HorDist;
                        else
                            overhangLeft = routeHighStruct1HorDist;

                        horizontalSectionLength = boundingBoxWidth;
                    }
                    (componentDictionary[VERTSECTION1]).SetPropertyValue(routeStruct1Dist, "IJUAHgrOccLength", "Length");
                    (componentDictionary[VERTSECTION2]).SetPropertyValue(routeStruct2Dist, "IJUAHgrOccLength", "Length");
                    (componentDictionary[HORSECTION]).SetPropertyValue(horizontalSectionLength, "IJUAHgrOccLength", "Length");
                }

                double horSec1Length = 0, horSec2Length = 0, brace1Length = 0, brace2Length = 0, brPadOffset1 = 0, brPadOffset2 = 0;
                string leftBraceBOM, rightBraceBOM;
                if (showBrace)
                {
                    horSec1Length = routeStruct1Dist - plateThickness;
                    horSec2Length = routeStruct2Dist - plateThickness;
                    brace1Length = horSec1Length / Math.Cos(braceAngle);
                    brace2Length = horSec2Length / Math.Cos(braceAngle);
                    brPadOffset1 = (brace1Length) * Math.Sin(braceAngle) - sectionDepth;
                    brPadOffset2 = (brace2Length) * Math.Sin(braceAngle) - sectionDepth;
                  
                    leftBraceBOM = "Generic Brace" + sizeOfSection + " Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, brace1Length, UnitName.DISTANCE_MILLIMETER);
                    rightBraceBOM = "Generic Brace" + sizeOfSection + " Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, brace2Length, UnitName.DISTANCE_MILLIMETER);

                    (componentDictionary[leftBrace]).SetPropertyValue(brace1Length, "IJOAHgrHVACGenBrace", "L");
                    (componentDictionary[leftBrace]).SetPropertyValue(Math.PI / 2 - braceAngle, "IJOAHgrHVACGenBrace", "Angle");
                    PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[leftBrace].GetPropertyValue("IJOAHgrHVACGenBrace", "BraceOrient");
                    braceCodelist.PropValue = 2;
                    (componentDictionary[leftBrace]).SetPropertyValue(braceCodelist.PropValue, "IJOAHgrHVACGenBrace", "BraceOrient");
                    (componentDictionary[leftBrace]).SetPropertyValue(sectionWidth, "IJOAHgrHVACGenBrace", "W");
                    (componentDictionary[leftBrace]).SetPropertyValue(sectionDepth, "IJOAHgrHVACGenBrace", "Depth");
                    (componentDictionary[leftBrace]).SetPropertyValue(sectionThickness, "IJOAHgrHVACGenBrace", "FlangeT");
                    (componentDictionary[leftBrace]).SetPropertyValue(Math.PI / 2 - braceAngle, "IJOAHgrHVACGenBrace", "CutBackAngle1");
                    (componentDictionary[leftBrace]).SetPropertyValue(braceAngle, "IJOAHgrHVACGenBrace", "CutBackAngle2");
                    (componentDictionary[leftBrace]).SetPropertyValue(leftBraceBOM, "IJOAHgrHVACGenBrace", "InputBomDesc");

                    (componentDictionary[rightBrace]).SetPropertyValue(brace1Length, "IJOAHgrHVACGenBrace", "L");
                    (componentDictionary[rightBrace]).SetPropertyValue(Math.PI / 2 - braceAngle, "IJOAHgrHVACGenBrace", "Angle");
                    PropertyValueCodelist braceCodelist1 = (PropertyValueCodelist)componentDictionary[rightBrace].GetPropertyValue("IJOAHgrHVACGenBrace", "BraceOrient");
                    braceCodelist1.PropValue = 2;
                    (componentDictionary[rightBrace]).SetPropertyValue(braceCodelist1.PropValue, "IJOAHgrHVACGenBrace", "BraceOrient");
                    (componentDictionary[rightBrace]).SetPropertyValue(sectionWidth, "IJOAHgrHVACGenBrace", "W");
                    (componentDictionary[rightBrace]).SetPropertyValue(sectionDepth, "IJOAHgrHVACGenBrace", "Depth");
                    (componentDictionary[rightBrace]).SetPropertyValue(sectionThickness, "IJOAHgrHVACGenBrace", "FlangeT");
                    (componentDictionary[rightBrace]).SetPropertyValue(braceAngle, "IJOAHgrHVACGenBrace", "CutBackAngle1");
                    (componentDictionary[rightBrace]).SetPropertyValue(Math.PI / 2 - braceAngle, "IJOAHgrHVACGenBrace", "CutBackAngle2");
                    (componentDictionary[rightBrace]).SetPropertyValue(rightBraceBOM, "IJOAHgrHVACGenBrace", "InputBomDesc");
                }

                Plane[] horSecRouteIndexPlane = new Plane[2]; Plane[] vertSec1BraceIndexPlane = new Plane[2]; Plane[] vertSec2BraceIndexPlane = new Plane[2];
                Axis[] horSecRouteIndexAxis = new Axis[2]; Axis[] vertSec1BraceIndexAxis = new Axis[2]; Axis[] vertSec2BraceIndexAxis = new Axis[2];
                string routePort1 = string.Empty, routePort2 = string.Empty, horSecPort = string.Empty;

                if (Configuration == 1)
                {
                    horSecRouteIndexPlane[0] = Plane.YZ;
                    horSecRouteIndexPlane[1] = Plane.XY;
                    horSecRouteIndexAxis[0] = Axis.Y;
                    horSecRouteIndexAxis[1] = Axis.NegativeX;
                    vertSec1BraceIndexPlane[0] = Plane.ZX;
                    vertSec1BraceIndexPlane[1] = Plane.NegativeXY;
                    vertSec1BraceIndexAxis[0] = Axis.X;
                    vertSec1BraceIndexAxis[1] = Axis.X;
                    vertSec2BraceIndexPlane[0] = Plane.ZX;
                    vertSec2BraceIndexPlane[1] = Plane.XY;
                    vertSec2BraceIndexAxis[1] = Axis.X;
                    vertSec2BraceIndexAxis[1] = Axis.NegativeX;

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
                    vertSec1AxisOffset = -overhangRight;
                    vertSec2AxisOffset = overhangLeft;
                    horSecPort = "BeginCap";

                    (componentDictionary[HORSECTION]).SetPropertyValue(overhangRight, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[HORSECTION]).SetPropertyValue(overhangLeft, "IJUAHgrOccOverLength", "BeginOverLength");
                }
                else if (Configuration == 2)
                {
                    horSecRouteIndexPlane[0] = Plane.YZ;
                    horSecRouteIndexPlane[1] = Plane.XY;
                    horSecRouteIndexAxis[0] = Axis.Y;
                    horSecRouteIndexAxis[1] = Axis.X;
                    vertSec1BraceIndexPlane[0] = Plane.ZX;
                    vertSec1BraceIndexPlane[1] = Plane.XY;
                    vertSec1BraceIndexAxis[0] = Axis.X;
                    vertSec1BraceIndexAxis[1] = Axis.NegativeX;
                    vertSec2BraceIndexPlane[0] = Plane.ZX;
                    vertSec2BraceIndexPlane[1] = Plane.NegativeXY;
                    vertSec2BraceIndexAxis[1] = Axis.X;
                    vertSec2BraceIndexAxis[1] = Axis.X;

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
                    horSecPort = "EndCap";
                    vertSec1AxisOffset = -overhangLeft;
                    vertSec2AxisOffset = overhangRight;

                    (componentDictionary[HORSECTION]).SetPropertyValue(overhangLeft, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[HORSECTION]).SetPropertyValue(overhangRight, "IJUAHgrOccOverLength", "BeginOverLength");
                }

                // =============
                // Create Joints
                // =============
                string refPortName = string.Empty;
                int bolts = 0;
                for (dxClamp = idxClampBegin; dxClamp <= idxClampEnd; dxClamp++)
                {
                    if (dxClamp == idxClampBegin)
                        refPortName = "Route";
                    else
                        refPortName = "Route_" + (dxClamp - idxClampBegin + 1).ToString();

                    JointHelper.CreateRigidJoint(idxClamp[dxClamp], "Route", "-1", refPortName, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    if (showBolts)
                    {
                        JointHelper.CreateRigidJoint(idxClamp[dxClamp], "StartOther", boltPartKeys[bolts], "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -(flatBarThickness + sectionThickness), 0, 0);
                        bolts++;
                        JointHelper.CreateRigidJoint(idxClamp[dxClamp], "EndOther", boltPartKeys[bolts], "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -(flatBarThickness + sectionThickness), 0, 0);
                        bolts++;
                    }
                }
                // Add joint between Horizontal Section 2 and Route
                JointHelper.CreateRigidJoint(HORSECTION, horSecPort, "-1", routePort2, horSecRouteIndexPlane[0], horSecRouteIndexPlane[1], horSecRouteIndexAxis[0], horSecRouteIndexAxis[1], 0.0, 0.0, sectionDepth / 2);
                if (showBrace)
                {
                    //Add a Joint between horizontal section and the point that we want the angle to rotate around.
                    JointHelper.CreateRigidJoint(leftConnectionObject, "Connection", VERTSECTION1, "BeginCap", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X, sectionThickness, -sectionThickness, -sectionThickness);

                    JointHelper.CreateRigidJoint(leftConnectionObject, "Connection", leftBrace, "EndCap", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0, 0);

                    //Add a Joint between horizontal section and the point that we want the angle to rotate around.
                    JointHelper.CreateRigidJoint(rightConnectionObject, "Connection", VERTSECTION2, "EndCap", Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX, sectionThickness, -sectionThickness, sectionThickness);

                    JointHelper.CreateRigidJoint(rightConnectionObject, "Connection", rightBrace, "BeginCap", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0, 0);
                    if (includeLBrPlate)
                    {
                        (componentDictionary[leftBrPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                        (componentDictionary[leftBrPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                        (componentDictionary[leftBrPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                        // Add Joint Between the Plate and the Horizontal Section
                        JointHelper.CreateRigidJoint(leftBrPlate, "BotStructure", VERTSECTION1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -sectionWidth / 2, -brPadOffset1 - sectionDepth / 2);
                    }
                    if (includeRBrPlate)
                    {
                        (componentDictionary[rightBrPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                        (componentDictionary[rightBrPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                        (componentDictionary[rightBrPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                        // Add Joint Between the Plate and the Horizontal Section
                        JointHelper.CreateRigidJoint(rightBrPlate, "TopStructure", VERTSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, -sectionDepth / 2, -brPadOffset2 - sectionDepth / 2);
                    }
                }
                // Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(VERTSECTION1, "BeginCap", HORSECTION, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, 0, vertSec1AxisOffset, 0);

                // Add joint between Horizontal Section 2 and Vertical Section 2
                JointHelper.CreateRigidJoint(VERTSECTION2, "EndCap", HORSECTION, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0, -vertSec2AxisOffset, 0);

                if (!includeLCantiPlate && includeRCantiPlate)
                {
                    (componentDictionary[VERTSECTION1]).SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[rightPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    JointHelper.CreateRigidJoint(rightPlate, "BotStructure", VERTSECTION1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);
                }
                else if (includeLCantiPlate && !includeRCantiPlate)
                {
                    (componentDictionary[VERTSECTION2]).SetPropertyValue(plateThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[leftPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    // Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(leftPlate, "TopStructure", VERTSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);
                }
                else if (includeLCantiPlate && includeRCantiPlate)
                {
                    (componentDictionary[VERTSECTION1]).SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[VERTSECTION2]).SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[leftPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[rightPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    // Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(leftPlate, "BotStructure", VERTSECTION1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);

                    // Add Joint Between the Plate and the Vertical Beam
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
