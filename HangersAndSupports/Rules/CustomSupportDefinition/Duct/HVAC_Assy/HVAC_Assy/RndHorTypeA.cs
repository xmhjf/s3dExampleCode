//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RndHorTypeA.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RndHorTypeA
//   Author       :Manikanth
//   Creation Date:13-06-2014
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   13-06-2013     manikanth  CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//   27.04.2015     PVK        TR-CP-253033 Elevation CP not shown by default
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

    public class RndHorTypeA : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RndHorTypeA"
        //----------------------------------------------------------------------------------
        //Constants
        private const string CLAMP1 = "CLAMP1";
        private const string CLAMP2 = "CLAMP2";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";
        string leftPlate = "LEFTPLATE";
        string rightPlate = "RIGHTPLATE";
        string leftClampPlate = "LEFTCLAMPPLATE", rightClampPlate = "RIGHTCLAMPPLATE";

        string[] boltPartKeys = new string[8];

        string sectionSize, flatBarDim, angleOption, boltSize, sizeOfSection, sizeOfBolt;
        double plateThickness, clampLegLength, supAngle, flatBarThickness, flatBarWidth;
        int boltBegin, boltEnd;
        bool sectionFromRule, includeLeftPlate, includeRightPlate, showBolts, boltSizeFromRule, value;
        PropertyValueCodelist angleOptionCodeList;
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
                    supAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssySectAngle", "SectionAngle")).PropValue;
                    flatBarDim = (string)((PropertyValueString)support.GetPropertyValue("IJOAHgrHVACAssyFlatBarDim", "FlatBarDim")).PropValue;
                    clampLegLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyClamp", "LegLength")).PropValue;
                    showBolts = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "ShowBolts")).PropValue;
                    boltSizeFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSizeFromRule")).PropValue;
                    PropertyValueCodelist boltCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSize");
                    boltSize = boltCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltCodeList.PropValue).DisplayName;
                    angleOptionCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyAngleOpt", "AngleOpt");
                    angleOption = angleOptionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltCodeList.PropValue).DisplayName;

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
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RndHorTypeA.cs", 98);
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
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Bolt size is not available.", "", "RndHorTypeA.cs", 112);
                            return null;
                        }
                    }

                    parts.Add(new PartInfo(CLAMP1, "HgrHVACClamp1_1"));
                    parts.Add(new PartInfo(CLAMP2, "HgrHVACClamp1_1"));
                    parts.Add(new PartInfo(VERTSECTION1, "HgrHVACGenericBrace_L2_1"));
                    parts.Add(new PartInfo(VERTSECTION2, "HgrHVACGenericBrace_L2_1"));

                    int partCount;
                    if (includeLeftPlate)
                    {
                        parts.Add(new PartInfo(leftPlate, "Util_Plate_Metric_1"));
                        parts.Add(new PartInfo(leftClampPlate, "HgrHVACClampPlate_1"));
                    }
                    if (includeRightPlate)
                    {
                        parts.Add(new PartInfo(rightPlate, "Util_Plate_Metric_1"));
                        parts.Add(new PartInfo(rightClampPlate, "HgrHVACClampPlate_1"));
                    }
                    if (showBolts)
                    {
                        partCount = 0;
                        boltBegin = parts.Count + 1;
                        boltEnd = boltBegin + 1;
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
                    CmnException e1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
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

                Catalog m_oPlantCatalog = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog;
                BusinessObject sectionPart = m_oPlantCatalog.GetNamedObject(sectionSize);

                CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double sectionWidth = crossSection.Width;
                double sectionDepth = crossSection.Depth;
                double sectionThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                string supportingType;
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

                DuctObjectInfo ductInfo = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double ductRadius = ductInfo.OutsideDiameter / 2;

                //Auto Dimensioning of Supports
                Note note1 = CreateNote("Dim 1", (componentDictionary[VERTSECTION1]), "EndCap");
                note1.SetPropertyValue("", "IJGeneralNote", "Text");
                PropertyValueCodelist note1PropertyValueCL = (PropertyValueCodelist)note1.GetPropertyValue("IJGeneralNote", "Purpose");
                CodelistItem codeList1 = note1PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                note1.SetPropertyValue(codeList1, "IJGeneralNote", "Purpose");

                Note note2 = CreateNote("Dim 2", (componentDictionary[CLAMP1]), "StartOther");
                note2.SetPropertyValue("", "IJGeneralNote", "Text");
                PropertyValueCodelist note2PropertyValueCL = (PropertyValueCodelist)note2.GetPropertyValue("IJGeneralNote", "Purpose");
                CodelistItem codeList2 = note2PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                note2.SetPropertyValue(codeList2, "IJGeneralNote", "Purpose");

                //Gasket Note
                string gasketNote;
                string gasketBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACGasketDesc", "GasketDesc")).PropValue;
                if (string.IsNullOrEmpty(gasketBOMDesc))
                    gasketNote = "To line with 3mm thick Gasket all around for all contacted surface between Duct & Support";
                else
                    gasketNote = gasketBOMDesc;

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
                    note5.SetPropertyValue("", "IJGeneralNote", "Text");
                }
                else
                    DeleteNoteIfExists("Note 1");
                //=======================================
                //Do Something if more than one Structure
                //=======================================
                Boolean[] bIsOffsetApplied = HVACAssemblyServices.GetIsLugEndOffsetApplied(this);
                string[] idxStructPort = HVACAssemblyServices.GetIndexedStructPortName(this, bIsOffsetApplied);
                string leftStructPort = idxStructPort[0];
                string rightStructPort = idxStructPort[1];

                //Get the route structure distance
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                double routeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);

                string structPort1, structPort2;
                bool b = (Configuration == 1);
                if (b)
                {
                    structPort1 = leftStructPort;
                    structPort2 = rightStructPort;
                }
                else
                {
                    structPort1 = rightStructPort;
                    structPort2 = leftStructPort;
                }

                double platePlOffset1 = 0, plateAxoffset1 = 0, platePloffset2 = 0, plateAxoffset2 = 0, distRouteStruct1 = 0, distRouteStruct2 = 0;
                double vertSec1Length = 0, vertSec2Length = 0, vertSec1PLOffset = 0, vertSec1AxOffset = 0, vertSec2PlOffset = 0, vertSec2AxOffset = 0, vertSecOrgOffset = 0, leftPlateThk = 0, rightPlateThk = 0;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if ((Math.Abs(routeAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && Math.Abs(routeStructAngle) > (Math.PI - 0.001)))
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is valid only when the Route is horizontal. For vertical route, use Vertical Duct Support.", "", "RndHorTypeA.cs", 260);
                        return;
                    }
                    else
                    {
                        if ((((Math.Abs(routeStructAngle) < (0 + 0.001)) && (Math.Abs(routeStructAngle) > (0 - 0.001))) || ((Math.Abs(routeStructAngle) < (Math.PI + 0.001) && (Math.Abs(routeStructAngle) > (Math.PI - 0.001))))))
                        {
                            distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "Route", PortDistanceType.Horizontal);
                            distRouteStruct2 = distRouteStruct1;
                        }
                        else
                        {
                            distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "Route", PortDistanceType.Vertical);
                            distRouteStruct2 = distRouteStruct1;
                        }
                    }
                }
                else
                {
                    if ((Math.Abs(routeAngle) < (0 + 0.001) && (Math.Abs(routeAngle) > (0 - 0.001))) || ((Math.Abs(routeAngle) < (Math.PI + 0.001)) && (Math.Abs(routeAngle) > (Math.PI - 0.001))))
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support is valid only when the Route is horizontal. For vertical route, use Vertical Duct Support.", "", "RndHorTypeA.cs", 281);
                        return;
                    }
                    else
                    {
                        if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && (Math.Abs(routeStructAngle) > (Math.PI - 0.001))))
                        {
                            if (leftStructPort != rightStructPort)
                            {
                                distRouteStruct1 = RefPortHelper.DistanceBetweenPorts(leftStructPort, "Route", PortDistanceType.Horizontal);
                                distRouteStruct2 = RefPortHelper.DistanceBetweenPorts(rightStructPort, "Route", PortDistanceType.Horizontal);
                            }
                            else
                            {
                                distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "Route", PortDistanceType.Horizontal);
                                distRouteStruct2 = distRouteStruct1;
                            }
                        }
                        else
                        {
                            if (leftStructPort != rightStructPort)
                            {
                                if (supportingType.Equals("Slab-Steel"))
                                {
                                    distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "Route", PortDistanceType.Vertical);
                                    distRouteStruct2 = RefPortHelper.DistanceBetweenPorts("Struct_2", "Route", PortDistanceType.Vertical);
                                }
                                else if (supportingType.Equals("Steel-Slab"))
                                {
                                    distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Struct_2", "Route", PortDistanceType.Vertical);
                                    distRouteStruct2 = RefPortHelper.DistanceBetweenPorts("Structure", "Route", PortDistanceType.Vertical);
                                }
                                else if (supportingType.Equals("Steel"))
                                {
                                    distRouteStruct1 = RefPortHelper.DistanceBetweenPorts(leftStructPort, "Route", PortDistanceType.Vertical);
                                    distRouteStruct2 = RefPortHelper.DistanceBetweenPorts(rightStructPort, "Route", PortDistanceType.Vertical);
                                }
                            }
                            else
                            {
                                distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "Route", PortDistanceType.Vertical);
                                distRouteStruct2 = distRouteStruct1;
                            }
                        }
                    }
                }
                if (!includeLeftPlate && includeRightPlate)
                {
                    rightPlateThk = plateThickness;
                    leftPlateThk = 0;
                }
                else if (includeLeftPlate && !includeRightPlate)
                {
                    rightPlateThk = 0;
                    leftPlateThk = plateThickness;
                }
                else if (includeLeftPlate && includeRightPlate)
                {
                    rightPlateThk = plateThickness;
                    leftPlateThk = plateThickness;
                }
                else if (!includeLeftPlate && !includeRightPlate)
                {
                    rightPlateThk = 0;
                    leftPlateThk = 0;
                }
                Matrix4X4 routePortOrientation = new Matrix4X4(), structPortOrientation = new Matrix4X4(), struct2PortOrientation = new Matrix4X4();
                double angle1 = 0, angle2 = 0;
                if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                {
                    //Get the Ref Port Origin                 
                    routePortOrientation = RefPortHelper.PortLCS("BBR_High");
                    structPortOrientation = RefPortHelper.PortLCS("Structure");
                    struct2PortOrientation.SetIdentity();
                    if (SupportHelper.SupportingObjects.Count > 1)
                        struct2PortOrientation = RefPortHelper.PortLCS("Struct_2");

                    Vector temp1Vector = new Vector(), temp2Vector = new Vector(), normal1Vector = new Vector(), normal2Vector = new Vector(), temp3Vector = new Vector(), temp4Vector = new Vector();
                    temp3Vector = new Vector(structPortOrientation.ZAxis.X, structPortOrientation.ZAxis.Y, structPortOrientation.ZAxis.Z);
                    temp4Vector = new Vector(struct2PortOrientation.ZAxis.X, struct2PortOrientation.ZAxis.Y, struct2PortOrientation.ZAxis.Z);
                    normal1Vector = temp3Vector.GetOrthogonalVector();
                    normal2Vector = temp4Vector.GetOrthogonalVector();
                    temp1Vector = new Vector(routePortOrientation.Origin.X - structPortOrientation.Origin.X, routePortOrientation.Origin.Y - structPortOrientation.Origin.Y, routePortOrientation.Origin.Z - structPortOrientation.Origin.Z);
                    temp2Vector = new Vector(routePortOrientation.Origin.X - struct2PortOrientation.Origin.X, routePortOrientation.Origin.Y - struct2PortOrientation.Origin.Y, routePortOrientation.Origin.Z - struct2PortOrientation.Origin.Z);

                    angle1 = temp1Vector.Angle(temp3Vector, normal1Vector);
                    angle2 = temp2Vector.Angle(temp4Vector, normal2Vector);
                }

                double supAngle1 = 0, supAngle2 = 0;

                support.SetPropertyValue(supAngle, "IJOAHgrHVACAssySectAngle", "SectionAngle");

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    supAngle1 = supAngle2 = supAngle / 2;
                else
                {
                    if (supportingType.Equals("Slab"))
                    {
                        supAngle1 = supAngle2 = supAngle / 2;
                        PropertyValueCodelist angleCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyAngleOpt", "AngleOpt");
                        angleCodelist.PropValue = 2;
                        support.SetPropertyValue(angleCodelist.PropValue, "IJOAHgrHVACAssyAngleOpt", "AngleOpt");
                    }
                    else
                    {
                        if (supportingType.Equals("Slab-Steel"))
                        {
                            //User Defined
                            if (angleOptionCodeList.PropValue == 2)
                                supAngle1 = supAngle2 = supAngle / 2;
                            //Adjust to Supporting Steel
                            else if (angleOptionCodeList.PropValue == 1)
                            {
                                if (Configuration == 1)
                                {
                                    supAngle1 = supAngle / 2;
                                    supAngle2 = angle2;
                                }
                                else if (Configuration == 2)
                                {
                                    supAngle1 = supAngle / 2;
                                    supAngle2 = angle2;
                                }
                            }
                        }
                        else if (supportingType.Equals("Steel-Slab"))
                        {
                            if (angleOptionCodeList.PropValue == 2)
                                supAngle1 = supAngle2 = supAngle / 2;
                            else if (angleOptionCodeList.PropValue == 1)
                            {
                                if (Configuration == 1)
                                {
                                    supAngle1 = -angle1;
                                    supAngle2 = -supAngle / 2;
                                }
                                else if (Configuration == 2)
                                {
                                    supAngle1 = -supAngle1 / 2;
                                    supAngle2 = -angle1;
                                }
                            }
                        }
                        else if (supportingType.Equals("Steel"))
                        {
                            if (angleOptionCodeList.PropValue == 2)
                                supAngle1 = supAngle2 = supAngle / 2;
                            else if (angleOptionCodeList.PropValue == 1)
                            {
                                if (Configuration == 1)
                                {
                                    supAngle1 = -angle1;
                                    supAngle2 = angle2;
                                }
                                else if (Configuration == 2)
                                {
                                    supAngle1 = angle2;
                                    supAngle2 = -angle1;
                                }
                            }
                        }
                    }
                }
                vertSecOrgOffset = flatBarWidth / 2;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    platePlOffset1 = distRouteStruct1 - rightPlateThk;
                    plateAxoffset1 = (platePlOffset1) * Math.Tan(supAngle1);
                    platePloffset2 = distRouteStruct2 - leftPlateThk;
                    plateAxoffset2 = (platePloffset2) * Math.Tan(supAngle2);

                    vertSec1Length = (distRouteStruct1 / Math.Cos(supAngle1)) - ((ductRadius + flatBarThickness + rightPlateThk + rightPlateThk / Math.Cos(supAngle1))) + (sectionWidth / 2 * Math.Sin(supAngle1));
                    vertSec2Length = (distRouteStruct2 / Math.Cos(supAngle2)) - ((ductRadius + flatBarThickness + leftPlateThk + leftPlateThk / Math.Cos(supAngle2))) + (sectionWidth / 2 * Math.Sin(supAngle2));

                    vertSec1PLOffset = (ductRadius + flatBarThickness + rightPlateThk) * Math.Cos(supAngle1) - sectionWidth / 2 * Math.Sin(supAngle1);
                    vertSec1AxOffset = (ductRadius + flatBarThickness + rightPlateThk) * Math.Sin(supAngle1) + sectionWidth / 2 * Math.Cos(supAngle1);
                    vertSec2PlOffset = (ductRadius + flatBarThickness + leftPlateThk) * Math.Cos(supAngle2) - sectionWidth / 2 * Math.Sin(supAngle2);
                    vertSec2AxOffset = (ductRadius + flatBarThickness + leftPlateThk) * Math.Sin(supAngle2) + sectionWidth / 2 * Math.Cos(supAngle2);
                }
                else
                {
                    if (Configuration == 1)
                    {
                        platePlOffset1 = distRouteStruct1 - rightPlateThk;
                        plateAxoffset1 = (platePlOffset1) * Math.Tan(supAngle1);
                        platePloffset2 = distRouteStruct2 - leftPlateThk;
                        plateAxoffset2 = (platePloffset2) * Math.Tan(supAngle2);

                        vertSec1Length = distRouteStruct1 / Math.Cos(supAngle1) - (ductRadius + flatBarThickness + rightPlateThk + rightPlateThk / Math.Cos(supAngle1)) + sectionWidth / 2 * Math.Sin(supAngle1);
                        vertSec2Length = distRouteStruct2 / Math.Cos(supAngle2) - (ductRadius + flatBarThickness + leftPlateThk + leftPlateThk / Math.Cos(supAngle2)) + sectionWidth / 2 * Math.Sin(supAngle2);

                        vertSec1PLOffset = (ductRadius + flatBarThickness + rightPlateThk) * Math.Cos(supAngle1) - sectionWidth / 2 * Math.Sin(supAngle1);
                        vertSec1AxOffset = (ductRadius + flatBarThickness + rightPlateThk) * Math.Sin(supAngle1) + sectionWidth / 2 * Math.Cos(supAngle1);
                        vertSec2PlOffset = (ductRadius + flatBarThickness + leftPlateThk) * Math.Cos(supAngle2) - sectionWidth / 2 * Math.Sin(supAngle2);
                        vertSec2AxOffset = (ductRadius + flatBarThickness + leftPlateThk) * Math.Sin(supAngle2) + sectionWidth / 2 * Math.Cos(supAngle2);
                    }
                    else if (Configuration == 2)
                    {
                        platePlOffset1 = distRouteStruct2 - rightPlateThk;
                        plateAxoffset1 = (platePlOffset1) * Math.Tan(supAngle2);
                        platePloffset2 = (distRouteStruct2 - leftPlateThk);
                        plateAxoffset2 = (platePloffset2) * Math.Tan(supAngle1);

                        vertSec1Length = distRouteStruct2 / Math.Cos(supAngle1) - (ductRadius + flatBarThickness + rightPlateThk + rightPlateThk / Math.Cos(supAngle1)) + sectionWidth / 2 * Math.Sin(supAngle1);
                        vertSec2Length = distRouteStruct1 / Math.Cos(supAngle2) - (ductRadius + flatBarThickness + leftPlateThk + leftPlateThk / Math.Cos(supAngle2)) + sectionWidth / 2 + Math.Sin(supAngle2);

                        vertSec1PLOffset = (ductRadius + flatBarThickness + rightPlateThk) * Math.Cos(supAngle1) - sectionWidth / 2 * Math.Sin(supAngle1);
                        vertSec1AxOffset = (ductRadius + flatBarThickness + rightPlateThk) * Math.Sin(supAngle1) + sectionWidth / 2 * Math.Cos(supAngle1);
                        vertSec2PlOffset = (ductRadius + flatBarThickness + leftPlateThk) * Math.Cos(supAngle2) - sectionWidth / 2 * Math.Sin(supAngle2);
                        vertSec2AxOffset = (ductRadius + flatBarThickness + leftPlateThk) * Math.Sin(supAngle2) + sectionWidth / 2 * Math.Cos(supAngle2);

                    }
                }

                (componentDictionary[VERTSECTION1]).SetPropertyValue(vertSec1Length, "IJOAHgrHVACGenBrace", "L");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(vertSec2Length, "IJOAHgrHVACGenBrace", "L");

                string boltDesc = string.Empty;
                string boltBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACBoltsDesc", "BoltDesc")).PropValue;
                if (boltBOMDesc == null)
                    boltDesc = boltSize + " " + "Bolts";
                else
                    boltDesc = boltBOMDesc;
                string clamp1BOMDesc, clamp2BOMDesc;

                //========================================
                // Set Values of Part Occurance Attributes
                //========================================
                clamp1BOMDesc = "Clamp 1 for " + 2 * ductRadius + "Duct OutsideRadius";
                clamp2BOMDesc = "Clamp 2 for " + 2 * ductRadius + "Duct OutsideRadius";

                (componentDictionary[CLAMP1]).SetPropertyValue(flatBarWidth, "IJUAHgrHVACClamp", "Width");
                (componentDictionary[CLAMP1]).SetPropertyValue(flatBarThickness, "IJUAHgrHVACClampThk", "Thickness");
                (componentDictionary[CLAMP1]).SetPropertyValue(clampLegLength, "IJUAHgrHVACClamp", "LegLength");
                (componentDictionary[CLAMP1]).SetPropertyValue(0.0001, "IJUAHgrOccLength", "Length");
                (componentDictionary[CLAMP1]).SetPropertyValue(ductRadius, "IJUAHgrHVACClamp", "Radius");
                (componentDictionary[CLAMP1]).SetPropertyValue(clamp1BOMDesc, "IJOAHgrHVACBomDesc", "InputBomDesc1");

                (componentDictionary[CLAMP2]).SetPropertyValue(flatBarWidth, "IJUAHgrHVACClamp", "Width");
                (componentDictionary[CLAMP2]).SetPropertyValue(flatBarThickness, "IJUAHgrHVACClampThk", "Thickness");
                (componentDictionary[CLAMP2]).SetPropertyValue(clampLegLength, "IJUAHgrHVACClamp", "LegLength");
                (componentDictionary[CLAMP2]).SetPropertyValue(0.0001, "IJUAHgrOccLength", "Length");
                (componentDictionary[CLAMP2]).SetPropertyValue(ductRadius, "IJUAHgrHVACClamp", "Radius");
                (componentDictionary[CLAMP2]).SetPropertyValue(clamp2BOMDesc, "IJOAHgrHVACBomDesc", "InputBomDesc1");

                if (showBolts)
                {
                    for (int boltIndex = boltBegin; boltIndex <= boltEnd; boltIndex++)
                    {
                        componentDictionary[boltPartKeys[boltIndex - boltBegin]].SetPropertyValue(4 * flatBarThickness, "IJOAHgrUtilMetricL", "L");
                        componentDictionary[boltPartKeys[boltIndex - boltBegin]].SetPropertyValue(clampLegLength / 8, "IJOAHgrUtilMetricRadius", "Radius");
                        componentDictionary[boltPartKeys[boltIndex - boltBegin]].SetPropertyValue(boltBOMDesc, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                    }
                }

                (componentDictionary[VERTSECTION1]).SetPropertyValue(sectionWidth, "IJOAHgrHVACGenBrace", "W");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(sectionDepth, "IJOAHgrHVACGenBrace", "Depth");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(sectionThickness, "IJOAHgrHVACGenBrace", "WebT");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(sectionThickness, "IJOAHgrHVACGenBrace", "FlangeT");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(Math.PI / 2, "IJOAHgrHVACGenBrace", "CutBackAngle1");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(-Math.PI / 2 - supAngle1, "IJOAHgrHVACGenBrace", "CutBackAngle2");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(supAngle1, "IJOAHgrHVACGenBrace", "Angle");
                PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJOAHgrHVACGenBrace", "BraceOrient");
                braceCodelist.PropValue = 1;
                (componentDictionary[VERTSECTION1]).SetPropertyValue(braceCodelist.PropValue, "IJOAHgrHVACGenBrace", "BraceOrient");

                (componentDictionary[VERTSECTION2]).SetPropertyValue(sectionWidth, "IJOAHgrHVACGenBrace", "W");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(sectionDepth, "IJOAHgrHVACGenBrace", "Depth");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(sectionThickness, "IJOAHgrHVACGenBrace", "WebT");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(sectionThickness, "IJOAHgrHVACGenBrace", "FlangeT");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(Math.PI / 2 - supAngle2, "IJOAHgrHVACGenBrace", "CutBackAngle1");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(Math.PI / 2, "IJOAHgrHVACGenBrace", "CutBackAngle2");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(supAngle2, "IJOAHgrHVACGenBrace", "Angle");
                PropertyValueCodelist braceCodelist1 = (PropertyValueCodelist)componentDictionary[VERTSECTION2].GetPropertyValue("IJOAHgrHVACGenBrace", "BraceOrient");
                braceCodelist1.PropValue = 1;
                (componentDictionary[VERTSECTION2]).SetPropertyValue(braceCodelist.PropValue, "IJOAHgrHVACGenBrace", "BraceOrient");
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

                Plane planeVertSec1Route1, planeVertSec1Route2, planeVertSec2Route1, planeVertSec2Route2;
                planeVertSec1Route1 = planeVertSec1Route2 = planeVertSec2Route1 = planeVertSec2Route2 = Plane.XY;
                Axis axisVertSec1Route1, axisVertSec1Route2, axisVertSec2Route1, axisVertSec2Route2;
                axisVertSec1Route1 = axisVertSec1Route2 = axisVertSec2Route1 = axisVertSec2Route2 = Axis.X;
                if (Configuration == 1)
                {
                    planeVertSec1Route1 = Plane.YZ;
                    planeVertSec1Route2 = Plane.XY;
                    axisVertSec1Route1 = Axis.Y;
                    axisVertSec1Route2 = Axis.NegativeX;
                    planeVertSec2Route1 = Plane.YZ;
                    planeVertSec2Route2 = Plane.NegativeXY;
                    axisVertSec2Route1 = Axis.Y;
                    axisVertSec2Route2 = Axis.NegativeX;
                }
                else if (Configuration == 2)
                {
                    planeVertSec1Route1 = Plane.YZ;
                    planeVertSec1Route2 = Plane.XY;
                    axisVertSec1Route1 = Axis.Y;
                    axisVertSec1Route2 = Axis.X;
                    planeVertSec2Route1 = Plane.YZ;
                    planeVertSec2Route2 = Plane.NegativeXY;
                    axisVertSec2Route1 = Axis.Y;
                    axisVertSec2Route2 = Axis.X;
                }
                //Add Joint Between the UBolt and Route
                JointHelper.CreateRigidJoint(CLAMP1, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                //Add Joint Between the UBolt and Route
                JointHelper.CreateRigidJoint(CLAMP2, "Route", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(VERTSECTION1, "BeginCap", "-1", "Route", planeVertSec1Route1, planeVertSec1Route2, axisVertSec1Route1, axisVertSec1Route2, -vertSec1PLOffset, vertSec1AxOffset, vertSecOrgOffset);
                //Add joint between Horizontal Section 2 and Vertical Section 2
                JointHelper.CreateRigidJoint(VERTSECTION2, "EndCap", "-1", "Route", planeVertSec2Route1, planeVertSec2Route2, axisVertSec2Route1, axisVertSec2Route2, vertSec2PlOffset, vertSec2AxOffset, vertSecOrgOffset);
                if (showBolts)
                {
                    //Add Joint Between the UBolt and Route
                    JointHelper.CreateRigidJoint(CLAMP1, "StartOther", boltPartKeys[0], "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -(3 * flatBarThickness / 2), 0, 0);
                    //Add Joint Between the UBolt and Route
                    JointHelper.CreateRigidJoint(CLAMP1, "EndOther", boltPartKeys[1], "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -(3 * flatBarThickness / 2), 0, 0);
                }
                if (!includeLeftPlate && includeRightPlate)
                {
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[rightPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");


                    (componentDictionary[rightClampPlate]).SetPropertyValue(2 * sectionWidth, "IJUAHgrHVACClampPlate", "Width");
                    (componentDictionary[rightClampPlate]).SetPropertyValue(2 * sectionDepth, "IJUAHgrHVACClampPlate", "Depth");
                    (componentDictionary[rightClampPlate]).SetPropertyValue(plateThickness, "IJUAHgrHVACClampPlate", "Thickness");
                    (componentDictionary[rightClampPlate]).SetPropertyValue(ductRadius + flatBarThickness, "IJUAHgrHVACClampPlate", "Radius");
                    (componentDictionary[rightClampPlate]).SetPropertyValue(-supAngle1, "IJUAHgrHVACClampPlate", "Angle");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(rightPlate, "TopStructure", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, -platePloffset2, -plateAxoffset2, 0);
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(rightClampPlate, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                else if (includeLeftPlate && !includeRightPlate)
                {
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[leftPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    (componentDictionary[leftClampPlate]).SetPropertyValue(2 * sectionWidth, "IJUAHgrHVACClampPlate", "Width");
                    (componentDictionary[leftClampPlate]).SetPropertyValue(2 * sectionDepth, "IJUAHgrHVACClampPlate", "Depth");
                    (componentDictionary[leftClampPlate]).SetPropertyValue(plateThickness, "IJUAHgrHVACClampPlate", "Thickness");
                    (componentDictionary[leftClampPlate]).SetPropertyValue(ductRadius + flatBarThickness, "IJUAHgrHVACClampPlate", "Radius");
                    (componentDictionary[leftClampPlate]).SetPropertyValue(supAngle2, "IJUAHgrHVACClampPlate:", "Angle");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(leftPlate, "TopStructure", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, -platePloffset2, plateAxoffset2, 0);
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(leftClampPlate, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                if (includeLeftPlate && includeRightPlate)
                {
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[rightPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[rightPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    (componentDictionary[rightClampPlate]).SetPropertyValue(2 * sectionWidth, "IJUAHgrHVACClampPlate", "Width");
                    (componentDictionary[rightClampPlate]).SetPropertyValue(2 * sectionDepth, "IJUAHgrHVACClampPlate", "Depth");
                    (componentDictionary[rightClampPlate]).SetPropertyValue(plateThickness, "IJUAHgrHVACClampPlate", "Thickness");
                    (componentDictionary[rightClampPlate]).SetPropertyValue(ductRadius + flatBarThickness, "IJUAHgrHVACClampPlate", "Radius");
                    (componentDictionary[rightClampPlate]).SetPropertyValue(-supAngle1, "IJUAHgrHVACClampPlate", "Angle");

                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[leftPlate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[leftPlate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    (componentDictionary[leftClampPlate]).SetPropertyValue(2 * sectionWidth, "IJUAHgrHVACClampPlate", "Width");
                    (componentDictionary[leftClampPlate]).SetPropertyValue(2 * sectionDepth, "IJUAHgrHVACClampPlate", "Depth");
                    (componentDictionary[leftClampPlate]).SetPropertyValue(plateThickness, "IJUAHgrHVACClampPlate", "Thickness");
                    (componentDictionary[leftClampPlate]).SetPropertyValue(ductRadius + flatBarThickness, "IJUAHgrHVACClampPlate", "Radius");
                    (componentDictionary[leftClampPlate]).SetPropertyValue(supAngle2, "IJUAHgrHVACClampPlate", "Angle");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(leftPlate, "TopStructure", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, -platePloffset2, plateAxoffset2, 0);
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(leftClampPlate, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(rightPlate, "TopStructure", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, -platePlOffset1, -plateAxoffset1, 0);
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(rightClampPlate, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
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

                    routeConnections.Add(new ConnectionInfo(CLAMP1, 1)); // partindex, routeindex

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





