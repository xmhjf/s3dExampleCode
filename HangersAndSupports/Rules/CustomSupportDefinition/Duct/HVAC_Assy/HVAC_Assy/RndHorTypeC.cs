//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RndHorTypeC.cs
//   HVAC_Assy,Ingr.SP3D.Content.Support.Rules.RndHorTypeC
//   Author       :Manikanth
//   Creation Date:13-06-2014
//   Description:


//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   13-06-2013     manikanth  CR-CP-224486 Convert HS_HVAC_Assy to C# .Net
//   27.04.2015     PVK        TR-CP-253033 Elevation CP not shown by default
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

    public class RndHorTypeC : CustomSupportDefinition
    {
        private const string VERTSECTION = "VERTSECTION";
        private const string CLAMP = "CLAMP";
        string bolt = "BOLT";
        string plate = "PLATE";

        string sectionSize, flatBarDim, sizeOfSection = string.Empty, boltSize, sizeOfBolt;
        double plateThickness, clampLegLength, flatBarThickness = 0, flatbarWidth = 0;
        bool sectionFromRule, includePlate, showBolts, boltSizeFromRule, value;

        string[] boltPartKeys = new string[8];

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    sectionFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionFromRule")).PropValue;
                    PropertyValueCodelist sectionCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionSize");
                    sectionSize = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionCodeList.PropValue).DisplayName;
                    includePlate = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyPlate", "IncludePlate")).PropValue;
                    plateThickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyPlateThk", "PlateThick")).PropValue;
                    flatBarDim = (string)((PropertyValueString)support.GetPropertyValue("IJOAHgrHVACAssyFlatBarDim", "FlatBarDim")).PropValue;
                    clampLegLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHVACAssyClamp", "LegLength")).PropValue;
                    showBolts = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "ShowBolts")).PropValue;
                    boltSizeFromRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSizeFromRule")).PropValue;
                    PropertyValueCodelist boltCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssyBolts", "BoltSize");
                    boltSize = boltCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(boltCodeList.PropValue).DisplayName;

                    string[] position = flatBarDim.Split('x');
                    flatbarWidth = double.Parse(position[0]) / 1000;
                    flatBarThickness = double.Parse(position[1]) / 1000;

                    if (sectionFromRule)
                        value = GenericHelper.GetDataByRule("HVACHgrSectionSize", (BusinessObject)support, out sizeOfSection);
                    else if (!sectionFromRule)
                    {
                        PropertyValueCodelist sectionSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrHVACAssySection", "SectionSize");
                        sizeOfSection = sectionSizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodeList.PropValue).DisplayName;

                        if (sizeOfSection.ToUpper().Equals("NONE") || sizeOfSection.Equals(string.Empty))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "RndHorTypeC.cs", 82);
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
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Bolt size is not available.", "", "RndHorTypeC.cs", 96);
                            return null;
                        }
                    }
                    parts.Add(new PartInfo(CLAMP, "HgrHVACClamp2_1"));
                    parts.Add(new PartInfo(VERTSECTION, sizeOfSection));
                    if (includePlate)
                        parts.Add(new PartInfo(plate, "Util_Plate_Metric_1"));

                    if (showBolts)
                        parts.Add(new PartInfo(bolt, "Util_Fixed_Cyl_Metric_1"));
                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in GetCatalogParts Method of Hvac_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                BusinessObject horizontalSectionPart = (componentDictionary[VERTSECTION]).GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double sectionDepth = crosssection.Depth;
                double sectionWidth = crosssection.Width;
                double steelThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

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
                    note5.SetPropertyValue(gasketNote, "IJGeneralNote", "Text");
                }
                else
                    DeleteNoteIfExists("Note 1");
                DuctObjectInfo ductInfo = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double ductRadius = ductInfo.OutsideDiameter / 2;

                Boolean[] bIsOffsetApplied = HVACAssemblyServices.GetIsLugEndOffsetApplied(this);
                string[] idxStructPort = HVACAssemblyServices.GetIndexedStructPortName(this, bIsOffsetApplied);
                string leftStructPort = idxStructPort[0];
                string rightStructPort = idxStructPort[1];

                PropertyValueCodelist hor1beginMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (hor1beginMiterCodelist.PropValue == -1)
                    hor1beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist hor1endMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (hor1endMiterCodelist.PropValue == -1)
                    hor1endMiterCodelist.PropValue = 1;

                (componentDictionary[VERTSECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[VERTSECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[VERTSECTION]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[VERTSECTION]).SetPropertyValue(hor1beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[VERTSECTION]).SetPropertyValue(hor1endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                BusinessObject clamp = (componentDictionary[CLAMP]).GetRelationship("madeFrom", "part").TargetObjects[0];
                double clampClearance = (double)((PropertyValueDouble)clamp.GetPropertyValue("IJUAHgrHVACClamp", "Clearance")).PropValue;

                string boltDesc;
                string boltBOMDesc = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrHVACBoltsDesc", "BoltDesc")).PropValue;
                if (boltBOMDesc == null)
                    boltDesc = boltSize + " " + "Bolts";
                else
                    boltDesc = boltBOMDesc;
                if (showBolts)
                {
                    componentDictionary[bolt].SetPropertyValue((flatBarThickness + steelThickness + 0.01), "IJOAHgrUtilMetricL", "L");
                    componentDictionary[bolt].SetPropertyValue(clampLegLength / 8, "IJOAHgrUtilMetricRadius", "Radius");
                    componentDictionary[bolt].SetPropertyValue(boltDesc, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
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
                string clampBomDesc = "Clamp for " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, 2 * ductRadius, UnitName.DISTANCE_MILLIMETER) + " Duct outside diameter";

                (componentDictionary[CLAMP]).SetPropertyValue(clampLegLength, "IJUAHgrHVACClamp", "LegLength");
                (componentDictionary[CLAMP]).SetPropertyValue(flatbarWidth, "IJUAHgrHVACClamp", "Width");
                (componentDictionary[CLAMP]).SetPropertyValue(flatBarThickness, "IJUAHgrHVACClampThk", "Thickness");
                (componentDictionary[CLAMP]).SetPropertyValue(ductRadius, "IJUAHgrHVACClamp", "Radius");
                (componentDictionary[CLAMP]).SetPropertyValue(clampBomDesc, "IJOAHgrHVACBomDesc", "InputBomDesc1");

                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                double routeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);

                double routeStructDistance;
                if ((Math.Abs(routeAngle) < (0 + 0.001) && Math.Abs(routeAngle) > (0 - 0.001)) || (Math.Abs(routeAngle) < (Math.PI + 0.001) && Math.Abs(routeAngle) > (Math.PI - 0.001)))
                    routeStructDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                else
                {
                    if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && Math.Abs(routeStructAngle) > (Math.PI - 0.001)))
                        routeStructDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                    else
                        routeStructDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                }
                (componentDictionary[VERTSECTION]).SetPropertyValue(routeStructDistance - ductRadius - flatBarThickness, "IJUAHgrOccLength", "Length");

                Plane plane1 = new Plane(); Plane plane2 = new Plane();
                Axis axis1 = new Axis(); Axis axis2 = new Axis();
                if (Configuration == 1)
                {
                    plane1 = Plane.XY;
                    plane2 = Plane.XY;
                    axis1 = Axis.X;
                    axis2 = Axis.X;
                }
                else if (Configuration == 2)
                {
                    plane1 = Plane.XY;
                    plane2 = Plane.XY;
                    axis1 = Axis.Y;
                    axis2 = Axis.X;
                }
                //Add Joint Between the UBolt and Route
                JointHelper.CreateRigidJoint(CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                //Add joint between Horizontal Section 1 and Vertical Section 1
                JointHelper.CreateRigidJoint(VERTSECTION, "BeginCap", CLAMP, "Structure", plane1, plane2, axis1, axis2, clampLegLength / 2, -clampClearance - flatBarThickness, sectionDepth / 2);
                if (showBolts)
                    //Add Joint Between the UBolt and Route
                    JointHelper.CreateRigidJoint(CLAMP, "Structure", bolt, "StartOther", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, (flatBarThickness + steelThickness + 2 * clampClearance), 0);

                if (includePlate)
                {
                    (componentDictionary[VERTSECTION]).SetPropertyValue(-plateThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[plate]).SetPropertyValue(2 * sectionWidth, "IJOAHgrUtilMetricWidth", "Width");
                    (componentDictionary[plate]).SetPropertyValue(2 * sectionDepth, "IJOAHgrUtilMetricDepth", "Depth");
                    (componentDictionary[plate]).SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(plate, "BotStructure", VERTSECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 2, sectionDepth / 2);
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

                    routeConnections.Add(new ConnectionInfo(CLAMP, 1)); // partindex, routeindex

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

                    structConnections.Add(new ConnectionInfo(VERTSECTION, 1)); // partindex, routeindex

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