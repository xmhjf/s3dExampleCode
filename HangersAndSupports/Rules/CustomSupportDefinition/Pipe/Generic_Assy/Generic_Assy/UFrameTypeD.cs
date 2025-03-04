//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   UFrameTypeD.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeD
//   Author       : Hema
//   Creation Date: 17-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-Sep-2013    Hema     CR-CP-224494 Converted Generic_Assy to C# .Net 
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
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

    public class UFrameTypeD : CustomSupportDefinition
    {
        private const string HORIZONTALSECTION = "HORIZONTALSECTION";
        private const string UBOLT = "UBOLT";
        private const string PLATE1 = "PLATE1";
        private const string PLATE2 = "PLATE2";

        string sectionSize;
        int sectionFromRule, sectionSizeValue, lengthFromRule;
        Double inset, hangerLength;
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
                    sectionFromRule = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionFromRule")).PropValue;
                    sectionSizeValue = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionSize")).PropValue;
                    inset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrGenAssyInset", "Inset")).PropValue;
                    lengthFromRule = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyHgrL", "LengthFromRule")).PropValue;
                    hangerLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyHgrL", "Length")).PropValue;

                    MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                    // Get the section size
                    if (sectionFromRule == 1)
                        GenericHelper.GetDataByRule("HgrSectionSize", support, out sectionSize);
                    else
                    {
                        sectionSize = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sectionSizeValue).DisplayName;
                        if (sectionSize.ToUpper().Equals("NONE") || sectionSize.Equals("0"))
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "UFrameTypeD.cs", 70);
                    }

                    // Get the UBolt
                    string uBoltPart = string.Empty;
                    GenericHelper.GetDataByRule("HgrUBoltSelection", support, out uBoltPart);

                    parts.Add(new PartInfo(HORIZONTALSECTION, sectionSize));
                    parts.Add(new PartInfo(UBOLT, uBoltPart));
                    parts.Add(new PartInfo(PLATE1, "Util_Plate_Metric_1"));
                    parts.Add(new PartInfo(PLATE2, "Util_Plate_Metric_1"));

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
        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                //Auto Dimensioning of Supports
                GenericAssemblyServices.CreateDimensionNote(this, "Dim1", componentDictionary[UBOLT], "Route");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim2", componentDictionary[HORIZONTALSECTION], "BeginCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim3", componentDictionary[HORIZONTALSECTION], "EndCap");

                double sectionDepth, sectionWidth, sectionThickness;
                //Get Section Structure dimensions
                BusinessObject horizontalSection = componentDictionary[HORIZONTALSECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)horizontalSection.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                sectionDepth = crossSection.Depth;
                sectionWidth = crossSection.Width;
                try
                {
                    sectionThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                }
                catch
                {
                    sectionThickness = 0.0;
                }
                double L, H;
                L = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeD_Dim", "IJUAHgrGenSrvDBrkDim", "L", "IJUAHgrGenSrvDBrkDim", "LBarSize", sectionSize);
                H = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeD_Dim", "IJUAHgrGenSrvDBrkDim", "H", "IJUAHgrGenSrvDBrkDim", "LBarSize", sectionSize);

                //==================================
                //Set properties on part occurrences
                //==================================

                double plateThickness = (double)((PropertyValueDouble)componentDictionary[PLATE1].GetPropertyValue("IJOAHgrUtilMetricT", "T")).PropValue;

                string plate1Bom = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateThickness, UnitName.DISTANCE_MILLIMETER) + "Plate Steel," + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, H, UnitName.DISTANCE_MILLIMETER) + "X" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L, UnitName.DISTANCE_MILLIMETER);

                componentDictionary[PLATE1].SetPropertyValue(H, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[PLATE1].SetPropertyValue(L, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[PLATE1].SetPropertyValue(plate1Bom, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                componentDictionary[PLATE2].SetPropertyValue(H, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[PLATE2].SetPropertyValue(L, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[PLATE2].SetPropertyValue(plate1Bom, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                //For UBolt
                Part uBolt = (Part)componentDictionary[UBOLT].GetRelationship("madeFrom", "part").TargetObjects[0];
                String uBoltBom = uBolt.PartDescription;
                componentDictionary[UBOLT].SetPropertyValue(sectionThickness, "IJOAHgrGenUBoltFlgThick", "FlgThick");
                componentDictionary[UBOLT].SetPropertyValue(uBoltBom, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");

                //Get the UBolt Type
                //Based on the type, we need to change the vertical offset to include Pad for UBolt Type E
                string uBoltType;
                GenericHelper.GetDataByRule("HgrUBoltType", support, out uBoltType);

                //Get Route object
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = pipeInfo.OutsideDiameter, padThickness = 0;
                if (uBoltType.Equals("E"))
                {
                    padThickness = 0.005; //As per email from China, Pad thickness = 5 mm for UBolt Type E,
                    componentDictionary[UBOLT].SetPropertyValue(sectionWidth, "IJOAHgrGenUBoltPadL", "PadL");
                }
                //For Hanger Beam
                if (lengthFromRule == 1)
                {
                    Collection<object> hangerLengthObject = new Collection<object>();
                    GenericHelper.GetDataByRule("HgrLength", support, out hangerLengthObject);
                    if (hangerLengthObject[0] == null)
                        hangerLength = Convert.ToDouble(hangerLengthObject[1]);
                    else
                        hangerLength = Convert.ToDouble(hangerLengthObject[0]);
                }
                componentDictionary[HORIZONTALSECTION].SetPropertyValue(hangerLength, "IJUAHgrOccLength", "Length");

                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;

                int sectioSizeValue = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sectionSizeValue).Value;
                // PropertyValueCodelist secsiz = (PropertyValueCodelist)support.GetPropertyValue("", "");

                support.SetPropertyValue(sectioSizeValue, "IJOAHgrGenAssySection", "SectionSize");
                support.SetPropertyValue(hangerLength, "IJOAHgrGenAssyHgrL", "Length");

                //=============
                //Create Joints
                //=============
                if ((SupportHelper.SupportingObjects.Count == 0))
                    //Add Joint Between the Horizontal Beam and Route
                    JointHelper.CreateRigidJoint(UBOLT, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                else
                    //Add Joint Between the Horizontal Beam and Route
                    JointHelper.CreateRigidJoint(UBOLT, "Route", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);

                //Add Joint Between the Horizontal Section and U Bolt
                JointHelper.CreateRigidJoint(HORIZONTALSECTION, "Neutral", UBOLT, "Route", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX, -pipeDiameter / 2 - sectionDepth / 2 - padThickness, 0, 0);

                //Add Joint Between the Horizontal Section and Plate1
                JointHelper.CreateRigidJoint(HORIZONTALSECTION, "BeginCap", PLATE1, "TopStructure", Plane.ZX, Plane.NegativeYZ, Axis.Z, Axis.Z, sectionDepth / 2, -H / 2, inset);

                //Add Joint Between the Horizontal Section and Plate2
                JointHelper.CreateRigidJoint(HORIZONTALSECTION, "EndCap", PLATE2, "BotStructure", Plane.ZX, Plane.NegativeYZ, Axis.Z, Axis.Z, sectionDepth / 2, -H / 2, -inset);
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
                return 1;
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

                    routeConnections.Add(new ConnectionInfo(UBOLT, 1)); //partindex, routeindex

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

                    //We are not connecting to any structure so we have nothing to return

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
