//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   UFrameTypeI.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeI
//   Author       : Rajeswari
//   Creation Date: 20-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 20-Sep-2013  Rajeswari CR-CP-224494 Convert HS_Generic_Assy to C# .Net
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
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class UFrameTypeI : CustomSupportDefinition
    {
        private const string HORSECTION = "HORSECTION";
        private const string UBOLT = "UBOLT";
        private const string BRACKET1 = "BRACKET1";
        private const string BRACKET2 = "BRACKET2";

        string sectionSize = string.Empty;
        int sizeOfSection;
        MetadataManager metadataManager;
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

                    int sectionFromRule = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionFromRule")).PropValue;
                    sizeOfSection = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionSize")).PropValue;

                    // Get the Section Size
                    metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                    //Collection<object> dataRuleobject = new Collection<object>();
                    if (sectionFromRule == 1)
                        GenericHelper.GetDataByRule("HgrSectionSize", support, out sectionSize);
                    else
                    {
                        sectionSize = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sizeOfSection).DisplayName;
                        if (sectionSize.ToUpper() == "NONE" || string.IsNullOrEmpty(sectionSize))
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "UFrameTypeH.cs", 102);
                    }

                    // Get the UBolt
                    string uBoltPart = string.Empty;
                    GenericHelper.GetDataByRule("HgrUBoltSelection", support, out uBoltPart);

                    // Add the parts to the collection
                    parts.Add(new PartInfo(HORSECTION, sectionSize));
                    parts.Add(new PartInfo(UBOLT, uBoltPart));
                    parts.Add(new PartInfo(BRACKET1, "HgrBracketWithTwoProfile"));
                    parts.Add(new PartInfo(BRACKET2, "HgrBracketWithTwoProfile"));
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

                // Auto Dimensioning of Supports
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_1", (SupportComponent)componentDictionary[UBOLT], "Route");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_2 ", (SupportComponent)componentDictionary[HORSECTION], "BeginCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_3 ", (SupportComponent)componentDictionary[HORSECTION], "EndCap");

                // Get Section Structure dimensions
                double sectionWidth = 0.0, sectionDepth = 0.0, sectionThickness = 0.0;
                BusinessObject horSectionPart = componentDictionary[HORSECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                sectionWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                sectionDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                sectionThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = pipeInfo.OutsideDiameter + 2 * pipeInfo.InsulationThickness;

                // =======================================
                // Set Values of Part Occurance Attributes
                // =======================================
                // For UBolt
                Part part = (Part)componentDictionary[UBOLT].GetRelationship("madeFrom", "part").TargetObjects[0];
                string uBoltBOM = part.PartDescription;

                componentDictionary[UBOLT].SetPropertyValue(sectionThickness, "IJOAHgrGenUBoltFlgThick", "FlgThick");
                componentDictionary[UBOLT].SetPropertyValue(uBoltBOM, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");

                // Get the UBolt Type. Based on the type, we need to change the vertical offset to include Pad for UBolt Type E
                string uBoltType = string.Empty;
                GenericHelper.GetDataByRule("HgrUBoltType", support, out uBoltType);

                // Check Insulation
                double padThicknessValue = 0;
                PipeObjectInfo pipeinfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeODWoInsul = GenericAssemblyServices.GetPipeODwithoutInsulatation(pipeDiameter, pipeinfo);
                if (uBoltType == "E")
                {
                    padThicknessValue = 0.005;
                    componentDictionary[UBOLT].SetPropertyValue(sectionWidth, "IJOAHgrGenUBoltPadL", "PadL");
                }

                // For Brackets
                double L, W, T;
                L = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeI_Dim", "IJUAHgrGenSrvIBrkDim", "L", "IJUAHgrGenSrvIBrkDim", "LBarSize", sectionSize);
                W = GenericAssemblyServices.GetDataByCondition("GenServ_FrmTypeI_Dim", "IJUAHgrGenSrvIBrkDim", "W", "IJUAHgrGenSrvIBrkDim", "LBarSize", sectionSize);
                T = (double)((PropertyValueDouble)componentDictionary[BRACKET1].GetPropertyValue("IJUAHgrBracketOcc", "BracketThickness")).PropValue;

                componentDictionary[BRACKET1].SetPropertyValue(W, "IJUAHgrBracketOcc", "AngleWidth");
                componentDictionary[BRACKET1].SetPropertyValue(W, "IJUAHgrBracketOcc", "AngleHeight");
                componentDictionary[BRACKET1].SetPropertyValue(L, "IJUAHgrBracketOcc", "BottomHeight");
                componentDictionary[BRACKET1].SetPropertyValue(L, "IJUAHgrBracketOcc", "TopWidth");
                componentDictionary[BRACKET1].SetPropertyValue(0.0001, "IJUAHgrBracketOcc", "CornerRadius");
                componentDictionary[BRACKET1].SetPropertyValue(T, "IJUAHgrBracketOcc", "BracketThickness");

                componentDictionary[BRACKET2].SetPropertyValue(W, "IJUAHgrBracketOcc", "AngleWidth");
                componentDictionary[BRACKET2].SetPropertyValue(W, "IJUAHgrBracketOcc", "AngleHeight");
                componentDictionary[BRACKET2].SetPropertyValue(L, "IJUAHgrBracketOcc", "BottomHeight");
                componentDictionary[BRACKET2].SetPropertyValue(L, "IJUAHgrBracketOcc", "TopWidth");
                componentDictionary[BRACKET2].SetPropertyValue(0.0001, "IJUAHgrBracketOcc", "CornerRadius");
                componentDictionary[BRACKET2].SetPropertyValue(T, "IJUAHgrBracketOcc", "BracketThickness");

                // For Hgr Beam
                int lengthFromRule = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyHgrL", "LengthFromRule")).PropValue;
                Collection<object> hgrlenghObj = new Collection<object>();
                Double hgrLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyHgrL", "Length")).PropValue;
                if (lengthFromRule == 1)
                {
                    GenericHelper.GetDataByRule("HgrLength", support, out hgrlenghObj);
                    if (hgrlenghObj[0] == null)
                        hgrLength = Convert.ToDouble(hgrlenghObj[1]);
                    else
                        hgrLength = Convert.ToDouble(hgrlenghObj[0]);
                }
                componentDictionary[HORSECTION].SetPropertyValue(hgrLength, "IJUAHgrOccLength", "Length");

                // Set properties on Assembly
                sizeOfSection = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sectionSize).Value;
                support.SetPropertyValue(sizeOfSection, "IJOAHgrGenAssySection", "SectionSize");
                support.SetPropertyValue(hgrLength, "IJOAHgrGenAssyHgrL", "Length");

                // =======================================
                // Create Joints
                // =======================================
                double bracketOffset = 0.1;
                // =======================================
                // Do Something if Structure is selected
                // =======================================
                if (SupportHelper.SupportingObjects.Count == 0)
                    // Add Joint Between the Horizontal Beam and Route
                    JointHelper.CreateRigidJoint(UBOLT, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                else
                    // Add Joint Between the Horizontal Beam and Route
                    JointHelper.CreateRigidJoint(UBOLT, "Route", "-1", "Route", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);

                // Add Joint Between the Horizontal Beam and U Bolt
                JointHelper.CreateRigidJoint(HORSECTION, "Neutral", UBOLT, "Route", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX, -pipeODWoInsul / 2 - sectionDepth / 2 - padThicknessValue, 0, 0);
                // Add Joint Between the Horizontal Beam and Bracket 1
                JointHelper.CreateRigidJoint(HORSECTION, "BeginCap", BRACKET1, "HgrPort_1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, -bracketOffset, -bracketOffset, 0);
                // Add Joint Between the Horizontal Beam and Bracket 2
                JointHelper.CreateRigidJoint(HORSECTION, "EndCap", BRACKET2, "HgrPort_1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, -bracketOffset, -bracketOffset, T);
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
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(UBOLT, 1));

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
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    //structConnections.Add(new ConnectionInfo(VERTSECTION, 1));

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

