//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   UFrameTypeH.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeH
//   Author       : Rajeswari
//   Creation Date: 20-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 20-Sep-2013  Rajeswari CR-CP-224494 Convert HS_Generic_Assy to C# .Net
// 22-Jan-2015    PVK     TR-CP-264951  Resolve coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Content.Support.Symbols;
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

    public class UFrameTypeH : CustomSupportDefinition
    {
        private const string HORSECTION1 = "HORSECTION1";
        private const string HORSECTION2 = "HORSECTION2";
        private const string HORSECTION3 = "HORSECTION3";
        private const string HORSECTION4 = "HORSECTION4";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";
        private const string VERTSECTION3 = "VERTSECTION3";
        private const string VERTSECTION4 = "VERTSECTION4";
        private const string UBOLT = "UBOLT";

        private const string LEFTPAD1 = "LEFTPAD1";
        private const string LEFTPAD2 = "LEFTPAD2";
        private const string RIGHTPAD1 = "RIGHTPAD1";
        private const string RIGHTPAD2 = "RIGHTPAD2";

        string sectionSize = string.Empty;
        int lengthFromRule, includeLeftPad, includeRightPad, sizeOfSection;
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

                    // Get the attributes from assembly
                    int sectionFromRule = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionFromRule")).PropValue;
                    sizeOfSection = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionSize")).PropValue;
                    includeLeftPad = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyLeftPad", "IncludeLeftPad")).PropValue;
                    includeRightPad = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyRightPad", "IncludeRightPad")).PropValue;
                    lengthFromRule = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyHgrL", "LengthFromRule")).PropValue;

                    if (includeLeftPad == 2)
                        GenericHelper.GetDataByRule("HgrIncludePad", support, out includeLeftPad);
                    if (includeRightPad == 2)
                        GenericHelper.GetDataByRule("HgrIncludePad", support, out includeRightPad);

                    metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                    sectionSize = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sizeOfSection).DisplayName;
                    // Get the Section Size
                    if (sectionFromRule == 1)
                        GenericHelper.GetDataByRule("HgrSectionSize", support, out sectionSize);
                    else
                    {
                        sectionSize = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sizeOfSection).DisplayName;
                        if (sectionSize.ToUpper() == "NONE" || string.IsNullOrEmpty(sectionSize))
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "UFrameTypeH.cs", 102);
                    }

                    // Get the nUBolt
                    string uBoltPart = string.Empty;
                    GenericHelper.GetDataByRule("HgrUBoltSelection", support, out uBoltPart);

                    // Get the Angle Pad
                    string anglePadPart = GenericAssemblyServices.GetDataByConditionString("GenServ_AnglePadDim", "IJUAHgrGenServAnglePad", "PadPart", "IJUAHgrGenServAnglePad", "SectionSize", sectionSize);

                    // Create the list of parts
                    parts.Add(new PartInfo(HORSECTION1, sectionSize));
                    parts.Add(new PartInfo(HORSECTION2, sectionSize));
                    parts.Add(new PartInfo(HORSECTION3, sectionSize));
                    parts.Add(new PartInfo(VERTSECTION1, sectionSize));
                    parts.Add(new PartInfo(VERTSECTION2, sectionSize));
                    parts.Add(new PartInfo(VERTSECTION3, sectionSize));
                    parts.Add(new PartInfo(VERTSECTION4, sectionSize));
                    parts.Add(new PartInfo(UBOLT, uBoltPart));

                    if (includeLeftPad == 1)
                    {
                        parts.Add(new PartInfo(LEFTPAD1, anglePadPart));
                        parts.Add(new PartInfo(LEFTPAD2, anglePadPart));
                    }
                    if (includeRightPad == 1)
                    {
                        parts.Add(new PartInfo(RIGHTPAD1, anglePadPart));
                        parts.Add(new PartInfo(RIGHTPAD2, anglePadPart));
                    }
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
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_2 ", (SupportComponent)componentDictionary[HORSECTION1], "BeginCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_3 ", (SupportComponent)componentDictionary[HORSECTION1], "EndCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_4 ", (SupportComponent)componentDictionary[HORSECTION2], "BeginCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_5 ", (SupportComponent)componentDictionary[HORSECTION2], "EndCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_6 ", (SupportComponent)componentDictionary[HORSECTION3], "BeginCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_7 ", (SupportComponent)componentDictionary[HORSECTION3], "EndCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_8 ", (SupportComponent)componentDictionary[VERTSECTION1], "EndCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_9 ", (SupportComponent)componentDictionary[VERTSECTION2], "EndCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_10 ", (SupportComponent)componentDictionary[VERTSECTION3], "BeginCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim_11 ", (SupportComponent)componentDictionary[VERTSECTION4], "BeginCap");

                // Get Section Structure dimensions
                double sectionWidth = 0.0, sectionDepth = 0.0, sectionThickness = 0.0;
                BusinessObject horSectionPart = componentDictionary[HORSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                sectionWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                sectionDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                sectionThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                // =======================================
                // Set Values of Part Occurance Attributes
                // =======================================
                PropertyValueCodelist beginMitercodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMitercodelist.PropValue == -1)
                    beginMitercodelist.PropValue = 1;
                PropertyValueCodelist endMitercodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMitercodelist.PropValue == -1)
                    endMitercodelist.PropValue = 1;
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERTSECTION1].SetPropertyValue(beginMitercodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERTSECTION1].SetPropertyValue(endMitercodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                beginMitercodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION2].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMitercodelist.PropValue == -1)
                    beginMitercodelist.PropValue = 1;
                endMitercodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION2].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMitercodelist.PropValue == -1)
                    endMitercodelist.PropValue = 1;
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERTSECTION2].SetPropertyValue(beginMitercodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERTSECTION2].SetPropertyValue(endMitercodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                beginMitercodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION3].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMitercodelist.PropValue == -1)
                    beginMitercodelist.PropValue = 1;
                endMitercodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION3].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMitercodelist.PropValue == -1)
                    endMitercodelist.PropValue = 1;
                componentDictionary[VERTSECTION3].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION3].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION3].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERTSECTION3].SetPropertyValue(beginMitercodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERTSECTION3].SetPropertyValue(endMitercodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                beginMitercodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION4].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMitercodelist.PropValue == -1)
                    beginMitercodelist.PropValue = 1;
                endMitercodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION4].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMitercodelist.PropValue == -1)
                    endMitercodelist.PropValue = 1;
                componentDictionary[VERTSECTION4].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION4].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION4].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERTSECTION4].SetPropertyValue(beginMitercodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERTSECTION4].SetPropertyValue(endMitercodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                Double hangerLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyHgrL", "Length")).PropValue;
                Double H = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyTypeHDim", "H")).PropValue;
                Double L1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyTypeHDim", "L1")).PropValue;
                Double L2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyTypeHDim", "L2")).PropValue;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = pipeInfo.OutsideDiameter + 2 * pipeInfo.InsulationThickness;
                if (L2 < L1 + sectionWidth + pipeDiameter / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "L2 dimension should be greater than L1.", "", "UFrameTypeH.cs", 229);

                componentDictionary[VERTSECTION1].SetPropertyValue(H, "IJUAHgrOccLength", "Length");
                componentDictionary[VERTSECTION2].SetPropertyValue(H, "IJUAHgrOccLength", "Length");
                componentDictionary[VERTSECTION3].SetPropertyValue(H, "IJUAHgrOccLength", "Length");
                componentDictionary[VERTSECTION4].SetPropertyValue(H, "IJUAHgrOccLength", "Length");
                
                // For Hanger beam which connected to UBolt
                if (lengthFromRule == 1)
                {
                    Collection<object> HgrLengthObj = new Collection<object>();
                    GenericHelper.GetDataByRule("HgrLength", support, out HgrLengthObj);
                    if (HgrLengthObj[0] == null)
                        hangerLength = Convert.ToDouble(HgrLengthObj[1]);
                    else
                        hangerLength = Convert.ToDouble(HgrLengthObj[0]);
                }
               
                componentDictionary[HORSECTION1].SetPropertyValue(2 * hangerLength, "IJUAHgrOccLength", "Length");
                
                // For UBolt
                Part part = (Part)componentDictionary[UBOLT].GetRelationship("madeFrom", "part").TargetObjects[0];
                string uBoltBOM = part.PartDescription;

                // Angle Pad
                string anglePadBOM = string.Empty;

                if (includeLeftPad == 1 && includeRightPad == 2)
                    part = (Part)componentDictionary[LEFTPAD1].GetRelationship("madeFrom", "part").TargetObjects[0];
                else if (includeRightPad == 1 && includeLeftPad == 2)
                    part = (Part)componentDictionary[RIGHTPAD1].GetRelationship("madeFrom", "part").TargetObjects[0];
                else if (includeLeftPad == 1 && includeRightPad == 1)
                    part = (Part)componentDictionary[LEFTPAD1].GetRelationship("madeFrom", "part").TargetObjects[0];

                if (includeLeftPad == 1 || includeRightPad == 1)
                    anglePadBOM = part.PartDescription;

                componentDictionary[UBOLT].SetPropertyValue(sectionThickness, "IJOAHgrGenUBoltFlgThick", "FlgThick");
                componentDictionary[UBOLT].SetPropertyValue(uBoltBOM, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");

                // Get the UBolt Type. Based on the type, we need to change the vertical offset to include Pad for UBolt Type E
                string uBoltType = string.Empty;
                GenericHelper.GetDataByRule("HgrUBoltType", support, out uBoltType);

                // Check Insulation
                double padThick = 0;
                PipeObjectInfo pipeinfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeODWoInsulation = GenericAssemblyServices.GetPipeODwithoutInsulatation(pipeDiameter, pipeinfo);
                if (uBoltType == "E")
                {
                    padThick = 0.005;
                    componentDictionary[UBOLT].SetPropertyValue(sectionWidth, "IJOAHgrGenUBoltPadL", "PadL");
                }
                componentDictionary[HORSECTION2].SetPropertyValue(L2 - pipeODWoInsulation / 2, "IJUAHgrOccLength", "Length");
                componentDictionary[HORSECTION3].SetPropertyValue(L2 - pipeODWoInsulation / 2, "IJUAHgrOccLength", "Length");

                // Set properties on Assembly
                sizeOfSection = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sectionSize).Value;
                support.SetPropertyValue(sizeOfSection, "IJOAHgrGenAssySection", "SectionSize");
                support.SetPropertyValue(hangerLength, "IJOAHgrGenAssyHgrL", "Length");

                double overhang = 0.01;
                // See which direction the pipe is going - Up/Down
                double pipeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                GenericAssemblyServices.ConfigIndex configIndex1 = new GenericAssemblyServices.ConfigIndex(), configIndex2 = new GenericAssemblyServices.ConfigIndex(), configIndex3 = new GenericAssemblyServices.ConfigIndex(), configIndex4 = new GenericAssemblyServices.ConfigIndex(), configIndex5 = new GenericAssemblyServices.ConfigIndex();
                string hangerPort1 = string.Empty, hangerPort2 = string.Empty, overLength1 = string.Empty, overLength2 = string.Empty, padPort1 = string.Empty, padPort2 = string.Empty;

                if (HgrCompareDoubleService.cmpdbl(pipeAngle , 0) == true)
                {
                    configIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.X);
                    configIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeYZ, Axis.Z, Axis.Y);
                    configIndex3 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeYZ, Axis.Z, Axis.NegativeY);
                    configIndex4 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.X);
                    configIndex5 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.X);

                    hangerPort1 = "BeginCap";
                    hangerPort2 = "EndCap";
                    overLength1 = "BeginOverLength";
                    overLength2 = "EndOverLength";
                    padPort1 = "HgrPort_1";
                    padPort2 = "HgrPort_2";
                }
                else if ((pipeAngle > Math.PI - 0.00001) && (pipeAngle < Math.PI + 0.00001))
                {
                    configIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX);
                    configIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.X);
                    configIndex3 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.NegativeX);
                    configIndex4 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X);
                    configIndex5 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);

                    hangerPort1 = "EndCap";
                    hangerPort2 = "BeginCap";
                    overLength1 = "EndOverLength";
                    overLength2 = "BeginOverLength";
                    padPort1 = "HgrPort_2";
                    padPort2 = "HgrPort_1";
                }
                else
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "This support is placed on vertical pipes only.", "", "UFrameTypeH.cs", 327);

                // =======================================
                // Create Joints
                // =======================================
                double padThickness = 0;
                if (includeLeftPad == 1 && includeRightPad == 2)
                {
                    padThickness = (double)((PropertyValueDouble)componentDictionary[LEFTPAD1].GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                    componentDictionary[LEFTPAD1].SetPropertyValue(anglePadBOM, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");
                    componentDictionary[LEFTPAD2].SetPropertyValue(anglePadBOM, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");
                    componentDictionary[VERTSECTION1].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", overLength2);
                    componentDictionary[VERTSECTION2].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", overLength2);

                    // Add Joint Between the Pad and the Vertical Beam 1
                    JointHelper.CreateRigidJoint(LEFTPAD1, padPort2, VERTSECTION1, hangerPort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                    // Add Joint Between the Pad and the Vertical Beam 1
                    JointHelper.CreateRigidJoint(LEFTPAD2, padPort2, VERTSECTION2, hangerPort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                }
                else if (includeLeftPad == 2 && includeRightPad == 1)
                {
                    padThickness = (double)((PropertyValueDouble)componentDictionary[RIGHTPAD1].GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                    componentDictionary[RIGHTPAD1].SetPropertyValue(anglePadBOM, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");
                    componentDictionary[RIGHTPAD2].SetPropertyValue(anglePadBOM, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");
                    componentDictionary[VERTSECTION3].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", overLength1);
                    componentDictionary[VERTSECTION4].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", overLength1);

                    // Add Joint Between the Pad and the Vertical Beam 2
                    JointHelper.CreateRigidJoint(RIGHTPAD1, padPort1, VERTSECTION3, hangerPort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                    // Add Joint Between the Pad and the Vertical Beam 2
                    JointHelper.CreateRigidJoint(RIGHTPAD2, padPort1, VERTSECTION4, hangerPort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                }
                else if (includeLeftPad == 1 && includeRightPad == 1)
                {
                    componentDictionary[LEFTPAD1].SetPropertyValue(anglePadBOM, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");
                    componentDictionary[LEFTPAD2].SetPropertyValue(anglePadBOM, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");
                    componentDictionary[RIGHTPAD1].SetPropertyValue(anglePadBOM, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");
                    componentDictionary[RIGHTPAD2].SetPropertyValue(anglePadBOM, "IJOAHgrGenPartBOMDesc", "BOM_DESC1");
                    componentDictionary[VERTSECTION1].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", overLength2);
                    componentDictionary[VERTSECTION2].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", overLength2);
                    componentDictionary[VERTSECTION3].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", overLength1);
                    componentDictionary[VERTSECTION4].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", overLength1);

                    // Add Joint Between the Pad and the Vertical Beam 1
                    JointHelper.CreateRigidJoint(LEFTPAD1, padPort2, VERTSECTION1, hangerPort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                    // Add Joint Between the Pad and the Vertical Beam 1
                    JointHelper.CreateRigidJoint(LEFTPAD2, padPort2, VERTSECTION2, hangerPort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);

                    // Add Joint Between the Pad and the Vertical Beam 2
                    JointHelper.CreateRigidJoint(RIGHTPAD1, padPort1, VERTSECTION3, hangerPort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                    // Add Joint Between the Pad and the Vertical Beam 2
                    JointHelper.CreateRigidJoint(RIGHTPAD2, padPort1, VERTSECTION4, hangerPort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                }

                // Add Joint Between the UBolt and Route
                JointHelper.CreateRigidJoint(UBOLT, "Route", "-1", "Route", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                // Add Joint Between the Horizontal Section and U Bolt '1189
                JointHelper.CreateRigidJoint(HORSECTION1, "Neutral", UBOLT, "Route", configIndex1.A, configIndex1.B, configIndex1.C, configIndex1.D, -pipeODWoInsulation / 2 - sectionDepth / 2 - padThick, 0, 0);
                // Add Joint Between the Horizontal Section 1 and Horizontal Section 2
                JointHelper.CreateRigidJoint(HORSECTION1, "BeginCap", HORSECTION2, "EndCap", configIndex2.A, configIndex2.B, configIndex2.C, configIndex2.D, 0, 0, overhang);
                // Add Joint Between the Horizontal Section 1 and Horizontal Section 3
                JointHelper.CreateRigidJoint(HORSECTION1, "EndCap", HORSECTION3, "BeginCap", configIndex3.A, configIndex3.B, configIndex3.C, configIndex3.D, 0, 0, -overhang);
                // Add Joint Between the Horizontal Section 2 and Vertical Section 1
                JointHelper.CreateRigidJoint(HORSECTION2, "BeginCap", VERTSECTION1, hangerPort1, configIndex4.A, configIndex4.B, configIndex4.C, configIndex4.D, 0, 0, 0);
                // Add Joint Between the Horizontal Section 3 and Vertical Section 2
                JointHelper.CreateRigidJoint(HORSECTION2, "BeginCap", VERTSECTION2, hangerPort1, configIndex4.A, configIndex4.B, configIndex4.C, configIndex4.D, 0, L1, sectionThickness);
                // Add Joint Between the Horizontal Section 2 and Vertical Section 1
                JointHelper.CreateRigidJoint(HORSECTION3, "EndCap", VERTSECTION3, hangerPort2, configIndex5.A, configIndex5.B, configIndex5.C, configIndex5.D, 0, 0, 0);
                // Add Joint Between the Horizontal Section 3 and Vertical Section 2
                JointHelper.CreateRigidJoint(HORSECTION3, "EndCap", VERTSECTION4, hangerPort2, configIndex5.A, configIndex5.B, configIndex5.C, configIndex5.D, 0, -L1, sectionThickness);
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
