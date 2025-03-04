//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   UFrameTypeG.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeG
//   Author       : Rajeswari
//   Creation Date: 20-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 20-Sep-2013  Rajeswari  CR-CP-224494 Convert HS_Generic_Assy to C# .Net
// 31-Oct-2014     PVK     TR-CP-260301	Resolve coverity issues found in August 2014 report
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//   06/06/2016     Vinay   TR-CP-296065	Fix new coverity issues found in H&S Content 
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

    public class UFrameTypeG : CustomSupportDefinition
    {
        private const string VERTSECTION = "VERTSECTION";
        private const string SECPAD = "SECPAD";
        private const string BRACE = "BRACE";
        private const string CONNECTIONOBJ = "CONNECTIONOBJ";
        private const string BRACEPAD = "BRACEPAD";

        string sectionSize;
        int showBrace, includePad, includeBracePad, sizeOfSection, angle, routeCount, uBoltBeginIndex, uBoltEndIndex, routeIndex, uBoltIndex;

        string[] uboltPartKeys = new string[0];
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

                    // Get Route object
                    routeCount = SupportHelper.SupportedObjects.Count;

                    // Get the attributes from assembly
                    showBrace = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrace", "ShowBrace")).PropValue;
                    int sectionFromRule = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionFromRule")).PropValue;
                    sizeOfSection = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionSize")).PropValue;
                    includePad = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyPad", "IncludePad")).PropValue;
                    includeBracePad = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrPad", "IncludeBracePad")).PropValue;
                    angle = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrace", "BraceAngle")).PropValue;
                    metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                    // Get the section size
                    if (sectionFromRule == 1)
                        GenericHelper.GetDataByRule("HgrSectionSize", support, out sectionSize);
                    else
                    {
                        sectionSize = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sizeOfSection).DisplayName;
                        if (sectionSize.ToUpper() == "NONE" || string.IsNullOrEmpty(sectionSize))
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "UFrameTypeG.cs", 80);
                    }

                    // Get the UBolt
                    string[] uBoltParts = new string[0];
                    Collection<object> uBoltPart = new Collection<object>();
                    GenericHelper.GetDataByRule("HgrUBoltSelection", support, out uBoltPart);
                    if (uBoltPart != null)
                    {
                        for (uBoltIndex = 1; uBoltIndex <= routeCount; uBoltIndex++)
                        {
                            Array.Resize(ref uBoltParts, uBoltIndex);

                            if (uBoltPart[0] == null)
                                uBoltParts[uBoltIndex - 1] = uBoltPart[uBoltIndex].ToString();
                            else
                                uBoltParts[uBoltIndex - 1] = uBoltPart[uBoltIndex - 1].ToString();
                        }
                    }
                    // Get the Angle Pad
                    string anglePadPart = GenericAssemblyServices.GetDataByConditionString("GenServ_AnglePadDim", "IJUAHgrGenServAnglePad", "PadPart", "IJUAHgrGenServAnglePad", "SectionSize", sectionSize);

                    // Create list of parts
                    parts.Add(new PartInfo(VERTSECTION, sectionSize));
                    uBoltBeginIndex = 2; uBoltEndIndex = routeCount + 1;
                    for (uBoltIndex = uBoltBeginIndex; uBoltIndex <= uBoltEndIndex; uBoltIndex++)
                    {
                        if (uBoltIndex == uBoltBeginIndex)
                            routeIndex = 1;
                        else
                            routeIndex = uBoltIndex - uBoltBeginIndex + 1;
                        Array.Resize(ref uboltPartKeys, uBoltIndex);
                        uboltPartKeys[uBoltIndex - 1] = "UBOLT" + uBoltIndex;
                        parts.Add(new PartInfo(uboltPartKeys[uBoltIndex - 1], uBoltParts[routeIndex - 1]));
                    }

                    if (includePad == 1)
                        parts.Add(new PartInfo(SECPAD, anglePadPart));

                    if (showBrace == 1)
                    {
                        parts.Add(new PartInfo(BRACE, "HgrGen_GenericBrace_1"));
                        parts.Add(new PartInfo(CONNECTIONOBJ, "Log_Conn_Part_1"));
                    }

                    if (includeBracePad == 1)
                    {
                        if (showBrace == 1)
                            parts.Add(new PartInfo(BRACEPAD, anglePadPart));
                        else if (showBrace == 2)
                        {
                            includeBracePad = 2;
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Brace is not available. Resetting the value.", "", "UFrameTypeG.cs", 124);
                        }
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
                int textIndex = 0;
                for (int index = uBoltBeginIndex; index <= uBoltEndIndex; index++)
                {
                    if (index == uBoltBeginIndex)
                        textIndex = 1;
                    else
                        textIndex = index - uBoltBeginIndex + 1;
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, (SupportComponent)componentDictionary[uboltPartKeys[index - 1]], "Route");
                }
                textIndex = routeCount + 1;
                GenericAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, (SupportComponent)componentDictionary[VERTSECTION], "BeginCap");
                textIndex = textIndex + 1;
                GenericAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, (SupportComponent)componentDictionary[VERTSECTION], "EndCap");

                if (showBrace == 1)
                {
                    textIndex = textIndex + 1;
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, (SupportComponent)componentDictionary[BRACE], "BeginCap");
                    textIndex = textIndex + 1;
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim " + textIndex, (SupportComponent)componentDictionary[BRACE], "EndCap");
                }

                // Get Section Structure dimensions
                double sectionWidth = 0.0, sectionDepth = 0.0, sectionThickness = 0.0;
                BusinessObject vertSectionPart = componentDictionary[VERTSECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)vertSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                sectionWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                sectionDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                sectionThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                // Get the largest Pipe Dia

                double[] pipeDiameter = new double[0];
                for (int i = 1; i <= routeCount; i++)
                {
                    Array.Resize(ref pipeDiameter, i);
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(i);
                    pipeDiameter[i - 1] = pipeInfo.OutsideDiameter + 2.0 * pipeInfo.InsulationThickness;
                }
                // Check if overhangs are small
                Double overhang = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyOverhang", "Overhang")).PropValue;
                Collection<object> overHang1Collection = new Collection<object>();
                GenericHelper.GetDataByRule("HgrOverhang", support, out overHang1Collection);
                double overHang1 = 0;
                if (overHang1Collection != null)
                {
                    if (overHang1Collection[0] == null)
                        overHang1 = Convert.ToDouble(overHang1Collection[1]); // For Straight sections, A is the OH dimension
                    else
                        overHang1 = Convert.ToDouble(overHang1Collection[0]); // For Straight sections, A is the OH dimension
                }

                if (!GenericAssemblyServices.CmpDblEqual(overhang, overHang1))
                {
                    if (GenericAssemblyServices.CmpDblLessThan(overhang, overHang1))
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Parts" + ": " + "WARNING: " + "Overhang is too small. Please check.", "", "UFrameTypeG.cs", 202);
                }

                // Get the Hgr Overhang from the catalog. Check this value with the value on assembly. If
                // both are not equal, it means that the user modified it. Then use the value given by user
                Double braceAngle = Convert.ToDouble(metadataManager.GetCodelistInfo("GenAssyBraceAngle", "UDP").GetCodelistItem(angle).DisplayName);
                BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                double catalogOverHang = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrGenAssyOverhang", "Overhang")).PropValue;
                if (HgrCompareDoubleService.cmpdbl(overhang , catalogOverHang)==true)
                {
                    GenericHelper.GetDataByRule("HgrOverhang", support, out overHang1Collection);
                    if (overHang1Collection != null)
                    {
                        if (overHang1Collection[0] == null)
                            overhang = Convert.ToDouble(overHang1Collection[2]);
                        else
                            overhang = Convert.ToDouble(overHang1Collection[1]);
                    }
                }
                // Get Pipe Slope
                double pipeAngle1 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                double routeStructDistance = 0;
                if ((pipeAngle1 < (0 + 0.0001) && pipeAngle1 > (0 - 0.0001)) || (pipeAngle1 < (Math.PI + 0.0001) && pipeAngle1 > (Math.PI - 0.0001)))
                    routeStructDistance = GenericAssemblyServices.GetMaximumRouteStructDistance(this, PortDistanceType.Horizontal);
                else if (pipeAngle1 < (Math.PI / 2 + 0.0001) && pipeAngle1 > (Math.PI / 2 - 0.0001))
                    routeStructDistance = GenericAssemblyServices.GetMaximumRouteStructDistance(this, PortDistanceType.Vertical);

                // =======================================
                // Do Something if more than one Structure
                // =======================================
                // Determine whether connecting to Steel or a Slab
                string connection = string.Empty;
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByReference && (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        connection = "Steel";
                    else
                        connection = "Slab";
                }
                else
                    connection = "Slab";

                // Determine whether connecting to Steel which is perpendicular. If so, give a error
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && Math.Abs(routeStructAngle) > (Math.PI - 0.001)))
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Please use Frame Support Type B for this Structure and Route configuration.", "", "UFrameTypeG.cs", 235);

                double horizontalDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                // =======================================
                // Set Values of Part Occurance Attributes
                // =======================================
                PropertyValueCodelist codelist = (PropertyValueCodelist)componentDictionary[VERTSECTION].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (codelist.PropValue == -1)
                    codelist.PropValue = 1;
                PropertyValueCodelist endmitercodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endmitercodelist.PropValue == -1)
                    endmitercodelist.PropValue = 1;
                componentDictionary[VERTSECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERTSECTION].SetPropertyValue(codelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERTSECTION].SetPropertyValue(endmitercodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                // For UBolt
                double[] pipeODWoInsulation = new double[0];
                string uBoltType = string.Empty;
                double padThicknessValue = 0;
                GenericHelper.GetDataByRule("HgrUBoltType", support, out uBoltType);
                for (uBoltIndex = uBoltBeginIndex; uBoltIndex <= uBoltEndIndex; uBoltIndex++)
                {
                    if (uBoltIndex == uBoltBeginIndex)
                        routeIndex = 1;
                    else
                        routeIndex = uBoltIndex - uBoltBeginIndex + 1;
                    Array.Resize(ref pipeODWoInsulation, routeIndex);
                    PipeObjectInfo pipeinfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                    pipeODWoInsulation[routeIndex - 1] = GenericAssemblyServices.GetPipeODwithoutInsulatation(pipeDiameter[routeIndex - 1], pipeinfo);
                    if (uBoltType == "E")
                    {
                        padThicknessValue = 0.005;
                        componentDictionary[uboltPartKeys[uBoltIndex - 1]].SetPropertyValue(sectionWidth, "IJOAHgrGenUBoltPadL", "PadL");
                    }
                    componentDictionary[uboltPartKeys[uBoltIndex - 1]].SetPropertyValue(sectionThickness, "IJOAHgrGenUBoltFlgThick", "FlgThick");
                }

                // Set properties on Assembly
                sizeOfSection = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sectionSize).Value;
                support.SetPropertyValue(sizeOfSection, "IJOAHgrGenAssySection", "SectionSize");
                support.SetPropertyValue(overhang, "IJOAHgrGenAssyOverhang", "Overhang");
                // =======================================
                // Create Joints
                // =======================================
                double vertSecLength = routeStructDistance + overhang;
                componentDictionary[VERTSECTION].SetPropertyValue(vertSecLength, "IJUAHgrOccLength", "Length");

                // For Braces
                double planeOffset = 0, padThickness = 0, angleRadian1 = 0, braceLength = 0, bracePadOffset = 0, braceHtOffset = 0, leftPlaneOffset = 0, braceHtOffFactor = 0;
                string padPort = string.Empty, sectionPort = string.Empty, hangerProperty = string.Empty, bracePort = string.Empty, sectionPort1 = string.Empty, sectionPort2 = string.Empty, bracePadPort = string.Empty;
                GenericAssemblyServices.ConfigIndex braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(), braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(), structConfigIndex = new GenericAssemblyServices.ConfigIndex(), routeConfigIndex = new GenericAssemblyServices.ConfigIndex();

                if (showBrace == 1)
                {
                    overHang1Collection = new Collection<object>();
                    GenericHelper.GetDataByRule("HgrBraceHeightOffset", support, out overHang1Collection);
                    if (overHang1Collection[0] == null)
                        braceHtOffFactor = Convert.ToDouble(overHang1Collection[1]);
                    else
                        braceHtOffFactor = Convert.ToDouble(overHang1Collection[0]);
                    braceLength = (1 - braceHtOffFactor) * vertSecLength / Math.Cos(braceAngle * Math.PI / 180) + ((sectionWidth - sectionThickness) * Math.Sin(braceAngle * Math.PI / 180));
                    bracePadOffset = (braceLength) * Math.Sin(braceAngle * Math.PI / 180);
                    braceHtOffset = braceHtOffFactor * vertSecLength;
                }

                if (connection == "Steel" && SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                {
                    double byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                    double byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

                    if (Math.Abs(byPointAngle2) > Math.PI / 2.0)
                    {
                        if (Math.Abs(byPointAngle1) < Math.PI / 2.0)
                        {
                            if (showBrace == 1)
                            {
                                braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                                braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                                if (Configuration == 1)
                                {
                                    sectionPort1 = "EndCap";
                                    sectionPort2 = "BeginCap";
                                    bracePort = "EndCap";
                                    bracePadPort = "HgrPort_1";
                                    leftPlaneOffset = braceHtOffset;
                                }
                                else
                                {
                                    sectionPort1 = "BeginCap";
                                    sectionPort2 = "EndCap";
                                    bracePort = "BeginCap";
                                    bracePadPort = "HgrPort_2";
                                    leftPlaneOffset = -braceHtOffset;
                                }
                            }

                            if (Configuration == 1)
                            {
                                structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                                routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                                planeOffset = horizontalDistance - pipeODWoInsulation[0] / 2 - padThicknessValue;
                                sectionPort = "BeginCap";
                                padPort = "HgrPort_1";
                                hangerProperty = "BeginOverLength";
                            }
                            else
                            {
                                structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                                routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX);
                                planeOffset = -horizontalDistance - pipeODWoInsulation[0] / 2 - padThicknessValue;
                                sectionPort = "EndCap";
                                padPort = "HgrPort_2";
                                hangerProperty = "EndOverLength";
                            }
                        }
                        else
                        {
                            if (showBrace == 1)
                            {
                                braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                                braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                                if (Configuration == 1)
                                {
                                    sectionPort1 = "BeginCap";
                                    sectionPort2 = "EndCap";
                                    bracePort = "BeginCap";
                                    bracePadPort = "HgrPort_2";
                                    leftPlaneOffset = -braceHtOffset;
                                }
                                else
                                {
                                    sectionPort1 = "EndCap";
                                    sectionPort2 = "BeginCap";
                                    bracePort = "EndCap";
                                    bracePadPort = "HgrPort_1";
                                    leftPlaneOffset = braceHtOffset;
                                }
                            }

                            if (Configuration == 1)
                            {
                                structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                                routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX);
                                planeOffset = horizontalDistance - pipeODWoInsulation[0] / 2 - padThicknessValue;
                                sectionPort = "EndCap";
                                padPort = "HgrPort_2";
                                hangerProperty = "EndOverLength";
                            }
                            else
                            {
                                structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                                routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                                planeOffset = -horizontalDistance - pipeODWoInsulation[0] / 2 - padThicknessValue;
                                sectionPort = "BeginCap";
                                padPort = "HgrPort_1";
                                hangerProperty = "BeginOverLength";
                            }
                        }
                    }
                    else // The structure is oriented in the opposite direction
                    {
                        if (Math.Abs(byPointAngle1) < Math.PI / 2.0)
                        {
                            if (showBrace == 1)
                            {
                                braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                                braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.NegativeX);
                                if (Configuration == 1)
                                {
                                    sectionPort1 = "BeginCap";
                                    sectionPort2 = "EndCap";
                                    bracePort = "BeginCap";
                                    bracePadPort = "HgrPort_2";
                                    leftPlaneOffset = braceHtOffset;
                                }
                                else
                                {
                                    sectionPort1 = "EndCap";
                                    sectionPort2 = "BeginCap";
                                    bracePort = "EndCap";
                                    bracePadPort = "HgrPort_1";
                                    leftPlaneOffset = -braceHtOffset;
                                }
                            }

                            if (Configuration == 1)
                            {
                                structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                                routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                                planeOffset = horizontalDistance - pipeODWoInsulation[0] / 2 - padThicknessValue;
                                sectionPort = "EndCap";
                                padPort = "HgrPort_2";
                                hangerProperty = "EndOverLength";
                            }
                            else
                            {
                                structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                                routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX);
                                planeOffset = -horizontalDistance - pipeODWoInsulation[0] / 2 - padThicknessValue;
                                sectionPort = "BeginCap";
                                padPort = "HgrPort_1";
                                hangerProperty = "BeginOverLength";
                            }
                        }
                        else
                        {
                            if (showBrace == 1)
                            {
                                braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                                braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                                if (Configuration == 1)
                                {
                                    sectionPort1 = "EndCap";
                                    sectionPort2 = "BeginCap";
                                    bracePort = "EndCap";
                                    bracePadPort = "HgrPort_1";
                                    leftPlaneOffset = braceHtOffset;
                                }
                                else
                                {
                                    sectionPort1 = "BeginCap";
                                    sectionPort2 = "EndCap";
                                    bracePort = "BeginCap";
                                    bracePadPort = "HgrPort_2";
                                    leftPlaneOffset = -braceHtOffset;
                                }
                            }

                            if (Configuration == 1)
                            {
                                structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                                routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX);
                                planeOffset = horizontalDistance - pipeODWoInsulation[0] / 2 - padThicknessValue;
                                sectionPort = "BeginCap";
                                padPort = "HgrPort_1";
                                hangerProperty = "BeginOverLength";
                            }
                            else
                            {
                                structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                                routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                                planeOffset = -horizontalDistance - pipeODWoInsulation[0] / 2 - padThicknessValue;
                                sectionPort = "EndCap";
                                padPort = "HgrPort_2";
                                hangerProperty = "EndOverLength";
                            }
                        }
                    }
                }
                if ((connection == "Steel" && SupportHelper.PlacementType == PlacementType.PlaceByStruct) || connection == "Slab")
                {
                    if (showBrace == 1)
                    {
                        braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                        braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                        sectionPort1 = "EndCap";
                        sectionPort2 = "BeginCap";
                        bracePort = "EndCap";
                        bracePadPort = "HgrPort_1";
                    }
                    if (Configuration == 1)
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        {
                            structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.Y);
                            routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX);
                            planeOffset = -pipeODWoInsulation[0] / 2 - padThicknessValue;
                            sectionPort = "BeginCap";
                            padPort = "HgrPort_1";
                            hangerProperty = "BeginOverLength";
                        }
                        else
                        {
                            if (connection == "Slab")
                            {
                                structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                                routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.NegativeX);
                                planeOffset = -pipeODWoInsulation[0] / 2 - padThicknessValue;
                                sectionPort = "BeginCap";
                                padPort = "HgrPort_1";
                                hangerProperty = "BeginOverLength";
                            }
                        }
                        leftPlaneOffset = braceHtOffset;
                    }
                    else
                    {
                        if (showBrace == 1)
                        {
                            braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                            braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                            sectionPort1 = "EndCap";
                            sectionPort2 = "BeginCap";
                            bracePort = "EndCap";
                            bracePadPort = "HgrPort_2";
                        }
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        {
                            structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);
                            routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                            planeOffset = -pipeODWoInsulation[0] / 2 - padThicknessValue;
                            sectionPort = "EndCap";
                            padPort = "HgrPort_2";
                            hangerProperty = "EndOverLength";
                        }
                        else
                        {
                            if (connection == "Slab")
                            {
                                structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                                routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                                planeOffset = -pipeODWoInsulation[0] / 2 - padThicknessValue;
                                sectionPort = "EndCap";
                                padPort = "HgrPort_2";
                                hangerProperty = "EndOverLength";
                            }
                        }
                        leftPlaneOffset = -braceHtOffset;
                    }
                }
                if (includePad == 1)
                {
                    padThickness = (double)((PropertyValueDouble)componentDictionary[SECPAD].GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                    componentDictionary[VERTSECTION].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", hangerProperty);

                    // Add Joint Between the Plate and the Vertical Section
                    JointHelper.CreateRigidJoint(SECPAD, padPort, VERTSECTION, sectionPort, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                }

                // Add Joint Between the Vertical Section and Steel
                JointHelper.CreateRigidJoint(VERTSECTION, sectionPort, "-1", "Structure", structConfigIndex.A, structConfigIndex.B, structConfigIndex.C, structConfigIndex.D, 0, planeOffset, sectionWidth / 2);

                string strRefPortName = string.Empty;
                for (uBoltIndex = uBoltBeginIndex; uBoltIndex <= uBoltEndIndex; uBoltIndex++)
                {
                    if (uBoltIndex == uBoltBeginIndex)
                        strRefPortName = "Route";
                    else
                        strRefPortName = "Route_" + Convert.ToString(uBoltIndex - uBoltBeginIndex + 1);

                    // Add Joint Between the UBolt and Route
                    JointHelper.CreateRigidJoint(uboltPartKeys[uBoltIndex - 1], "Route", "-1", strRefPortName, routeConfigIndex.A, routeConfigIndex.B, routeConfigIndex.C, routeConfigIndex.D, 0, 0, 0);
                }

                // For Braces
                if (showBrace == 1)
                {
                    angleRadian1 = braceAngle * Math.PI / 180;
                    if (includeBracePad == 1)
                    {
                        padThickness = (double)((PropertyValueDouble)componentDictionary[BRACEPAD].GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;

                        // Add Joint Between the Plate and the Vertical Section
                        JointHelper.CreateRigidJoint(BRACEPAD, bracePadPort, VERTSECTION, sectionPort2, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -bracePadOffset, sectionWidth / 3);
                        componentDictionary[BRACE].SetPropertyValue(braceLength - padThickness, "IJOAHgrGenericBrace", "L");
                    }
                    else
                        componentDictionary[BRACE].SetPropertyValue(braceLength, "IJOAHgrGenericBrace", "L");

                    codelist = (PropertyValueCodelist)componentDictionary[BRACE].GetPropertyValue("IJOAHgrGenericBrace", "BraceOrient");
                    if (codelist.PropValue == -1)
                        codelist.PropValue = 1;
                    componentDictionary[BRACE].SetPropertyValue(angleRadian1, "IJOAHgrGenericBrace", "Angle");
                    componentDictionary[BRACE].SetPropertyValue(codelist.PropValue = 2, "IJOAHgrGenericBrace", "BraceOrient");
                    componentDictionary[BRACE].SetPropertyValue(sectionWidth, "IJOAHgrGenericBrace", "W");
                    componentDictionary[BRACE].SetPropertyValue(sectionDepth, "IJOAHgrGenericBrace", "Depth");
                    componentDictionary[BRACE].SetPropertyValue(sectionThickness, "IJOAHgrGenericBrace", "FlangeT");

                    // Add a Joint between horizontal section and the point that we want the angle to rotate around.
                    JointHelper.CreateRigidJoint(CONNECTIONOBJ, "Connection", VERTSECTION, sectionPort1, braceConfigIndex1.A, braceConfigIndex1.B, braceConfigIndex1.C, braceConfigIndex1.D, leftPlaneOffset, sectionThickness, 0);
                    JointHelper.CreateRigidJoint(CONNECTIONOBJ, "Connection", BRACE, bracePort, braceConfigIndex2.A, braceConfigIndex2.B, braceConfigIndex2.C, braceConfigIndex2.D, 0, 0, 0);
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
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(uboltPartKeys[1], 1));

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
                    structConnections.Add(new ConnectionInfo(VERTSECTION, 1));

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
