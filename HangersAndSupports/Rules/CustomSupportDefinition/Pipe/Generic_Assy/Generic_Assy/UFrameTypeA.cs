//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   UFrameTypeA.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeA
//   Author       : Manikanth
//   Creation Date: 17-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-Sep-2013    Manikanth  CR-CP-224494 Convert HS_Generic_Assy to C# .Net 
//   31-Oct-2014     PVK       TR-CP-260301	Resolve coverity issues found in August 2014 report
//   22-Jan-2015     PVK       TR-CP-264951  Resolve coverity issues found in November 2014 report
//   06-June-2016    PVK       TR-CP-296065	Fix new coverity issues found in H&S Content 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
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

    public class UFrameTypeA : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string HORSECTION = "HORSECTION";
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";
        private const string UBOLT = "UBOLT";
        private string LEFTBRACE = "LEFTBRACE";
        private string RIGHTBRACE = "RIGHTBRACE";
        private string LEFTCONNOBJ = "LEFTCONNOBJ";
        private string RIGHTCONNOBJ = "RIGHTCONNOBJ";
        private string LEFTPAD = "LEFTPAD";
        private string RIGHTPAD = "RIGHTPAD";
        private string LEFTBRPAD = "LEFTBRPAD";
        private string RIGHTBRPAD = "RIGHTBRPAD";
        private string CONNOBJ1 = "CONNOBJ1";
        private string CONNOBJ2 = "CONNOBJ2";
        private string uBoltPart, angularPadPart, sectionSize, fromOrient;
        string sectionFromRule;
        private double braceAngle, overhangLeft, overhangRight, widthOffset, heightOffset;
        private bool value;
        PropertyValueCodelist offsetsFromRule, sectionFromRuleCodelist, sectionCodeList, overhangOption, leftPad, rightPad, showBrace, leftBracePad, rightBracePad, connection, frmOrientCodeList, braceAngleCodelist;
        Collection<object> collection;
        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                IEnumerable<BusinessObject> padPart = null;
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    sectionFromRuleCodelist = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionFromRule"));
                    sectionFromRule = (sectionFromRuleCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(sectionFromRuleCodelist.PropValue).DisplayName);
                    sectionCodeList = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionSize"));
                    sectionSize = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionCodeList.PropValue).DisplayName;
                    overhangOption = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyOHOpt", "OverhangOpt"));
                    leftPad = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyLeftPad", "IncludeLeftPad"));
                    rightPad = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyRightPad", "IncludeRightPad"));
                    showBrace = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrace", "ShowBrace"));
                    leftBracePad = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyLeftBrPad", "IncludeLeftBrPad"));
                    rightBracePad = (((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyRightBrPad", "IncludeRightBrPad")));
                    overhangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyOHLeft", "OverhangLeft")).PropValue;
                    overhangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyOHRight", "OverhangRight")).PropValue;
                    connection = (((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyConn", "Connection")));
                    frmOrientCodeList = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyFrmOrient", "FrmOrient"));
                    fromOrient = (frmOrientCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(frmOrientCodeList.PropValue).DisplayName);
                    braceAngleCodelist = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrace", "BraceAngle"));
                    braceAngle = Convert.ToDouble(braceAngleCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(braceAngleCodelist.PropValue).DisplayName);
                    offsetsFromRule = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyFrmOffsets", "OffsetsFromRule"));

                    if (offsetsFromRule.PropValue == 1)
                    {
                        collection = new Collection<object>();
                        value = GenericHelper.GetDataByRule("HgrFrmOffsets", support, out collection);
                        if (collection != null)
                        {
                            if (collection[0] == null)
                            {
                                widthOffset = (double)collection[1];
                                heightOffset = (double)collection[2];
                            }
                            else
                            {
                                widthOffset = (double)collection[0];
                                heightOffset = (double)collection[1];
                            }
                        }
                    }
                    else
                    {
                        widthOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyFrmOffsets", "WidthOffset")).PropValue;
                        heightOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyFrmOffsets", "HeightOffset")).PropValue;
                    }

                    //Get the Section Size
                    if (sectionFromRule == "True")
                        value = GenericHelper.GetDataByRule("HgrSectionSize", (BusinessObject)support, out sectionSize);
                    else if (sectionFromRule == "False")
                    {
                        PropertyValueCodelist sectionSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionSize");
                        sectionSize = sectionSizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodeList.PropValue).DisplayName;

                        if (sectionSize.ToUpper() == "NONE" || sectionSize == string.Empty)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "UFrameTypeA.cs", 110);
                    }
                    collection = new Collection<object>();
                    value = GenericHelper.GetDataByRule("HgrUBoltSelection", support, out collection);
                    if (collection != null)
                    {
                        if (collection[0] == null)
                            uBoltPart = (string)collection[1];
                        else
                            uBoltPart = (string)collection[0];
                    }
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    PartClass padPartclass = (PartClass)catalogBaseHelper.GetPartClass("GenServ_AnglePadDim");
                    padPart = padPartclass.Parts;
                    if (padPartclass.PartClassType.Equals("HgrServiceClass"))
                        padPart = padPartclass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    else
                        padPart = padPartclass.Parts;
                    padPart = padPart.Where(part1 => (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrGenServAnglePad", "SectionSize")).PropValue == sectionSize);

                    if (padPart.Count() > 0)
                        angularPadPart = (string)((PropertyValueString)padPart.ElementAt(0).GetPropertyValue("IJUAHgrGenServAnglePad", "PadPart")).PropValue;

                    parts.Add(new PartInfo(HORSECTION, sectionSize));
                    parts.Add(new PartInfo(VERTSECTION1, sectionSize));
                    parts.Add(new PartInfo(VERTSECTION2, sectionSize));
                    parts.Add(new PartInfo(UBOLT, uBoltPart));
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                    {
                        parts.Add(new PartInfo(CONNOBJ1, "Log_Conn_Part_1"));
                        parts.Add(new PartInfo(CONNOBJ2, "Log_Conn_Part_1"));
                    }

                    if (leftPad.PropValue == 1)
                        parts.Add(new PartInfo(LEFTPAD, angularPadPart));
                    if (rightPad.PropValue == 1)
                        parts.Add(new PartInfo(RIGHTPAD, angularPadPart));

                    if (showBrace.PropValue == 1)
                    {
                        parts.Add(new PartInfo(LEFTBRACE, "HgrGen_GenericBrace_1"));
                        parts.Add(new PartInfo(LEFTCONNOBJ, "Log_Conn_Part_1"));
                        parts.Add(new PartInfo(RIGHTBRACE, "HgrGen_GenericBrace_1"));
                        parts.Add(new PartInfo(RIGHTCONNOBJ, "Log_Conn_Part_1"));
                    }

                    if (leftBracePad.PropValue == 1)
                    {
                        if (showBrace.PropValue == 1)
                            parts.Add(new PartInfo(LEFTBRPAD, angularPadPart));
                        else if (showBrace.PropValue == 2)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Brace is not available. Resetting the value.", "", "UFrameTypeA.cs", 157);
                            leftBracePad.PropValue = 2;
                            return parts;
                        }
                    }
                    if (rightBracePad.PropValue == 1)
                    {
                        if (showBrace.PropValue == 1)
                            parts.Add(new PartInfo(RIGHTBRPAD, angularPadPart));
                        else if (showBrace.PropValue == 2)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Brace is not available. Resetting the value.", "", "UFrameTypeA.cs", 168);
                            rightBracePad.PropValue = 2;
                            return parts;
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
                finally
                {
                    if (padPart is IDisposable)
                    {
                        ((IDisposable)padPart).Dispose(); // This line will be executed
                    }
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

                GenericAssemblyServices.CreateDimensionNote(this, "Dim 1", componentDictionary[UBOLT], "Route");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim 2", componentDictionary[HORSECTION], "BeginCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim 3", componentDictionary[HORSECTION], "EndCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim 4", componentDictionary[VERTSECTION1], "EndCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim 5", componentDictionary[VERTSECTION2], "BeginCap");

                if (showBrace.PropValue == 1)
                {
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim 6", componentDictionary[LEFTBRACE], "BeginCap");
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim 7", componentDictionary[LEFTBRACE], "EndCap");
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim 8", componentDictionary[RIGHTBRACE], "BeginCap");
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim 9", componentDictionary[RIGHTBRACE], "EndCap");
                }
                BusinessObject horizontalSectionPart = componentDictionary[HORSECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                double sectionWidth = 0.0, sectionDepth = 0.0, sectionThickness = 0.0;
                sectionWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                sectionDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                sectionThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                if (widthOffset > 3 * sectionWidth / 4)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Width Offset should be less than 3/4th of the Section Width.", "", "UFrameTypeA.cs", 214);
                if (heightOffset > 3 * sectionDepth / 4)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Height Offset should be less than 3/4th of the Section Width.", "", "UFrameTypeA.cs", 216);

                bool[] isOffsetApplied = GenericAssemblyServices.GetIsLugEndOffsetApplied(this);
                string[] structPort = GenericAssemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0], rightStructPort = structPort[1];
                double overHangLeft = 0, overHangRight = 0;
                value = GenericHelper.GetDataByRule("HgrOverhang", support, out  collection);
                if (collection != null)
                {
                    if (collection[0] == null)
                    {
                        overHangLeft = (double)collection[1];
                        overHangRight = (double)collection[1];
                    }
                    else
                    {
                        overHangLeft = (double)collection[0];
                        overHangRight = (double)collection[0];
                    }
                }
                if (overhangOption.PropValue == 3)
                {
                    if (HgrCompareDoubleService.cmpdbl(overhangLeft, overHangLeft) == false)
                    {
                        if (overhangLeft < overHangLeft)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Left overhang is too small. Please check.", "", "UFrameTypeA.cs", 230);
                    }
                    if (HgrCompareDoubleService.cmpdbl(overhangRight, overHangRight) == false)
                    {
                        if (overhangRight < overHangRight)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Right overhang is too small. Please check.", "", "UFrameTypeA.cs", 235);
                    }
                }
                if (overhangOption.PropValue == 1)
                    overhangLeft = overhangRight = overHangLeft;
                else if (overhangOption.PropValue == 2)
                {
                    double horDistRouteLeftStruct, horDistRouteRightStruct;
                    if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                    {
                        if (Configuration == 1 || Configuration == 4)
                        {
                            horDistRouteLeftStruct = RefPortHelper.DistanceBetweenPorts("Route", leftStructPort, PortDistanceType.Horizontal);
                            horDistRouteRightStruct = RefPortHelper.DistanceBetweenPorts("Route", rightStructPort, PortDistanceType.Horizontal);

                            overhangLeft = horDistRouteLeftStruct + sectionWidth / 2;
                            overhangRight = horDistRouteRightStruct + sectionWidth / 2;
                        }

                        else if (Configuration == 2 || Configuration == 3)
                        {
                            horDistRouteLeftStruct = RefPortHelper.DistanceBetweenPorts("Route", rightStructPort, PortDistanceType.Horizontal);
                            horDistRouteRightStruct = RefPortHelper.DistanceBetweenPorts("Route", leftStructPort, PortDistanceType.Horizontal);

                            overhangLeft = horDistRouteLeftStruct + sectionWidth / 2;
                            overhangRight = horDistRouteRightStruct + sectionWidth / 2;
                        }
                    }
                }
                PropertyValueCodelist topBeginMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (topBeginMiterCodelist.PropValue == -1)
                    topBeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist topEndMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (topEndMiterCodelist.PropValue == -1)
                    topEndMiterCodelist.PropValue = 1;

                PropertyValueCodelist angleBeginMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION1]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (angleBeginMiterCodelist.PropValue == -1)
                    angleBeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist angleEndMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION1]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (angleEndMiterCodelist.PropValue == -1)
                    angleEndMiterCodelist.PropValue = 1;

                PropertyValueCodelist angleBeginMiterCodelist1 = (PropertyValueCodelist)(componentDictionary[VERTSECTION2]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (angleBeginMiterCodelist1.PropValue == -1)
                    angleBeginMiterCodelist1.PropValue = 1;
                PropertyValueCodelist angleendMiterCodelist1 = (PropertyValueCodelist)(componentDictionary[VERTSECTION2]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (angleendMiterCodelist1.PropValue == -1)
                    angleendMiterCodelist1.PropValue = 1;

                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[HORSECTION]).SetPropertyValue(topBeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[HORSECTION]).SetPropertyValue(topEndMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                (componentDictionary[VERTSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(topBeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[VERTSECTION1]).SetPropertyValue(topEndMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                (componentDictionary[VERTSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(topBeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[VERTSECTION2]).SetPropertyValue(topEndMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                PipeObjectInfo pipeinfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = pipeinfo.OutsideDiameter, slope;

                Vector orientation = pipeinfo.Orientation;
                double distance = Math.Sqrt(orientation.X * orientation.X + orientation.Y * orientation.Y);
                if (distance < Math3d.DistanceTolerance)
                    slope = Math.PI / 2;
                else
                    slope = Math.Atan(Math.Abs(orientation.Z) / orientation.Length);

                double pipeAngle = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Z);
                double angledOffset = Math.Tan(pipeAngle) * sectionWidth, angledOffset2 = Math.Sin(pipeAngle) * sectionWidth / 2;

                double extraLength = 0, padThickness = 0, routeStructDistance, routeLeftStructDistance, routeRightStructDistance, verticalLength = 0, leftSectionLength = 0, rightSectionLength = 0;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                        extraLength = angledOffset;
                }

                (componentDictionary[UBOLT]).SetPropertyValue(sectionThickness, "IJOAHgrGenUBoltFlgThick", "FlgThick");
                value = GenericHelper.GetDataByRule("HgrUBoltType", support, out  collection);
                string uBoltType = string.Empty;
                if (collection != null)
                {
                    if (collection[0] == null)
                        uBoltType = (string)collection[1];
                    else
                        uBoltType = (string)collection[0];
                }

                if (uBoltType == "E")
                {
                    padThickness = 0.005;
                    (componentDictionary[UBOLT]).SetPropertyValue(sectionWidth, "IJOAHgrGenUBoltPadL", "PadL");
                }
                (componentDictionary[UBOLT]).SetPropertyValue(sectionThickness, "IJOAHgrGenUBoltFlgThick", "FlgThick");
                if ((componentDictionary[UBOLT]).SupportsInterface("IJOAHgrGenericUBoltE"))
                    (componentDictionary[UBOLT]).SetPropertyValue(pipeDiameter / 2, "IJOAHgrGenericUBoltE", "PipeRadius");
                else if ((componentDictionary[UBOLT]).SupportsInterface("IJOAHgrGenericUBolt"))
                    (componentDictionary[UBOLT]).SetPropertyValue(pipeDiameter / 2, "IJOAHgrGenericUBolt", "PipeRadius");
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                int size = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sectionSize).Value;

                support.SetPropertyValue(size, "IJOAHgrGenAssySection", "SectionSize");
                support.SetPropertyValue(overhangLeft, "IJOAHgrGenAssyOHLeft", "OverhangLeft");
                support.SetPropertyValue(overhangRight, "IJOAHgrGenAssyOHRight", "OverhangRight");
                (componentDictionary[HORSECTION]).SetPropertyValue(overhangLeft + overhangRight, "IJUAHgrOccLength", "Length");

                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);

                if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                {
                    if (HgrCompareDoubleService.cmpdbl(Math.Abs(slope) , 0)==true)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "This option is valid only for Sloped Route. Resetting the value..", "", "UFrameTypeA.cs", 346);
                        frmOrientCodeList.PropValue = 1;
                        support.SetPropertyValue(frmOrientCodeList.PropValue, "IJOAHgrGenAssyFrmOrient", "FrmOrient");
                    }
                    else
                    {
                        if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "This option is not yet implemented. Resetting the value.", "", "UFrameTypeA.cs", 354);
                            frmOrientCodeList.PropValue = 1;
                            support.SetPropertyValue(frmOrientCodeList.PropValue, "IJOAHgrGenAssyFrmOrient", "FrmOrient");
                        }
                    }
                }

                if ((Math.Abs(routeStructAngle) < (0 - slope + 0.001) && Math.Abs(routeStructAngle) > (0 - slope - 0.001)) || (Math.Abs(routeStructAngle) < ((Math.PI - slope) + 0.001) && Math.Abs(routeStructAngle) > ((Math.PI - slope) - 0.001)))
                {
                    routeStructDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                    routeLeftStructDistance = RefPortHelper.DistanceBetweenPorts("Route", leftStructPort, PortDistanceType.Horizontal);
                    routeRightStructDistance = RefPortHelper.DistanceBetweenPorts("Route", rightStructPort, PortDistanceType.Horizontal);
                }
                else
                {
                    routeStructDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                    routeLeftStructDistance = RefPortHelper.DistanceBetweenPorts("Route", leftStructPort, PortDistanceType.Vertical);
                    routeRightStructDistance = RefPortHelper.DistanceBetweenPorts("Route", rightStructPort, PortDistanceType.Vertical);
                }

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (Configuration == 1 || Configuration == 2)
                        verticalLength = routeStructDistance - pipeDiameter / 2 - padThickness + extraLength;
                    else if (Configuration == 3 || Configuration == 4)
                        verticalLength = routeStructDistance + pipeDiameter / 2 + sectionDepth + padThickness + extraLength;

                    (componentDictionary[VERTSECTION1]).SetPropertyValue(verticalLength, "IJUAHgrOccLength", "Length");
                    (componentDictionary[VERTSECTION2]).SetPropertyValue(verticalLength, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    if (Configuration == 1 || Configuration == 2)
                    {
                        leftSectionLength = routeLeftStructDistance - pipeDiameter / 2 - padThickness + extraLength;
                        rightSectionLength = routeRightStructDistance - pipeDiameter / 2 - padThickness + extraLength;
                    }
                    else if (Configuration == 3 || Configuration == 4)
                    {
                        leftSectionLength = routeLeftStructDistance + pipeDiameter / 2 + sectionDepth + padThickness + extraLength;
                        rightSectionLength = routeRightStructDistance + pipeDiameter / 2 + sectionDepth + padThickness + extraLength;
                    }
                    (componentDictionary[VERTSECTION1]).SetPropertyValue(leftSectionLength, "IJUAHgrOccLength", "Length");
                    (componentDictionary[VERTSECTION2]).SetPropertyValue(rightSectionLength, "IJUAHgrOccLength", "Length");
                }
                if (showBrace.PropValue == 1)
                {
                    double leftPlaneOffset = 0, leftAxisOffset = 0, rightPlaneOffset = 0, rightAxisOffset = 0, bracePlaneOffset = 0, braceAxisOffset = 0, braceOriginOffset = 0, angle1 = 0, leftBraceLength = 0, rightBraceLength = 0;
                    double leftPadOffset = 0, rightPadOffset = 0, leftBraceHtOffset = 0, rightBraceHtOffset, bracePadPlaneOffset1 = 0, bracePadAxisOffset1 = 0, bracePadOriginOffset1 = 0, bracePadPlaneOffset2 = 0, bracePadAxisOffset2 = 0, bracePadOriginOffset2 = 0;
                    string sectionPort1 = string.Empty, sectionPort2 = string.Empty, bracePort1 = string.Empty, bracePort2 = string.Empty;

                    value = GenericHelper.GetDataByRule("HgrBraceHeightOffset", support, out collection);
                    double braceHtFactor = 0;
                    if (collection != null)
                    {
                        if (collection[0] == null)
                            braceHtFactor = (double)collection[1];
                        else
                            braceHtFactor = (double)collection[0];
                    }
                    angle1 = braceAngle * (Math.PI / 180);
                    GenericAssemblyServices.ConfigIndex braceIndex1 = new GenericAssemblyServices.ConfigIndex(), braceIndex2 = new GenericAssemblyServices.ConfigIndex(), bracePadIndex = new GenericAssemblyServices.ConfigIndex();

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        leftBraceLength = (1 - braceHtFactor) * verticalLength / Math.Cos((braceAngle * (Math.PI / 180))) + ((sectionWidth - sectionThickness) * Math.Sin((braceAngle * (Math.PI / 180))));
                        leftPadOffset = (leftBraceLength) * Math.Sin((braceAngle * (Math.PI / 180)));
                        leftBraceHtOffset = braceHtFactor * verticalLength;
                        rightBraceLength = leftBraceLength;
                        rightBraceHtOffset = leftBraceHtOffset;
                        rightPadOffset = leftPadOffset;
                    }
                    else
                    {
                        leftBraceLength = (1 - braceHtFactor) * leftSectionLength / Math.Cos((braceAngle * (Math.PI / 180))) + ((sectionWidth - sectionThickness) * Math.Sin((braceAngle * (Math.PI))));
                        rightBraceLength = (1 - braceHtFactor) * rightSectionLength / Math.Cos(braceAngle * (Math.PI / 180)) + ((sectionWidth - sectionThickness) * Math.Sin((braceAngle * (Math.PI / 180))));
                        leftPadOffset = (leftBraceLength) * Math.Sin((braceAngle * (Math.PI / 180)));
                        rightPadOffset = (rightBraceLength) * Math.Sin(braceAngle * (Math.PI / 180));
                        leftBraceHtOffset = braceHtFactor * leftSectionLength;
                        rightBraceHtOffset = braceHtFactor * rightSectionLength;
                    }
                    if (Configuration == 1 || Configuration == 4)
                    {
                        braceIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.Y, Axis.NegativeX);
                        braceIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                        bracePadIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                        sectionPort1 = "BeginCap";
                        sectionPort2 = "EndCap";
                        bracePort1 = "BeginCap";
                        bracePort2 = "EndCap";

                        bracePadPlaneOffset1 = 0;
                        bracePadPlaneOffset2 = 0;
                        bracePadAxisOffset1 = -sectionWidth / 3;
                        bracePadAxisOffset2 = -sectionWidth / 3;
                        bracePadOriginOffset1 = leftPadOffset;
                        bracePadOriginOffset2 = rightPadOffset;

                    }
                    else if (Configuration == 2 || Configuration == 3)
                    {
                        braceIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                        braceIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                        bracePadIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                        sectionPort1 = "BeginCap";
                        sectionPort2 = "EndCap";
                        bracePort1 = "EndCap";
                        bracePort2 = "BeginCap";

                        bracePadPlaneOffset1 = 0;
                        bracePadPlaneOffset2 = 0;
                        bracePadAxisOffset1 = -leftPadOffset;
                        bracePadAxisOffset2 = -rightPadOffset;
                        bracePadOriginOffset1 = sectionWidth / 3;
                        bracePadOriginOffset2 = sectionWidth / 3;
                    }

                    if (leftBracePad.PropValue == 1 && rightBracePad.PropValue == 2)
                    {
                        padThickness = (double)((PropertyValueDouble)(componentDictionary[LEFTBRPAD]).GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                        //Add Joint Between the Plate and the Vertical Section
                        JointHelper.CreateRigidJoint(LEFTBRPAD, "HgrPort_2", VERTSECTION1, "EndCap", bracePadIndex.A, bracePadIndex.B, bracePadIndex.C, bracePadIndex.D, bracePadPlaneOffset1, bracePadAxisOffset1, bracePadOriginOffset1);
                    }
                    else if (leftBracePad.PropValue == 2 && rightBracePad.PropValue == 1)
                    {
                        padThickness = (double)((PropertyValueDouble)(componentDictionary[RIGHTBRPAD]).GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                        //Add Joint Between the Plate and the Vertical Section
                        JointHelper.CreateRigidJoint(RIGHTBRPAD, "HgrPort_1", VERTSECTION2, "BeginCap", bracePadIndex.A, bracePadIndex.B, bracePadIndex.C, bracePadIndex.D, bracePadPlaneOffset2, bracePadAxisOffset2, bracePadOriginOffset2);
                    }
                    else if (leftBracePad.PropValue == 1 && rightBracePad.PropValue == 1)
                    {
                        padThickness = (double)((PropertyValueDouble)(componentDictionary[LEFTBRPAD]).GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                        //Add Joint Between the Plate and the Vertical Section
                        JointHelper.CreateRigidJoint(LEFTBRPAD, "HgrPort_2", VERTSECTION1, "EndCap", bracePadIndex.A, bracePadIndex.B, bracePadIndex.C, bracePadIndex.D, bracePadPlaneOffset1, bracePadAxisOffset1, bracePadOriginOffset1);
                        //Add Joint Between the Plate and the Vertical Section
                        JointHelper.CreateRigidJoint(RIGHTBRPAD, "HgrPort_1", VERTSECTION2, "BeginCap", bracePadIndex.A, bracePadIndex.B, bracePadIndex.C, bracePadIndex.D, bracePadPlaneOffset2, bracePadAxisOffset2, bracePadOriginOffset2);
                    }
                    if (leftBracePad.PropValue == 1)
                        (componentDictionary[LEFTBRACE]).SetPropertyValue(leftBraceLength - padThickness, "IJOAHgrGenericBrace", "L");
                    else if (leftBracePad.PropValue == 2)
                        (componentDictionary[LEFTBRACE]).SetPropertyValue(leftBraceLength, "IJOAHgrGenericBrace", "L");

                    PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[LEFTBRACE].GetPropertyValue("IJOAHgrGenericBrace", "BraceOrient");
                    braceCodelist.PropValue = 2;
                    (componentDictionary[LEFTBRACE]).SetPropertyValue(angle1, "IJOAHgrGenericBrace", "Angle");
                    (componentDictionary[LEFTBRACE]).SetPropertyValue(braceCodelist.PropValue, "IJOAHgrGenericBrace", "BraceOrient");
                    (componentDictionary[LEFTBRACE]).SetPropertyValue(sectionWidth, "IJOAHgrGenericBrace", "W");
                    (componentDictionary[LEFTBRACE]).SetPropertyValue(sectionDepth, "IJOAHgrGenericBrace", "Depth");
                    (componentDictionary[LEFTBRACE]).SetPropertyValue(sectionThickness, "IJOAHgrGenericBrace", "FlangeT");

                    if (rightBracePad.PropValue == 1)
                        (componentDictionary[RIGHTBRACE]).SetPropertyValue(rightBraceLength - padThickness, "IJOAHgrGenericBrace", "L");
                    else if (rightBracePad.PropValue == 2)
                        (componentDictionary[RIGHTBRACE]).SetPropertyValue(rightBraceLength, "IJOAHgrGenericBrace", "L");

                    (componentDictionary[RIGHTBRACE]).SetPropertyValue(angle1, "IJOAHgrGenericBrace", "Angle");
                    (componentDictionary[RIGHTBRACE]).SetPropertyValue(braceCodelist.PropValue, "IJOAHgrGenericBrace", "BraceOrient");
                    (componentDictionary[RIGHTBRACE]).SetPropertyValue(sectionWidth, "IJOAHgrGenericBrace", "W");
                    (componentDictionary[RIGHTBRACE]).SetPropertyValue(sectionDepth, "IJOAHgrGenericBrace", "Depth");
                    (componentDictionary[RIGHTBRACE]).SetPropertyValue(sectionThickness, "IJOAHgrGenericBrace", "FlangeT");

                    if (Configuration == 1 || Configuration == 4)
                    {
                        leftPlaneOffset = leftBraceHtOffset;
                        rightPlaneOffset = -rightBraceHtOffset;
                    }
                    else if (Configuration == 2 || Configuration == 3)
                    {
                        leftPlaneOffset = -leftBraceHtOffset;
                        rightPlaneOffset = rightBraceHtOffset;
                    }

                    //Add a Joint between horizontal section and the point that we want the angle to rotate around.
                    JointHelper.CreateRigidJoint(LEFTCONNOBJ, "Connection", VERTSECTION1, sectionPort1, braceIndex1.A, braceIndex1.B, braceIndex1.C, braceIndex1.D, leftPlaneOffset, leftAxisOffset, 0);
                    JointHelper.CreateRigidJoint(LEFTCONNOBJ, "Connection", LEFTBRACE, bracePort2, braceIndex2.A, braceIndex2.B, braceIndex2.C, braceIndex2.D, bracePlaneOffset, braceAxisOffset, braceOriginOffset);
                    //Add a Joint between horizontal section and the point that we want the angle to rotate around.
                    JointHelper.CreateRigidJoint(RIGHTCONNOBJ, "Connection", VERTSECTION2, sectionPort2, braceIndex1.A, braceIndex1.B, braceIndex1.C, braceIndex1.D, rightPlaneOffset, rightAxisOffset, 0);
                    JointHelper.CreateRigidJoint(RIGHTCONNOBJ, "Connection", RIGHTBRACE, bracePort1, braceIndex2.A, braceIndex2.B, braceIndex2.C, braceIndex2.D, bracePlaneOffset, braceAxisOffset, braceOriginOffset);
                }
                GenericAssemblyServices.ConfigIndex configIndex1 = new GenericAssemblyServices.ConfigIndex(), configIndex2 = new GenericAssemblyServices.ConfigIndex(), structConfigIndex = new GenericAssemblyServices.ConfigIndex(), routeConfigIndex = new GenericAssemblyServices.ConfigIndex(), routeUboltIndex = new GenericAssemblyServices.ConfigIndex();

                string horizontalPort1 = string.Empty, horizontalPort2 = string.Empty, verticalPort1 = string.Empty, verticalPort2 = string.Empty;
                double horizontalPlaneOffset = 0, horizontalOriginOffset2 = 0, horizontalOriginOffset1 = 0, horizontalAxisOffset2 = 0, horizontalAxisOffset1 = 0, horizontalAxisOffset = 0, horizontalOriginOffset = 0, structAxisOffset = 0, structOriginOffset = 0, hoizontalPlaneOffset1 = 0, horizontalPlaneOffset2 = 0;

                if (Configuration == 1 || Configuration == 4)
                {
                    horizontalPort1 = "EndCap";
                    horizontalPort2 = "BeginCap";
                    verticalPort1 = "BeginCap";
                    verticalPort2 = "EndCap";
                }
                else if (Configuration == 2 || Configuration == 3)
                {
                    horizontalPort1 = "BeginCap";
                    horizontalPort2 = "EndCap";
                    verticalPort1 = "BeginCap";
                    verticalPort2 = "EndCap";
                }

                if (connection.PropValue == 1)
                {
                    widthOffset = 0;
                    heightOffset = 0;

                    if (Configuration == 1 || Configuration == 2)
                    {
                        hoizontalPlaneOffset1 = 0;
                        horizontalPlaneOffset2 = 0;
                        horizontalAxisOffset1 = 0;
                        horizontalAxisOffset2 = 0;
                        horizontalOriginOffset1 = 0;
                        horizontalOriginOffset2 = 0;
                    }
                    if (Configuration == 3 || Configuration == 4)
                    {
                        hoizontalPlaneOffset1 = 0;
                        horizontalPlaneOffset2 = 0;
                        horizontalAxisOffset1 = sectionDepth;
                        horizontalAxisOffset2 = sectionDepth;
                        horizontalOriginOffset1 = 0;
                        horizontalOriginOffset2 = 0;
                    }

                    if (Configuration == 1)
                    {
                        configIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        configIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.X);
                        routeUboltIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);

                        structAxisOffset = -sectionWidth / 2;
                        structOriginOffset = -overhangLeft;

                        horizontalPlaneOffset = pipeDiameter / 2 + padThickness;
                        horizontalAxisOffset = -overhangLeft;
                        horizontalOriginOffset = -sectionWidth / 2;
                    }
                    else if (Configuration == 2)
                    {
                        configIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.ZX, Axis.Z, Axis.X);
                        configIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX);
                        routeUboltIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);

                        structAxisOffset = sectionWidth / 2;
                        structOriginOffset = -overhangRight;

                        horizontalPlaneOffset = pipeDiameter / 2 + padThickness;
                        horizontalAxisOffset = overhangRight;
                        horizontalOriginOffset = sectionWidth / 2;
                    }
                    else if (Configuration == 3)
                    {
                        configIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.YZ, Axis.Y, Axis.Y);
                        configIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeYZ, Axis.Y, Axis.Y);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeY);
                        routeUboltIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);

                        structAxisOffset = sectionWidth / 2;
                        structOriginOffset = -overhangRight;

                        horizontalPlaneOffset = -pipeDiameter / 2 - padThickness;
                        horizontalAxisOffset = overhangRight;
                        horizontalOriginOffset = sectionWidth / 2;
                    }
                    else if (Configuration == 4)
                    {
                        configIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeZX, Axis.Y, Axis.X);
                        configIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.Y, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y);
                        routeUboltIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);

                        structAxisOffset = -sectionWidth / 2;
                        structOriginOffset = -overhangLeft;

                        horizontalPlaneOffset = -pipeDiameter / 2 - padThickness;
                        horizontalAxisOffset = -overhangLeft;
                        horizontalOriginOffset = -sectionWidth / 2;
                    }
                }
                else if (connection.PropValue == 2)
                {
                    (componentDictionary[VERTSECTION1]).SetPropertyValue(heightOffset, "IJUAHgrOccOverLength", "BeginOverLength");
                    (componentDictionary[VERTSECTION2]).SetPropertyValue(heightOffset, "IJUAHgrOccOverLength", "EndOverLength");


                    if (Configuration == 1)
                    {
                        configIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                        configIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.NegativeX);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY);
                        routeUboltIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);

                        structAxisOffset = sectionWidth / 2;
                        structOriginOffset = -overhangLeft;

                        hoizontalPlaneOffset1 = 0;
                        horizontalPlaneOffset2 = 0;
                        horizontalAxisOffset1 = -widthOffset;
                        horizontalAxisOffset2 = widthOffset;
                        horizontalOriginOffset1 = 0;
                        horizontalOriginOffset2 = 0;

                        horizontalPlaneOffset = pipeDiameter / 2 + padThickness;
                        horizontalAxisOffset = -overhangLeft;
                        horizontalOriginOffset = sectionWidth / 2;
                    }
                    else if (Configuration == 2)
                    {
                        configIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.X);
                        configIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeX);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.YZ, Axis.X, Axis.Y);
                        routeUboltIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);

                        structAxisOffset = -sectionWidth / 2;
                        structOriginOffset = -overhangRight;

                        hoizontalPlaneOffset1 = 0;
                        horizontalPlaneOffset2 = 0;
                        horizontalAxisOffset1 = 0;
                        horizontalAxisOffset2 = 0;
                        horizontalOriginOffset1 = widthOffset;
                        horizontalOriginOffset2 = -widthOffset;

                        horizontalPlaneOffset = pipeDiameter / 2 + padThickness;
                        horizontalAxisOffset = overhangRight;
                        horizontalOriginOffset = -sectionWidth / 2;
                    }
                    else if (Configuration == 3)
                    {
                        configIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY);
                        configIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeY);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeZX, Axis.X, Axis.X);
                        routeUboltIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);

                        structAxisOffset = -sectionWidth / 2;
                        structOriginOffset = -overhangRight;

                        hoizontalPlaneOffset1 = widthOffset;
                        horizontalPlaneOffset2 = -widthOffset;
                        horizontalAxisOffset1 = sectionDepth;
                        horizontalAxisOffset2 = sectionDepth;
                        horizontalOriginOffset1 = 0;
                        horizontalOriginOffset2 = 0;

                        horizontalPlaneOffset = -pipeDiameter / 2 - padThickness;
                        horizontalAxisOffset = overhangRight;
                        horizontalOriginOffset = -sectionWidth / 2;
                    }
                    else if (Configuration == 4)
                    {
                        configIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeYZ, Axis.Z, Axis.NegativeY);
                        configIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeYZ, Axis.Z, Axis.Y);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX);
                        routeUboltIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);

                        hoizontalPlaneOffset1 = 0;
                        horizontalPlaneOffset2 = 0;
                        horizontalAxisOffset1 = sectionDepth;
                        horizontalAxisOffset2 = sectionDepth;
                        horizontalOriginOffset1 = -widthOffset;
                        horizontalOriginOffset2 = widthOffset;

                        structAxisOffset = sectionWidth / 2;
                        structOriginOffset = -overhangLeft;

                        horizontalPlaneOffset = -pipeDiameter / 2 - padThickness;
                        horizontalAxisOffset = -overhangLeft;
                        horizontalOriginOffset = sectionWidth / 2;
                    }
                }
                if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                {
                    //Add Joint Between the Horizontal and Vertical Beams
                    JointHelper.CreateRigidJoint(HORSECTION, horizontalPort1, VERTSECTION1, verticalPort1, configIndex1.A, configIndex1.B, configIndex1.C, configIndex1.D, hoizontalPlaneOffset1, horizontalAxisOffset1, horizontalOriginOffset1);
                    //Add Joint Between the Horizontal and Vertical Beams
                    JointHelper.CreateRigidJoint(HORSECTION, horizontalPort2, VERTSECTION2, verticalPort2, configIndex2.A, configIndex2.B, configIndex2.C, configIndex2.D, horizontalPlaneOffset2, horizontalAxisOffset2, horizontalOriginOffset2);
                }
                else if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                {
                    if (Configuration == 1 || Configuration == 2)
                    {
                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HORSECTION, horizontalPort1, VERTSECTION1, verticalPort1, configIndex1.A, configIndex1.B, configIndex1.C, configIndex1.D, hoizontalPlaneOffset1, horizontalAxisOffset1, horizontalOriginOffset1);
                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HORSECTION, horizontalPort2, VERTSECTION2, verticalPort2, configIndex2.A, configIndex2.B, configIndex2.C, configIndex2.D, horizontalPlaneOffset2, horizontalAxisOffset2, horizontalOriginOffset2);
                    }
                    else if (Configuration == 3 || Configuration == 4)
                    {
                        //Add joints between Route and Beam
                        JointHelper.CreateRigidJoint(CONNOBJ1, "Connection", VERTSECTION1, verticalPort1, configIndex1.A, configIndex1.B, configIndex1.C, configIndex1.D, hoizontalPlaneOffset1, horizontalAxisOffset1, horizontalOriginOffset1);
                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HORSECTION, horizontalPort1, CONNOBJ1, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        //Add joints between Route and Beam
                        JointHelper.CreateRigidJoint(CONNOBJ2, "Connection", VERTSECTION2, verticalPort2, configIndex2.A, configIndex2.B, configIndex2.C, configIndex2.D, horizontalPlaneOffset2, horizontalAxisOffset2, horizontalOriginOffset2);
                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HORSECTION, horizontalPort2, CONNOBJ2, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }
                }
                if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                {
                    if ((componentDictionary[UBOLT]).SupportsInterface("IJOAHgrGenericUBoltE"))
                        (componentDictionary[UBOLT]).SetPropertyValue(0.0, "IJOAHgrGenericUBoltE", "Angle");
                    else if ((componentDictionary[UBOLT]).SupportsInterface("IJOAHgrGenericUBolt"))
                        (componentDictionary[UBOLT]).SetPropertyValue(0.0, "IJOAHgrGenericUBolt", "Angle");
                }
                else if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                {
                    if (Configuration == 1 || Configuration == 2)
                    {
                        if ((componentDictionary[UBOLT]).SupportsInterface("IJOAHgrGenericUBoltE"))
                            (componentDictionary[UBOLT]).SetPropertyValue(slope, "IJOAHgrGenericUBoltE", "Angle");
                        else if ((componentDictionary[UBOLT]).SupportsInterface("IJOAHgrGenericUBolt"))
                            (componentDictionary[UBOLT]).SetPropertyValue(slope, "IJOAHgrGenericUBolt", "Angle");
                    }
                    else if (Configuration == 3 || Configuration == 4)
                    {
                        if ((componentDictionary[UBOLT]).SupportsInterface("IJOAHgrGenericUBoltE"))
                            (componentDictionary[UBOLT]).SetPropertyValue(-slope, "IJOAHgrGenericUBoltE", "Angle");
                        else if ((componentDictionary[UBOLT]).SupportsInterface("IJOAHgrGenericUBolt"))
                            (componentDictionary[UBOLT]).SetPropertyValue(-slope, "IJOAHgrGenericUBolt", "Angle");
                    }
                }
                //Add Joint Between the UBolt and Route
                JointHelper.CreateRigidJoint(UBOLT, "Route", "-1", "Route", routeUboltIndex.A, routeUboltIndex.B, routeUboltIndex.C, routeUboltIndex.D, 0, 0, 0);

                if (leftPad.PropValue == 2 && rightPad.PropValue == 1)
                {
                    padThickness = (double)((PropertyValueDouble)(componentDictionary[RIGHTPAD]).GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                    (componentDictionary[VERTSECTION1]).SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    //Add Joint Between the Left and the Vertical Section 1
                    JointHelper.CreateRigidJoint(RIGHTPAD, "HgrPort_2", VERTSECTION1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                }
                else if (leftPad.PropValue == 1 && rightPad.PropValue == 2)
                {
                    padThickness = (double)((PropertyValueDouble)(componentDictionary[LEFTPAD]).GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                    (componentDictionary[VERTSECTION2]).SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(LEFTPAD, "HgrPort_1", VERTSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                }

                else if (leftPad.PropValue == 1 && rightPad.PropValue == 1)
                {
                    padThickness = (double)((PropertyValueDouble)(componentDictionary[RIGHTPAD]).GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                    (componentDictionary[VERTSECTION1]).SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    (componentDictionary[VERTSECTION2]).SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "BeginOverLength");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(LEFTPAD, "HgrPort_2", VERTSECTION1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(RIGHTPAD, "HgrPort_1", VERTSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                }


                if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    JointHelper.CreateRigidJoint("-1", "Route", HORSECTION, "EndCap", routeConfigIndex.A, routeConfigIndex.B, routeConfigIndex.C, routeConfigIndex.D, horizontalPlaneOffset, horizontalAxisOffset, horizontalOriginOffset);
                else if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                {
                    JointHelper.CreateRigidJoint("-1", "Structure", VERTSECTION1, "EndCap", structConfigIndex.A, structConfigIndex.B, structConfigIndex.C, structConfigIndex.D, 0, structAxisOffset, structOriginOffset);
                    JointHelper.CreatePlanarJoint("-1", "Structure", VERTSECTION2, "BeginCap", Plane.XY, Plane.XY, 0);
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
                return 4;
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

                    routeConnections.Add(new ConnectionInfo(HORSECTION, 1));       //partindex, routeindex

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
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    structConnections.Add(new ConnectionInfo(VERTSECTION1, 1));      //partindex, routeindex

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

        #region ICustomHgrBOMDescription Members
        //------------------------------------------------------------
        //BOM Description
        //-------------------------------------------------------------
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                string bomDescription = (string)((PropertyValueString)support.GetPropertyValue("IJOAHgrGenAssyBOMDesc", "BOM_DESC")).PropValue;

                if (bomDescription == "")
                    bomDescription = "Assembly Type A for Single Pipe";

                return bomDescription;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOMDescription." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        #endregion
    }
}