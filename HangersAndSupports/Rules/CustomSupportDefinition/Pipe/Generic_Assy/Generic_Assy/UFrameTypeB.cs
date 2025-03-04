//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   UFrameTypeB.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeB
//   Author       : Manikanth
//   Creation Date: 17-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-Sep-2013    Manikanth   CR-CP-224494 Convert HS_Generic_Assy to C# .Net 
//   31-Oct-2014    PVK         TR-CP-260301	Resolve coverity issues found in August 2014 report
//   22-Jan-2015    PVK         TR-CP-264951    Resolve coverity issues found in November 2014 report
//   28-Apr-2015    PVK	        Resolve Coverity issues found in April
//   29-Apr-2016    PVK	        TR-CP-292882	Resolve the issues found in Generic Assemblies
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

    public class UFrameTypeB : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string HORSECTION = "HORSECTION";
        private string CANTIPAD = "CANTIPAD";
        private string BRACE = "BRACE";
        private string BRACEPAD = "BRACEPAD";
        private string CONNOBJ = "CONNOBJ";
        private double overhang, braceAngle;
        private string fromOrient, sectionFromRule, sectionSize, anglePadPart;
        private int routeCount, uBolt, connectionObject, routeIndex;
        bool value;
        string[] ubolt, connectionPart;
        private double[] pipeDiameter;
        Collection<object> colllection;

        PropertyValueCodelist showBrace, sectionFromRuleCodelist, sectionCodeList, includeCantiPad, includeBracePad, connection, fromOrientCodeList, braceAngleCodelist;
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

                    showBrace = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrace", "ShowBrace"));
                    sectionFromRuleCodelist = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionFromRule"));
                    sectionFromRule = (sectionFromRuleCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(sectionFromRuleCodelist.PropValue).DisplayName);
                    sectionCodeList = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionSize"));
                    sectionSize = sectionCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionCodeList.PropValue).DisplayName;
                    includeCantiPad = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyCantiPad", "IncludeCantiPad");
                    includeBracePad = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrPad", "IncludeBracePad");
                    overhang = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyOverhang", "Overhang")).PropValue;
                    connection = (((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyConn", "Connection")));
                    fromOrientCodeList = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyFrmOrient", "FrmOrient"));
                    fromOrient = (fromOrientCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(fromOrientCodeList.PropValue).DisplayName);
                    braceAngleCodelist = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrace", "BraceAngle"));
                    braceAngle = Convert.ToDouble(braceAngleCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(braceAngleCodelist.PropValue).DisplayName);

                    routeCount = SupportHelper.SupportedObjects.Count;
                    pipeDiameter = new double[routeCount + 1];
                    for (routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                    {
                        PipeObjectInfo pipeinfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                        pipeDiameter[routeIndex] = pipeinfo.OutsideDiameter / 2;
                    }
                    if (sectionFromRule == "True")
                        value = GenericHelper.GetDataByRule("HgrSectionSize", (BusinessObject)support, out sectionSize);
                    else if (sectionFromRule == "False")
                    {
                        PropertyValueCodelist sectionSizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionSize");
                        sectionSize = sectionSizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodeList.PropValue).DisplayName;

                        if (sectionSize.ToUpper() == "NONE" || sectionSize == string.Empty)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "UFrameTypeB.cs", 1);
                    }
                    colllection = new Collection<object>();
                    string[] uBoltPart = new string[routeCount + 1];
                    for (int i = 1; i <= routeCount; i++)
                    {
                        uBoltPart[i] = "UBOLT" + i;
                        value = GenericHelper.GetDataByRule("HgrUBoltSelection", support, out colllection);
                        if (colllection != null)
                        {
                            if (colllection[0] == null)
                                uBoltPart[i] = (string)colllection[1];
                            else
                                uBoltPart[i] = (string)colllection[0];
                        }

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
                        anglePadPart = (string)((PropertyValueString)padPart.ElementAt(0).GetPropertyValue("IJUAHgrGenServAnglePad", "PadPart")).PropValue;

                    //Get the Hgr Overhang from the catalog. Check this value with the value on assembly. If
                    //both are not equal, it means that the user modified it. Then use the value given by user
                    double catOverHang = (double)((PropertyValueDouble)support.SupportDefinition.GetPropertyValue("IJOAHgrGenAssyOverhang", "Overhang")).PropValue;
                    if (HgrCompareDoubleService.cmpdbl(catOverHang, overhang) == true)
                    {
                        value = GenericHelper.GetDataByRule("HgrOverhang", support, out colllection);
                        if (colllection != null)
                        {
                            if (colllection[0] == null)
                                overhang = (double)colllection[1];
                            else
                                overhang = (double)colllection[0];
                        }
                    }

                    parts.Add(new PartInfo(HORSECTION, sectionSize));
                    ubolt = new string[routeCount + 2];
                    for (int uBolt = 2; uBolt <= routeCount + 1; uBolt++)
                    {
                        ubolt[uBolt] = "UBOLT" + uBolt;
                        if (uBolt == 2)
                            routeIndex = 1;
                        else
                            routeIndex = uBolt - 1;
                        parts.Add(new PartInfo(ubolt[uBolt], uBoltPart[routeIndex]));
                    }
                    connectionPart = new string[(2 * routeCount + 3)];
                    for (connectionObject = routeCount + 2; connectionObject <= (2 * routeCount + 2); connectionObject++)
                    {
                        connectionPart[connectionObject] = "connection" + connectionObject;
                        if (connectionObject == (2 * routeCount + 2))
                            routeIndex = 1;
                        else
                            routeIndex = connectionObject - routeCount + 3;
                        parts.Add(new PartInfo(connectionPart[connectionObject], "Log_Conn_Part_1"));
                    }

                    if (includeCantiPad.PropValue == 1)
                        parts.Add(new PartInfo(CANTIPAD, anglePadPart));

                    if (showBrace.PropValue == 1)
                    {
                        parts.Add(new PartInfo(BRACE, "HgrGen_GenericBrace_1"));
                        parts.Add(new PartInfo(CONNOBJ, "Log_Conn_Part_1"));
                    }
                    if (includeBracePad.PropValue == 1)
                    {
                        if (showBrace.PropValue == 1)
                            parts.Add(new PartInfo(BRACEPAD, anglePadPart));
                        else if (showBrace.PropValue == 2)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Brace is not available. Resetting the value..", "", "UFrameTypeA.cs", 1);
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

                double textIndex;
                for (int i = 2; i <= routeCount + 1; i++)
                {
                    if (i == 2)
                        textIndex = 1;
                    else
                        textIndex = i - 1;
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim" + textIndex, componentDictionary[ubolt[i]], "Route");
                }
                textIndex = routeCount + 1;
                GenericAssemblyServices.CreateDimensionNote(this, "Dim" + textIndex, componentDictionary[HORSECTION], "BeginCap");
                GenericAssemblyServices.CreateDimensionNote(this, "Dim" + textIndex + 1, componentDictionary[HORSECTION], "EndCap");

                if (showBrace.PropValue == 1)
                {
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim" + textIndex + 2, componentDictionary[BRACE], "BeginCap");
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim" + textIndex + 3, componentDictionary[BRACE], "EndCap");
                }

                BusinessObject horizontalSectionPart = componentDictionary[HORSECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                double sectionWidth = 0.0, SectionDepth = 0.0, SectionThickness = 0.0;
                sectionWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                SectionDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                SectionThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                string connection;
                //Determine whether connecting to Steel or a Slab
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    connection = "Slab";
                else
                {
                    if (SupportHelper.SupportingObjects.Count > 1)
                    {
                        connection = "steel";
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                            connection = "Slab";          //Two Slabs

                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Member))
                            connection = "Slab-Steel";    //Slab then Steel

                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                            connection = "Steel-Slab";    //Steel then Slab
                    }
                    else
                    {
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                            connection = "Steel";    //Steel                      
                        else
                            connection = "Slab";
                    }
                }

                double dRouteStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                if (Math.Abs(dRouteStructAngle) < (Math.PI / 2 + 0.001) && Math.Abs(dRouteStructAngle) > (Math.PI / 2 - 0.001))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Use Frame Support Type G for this Structure and Route configuration.", "", "UFrameTypeB.cs", 1);
                    return;
                }

                double angledOffset = 0, pipeAngle = 0;
                pipeAngle = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Z);
                angledOffset = Math.Tan(pipeAngle) * sectionWidth;

                PropertyValueCodelist topBeginMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (topBeginMiterCodelist.PropValue == -1)
                    topBeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist topEndMiterCodelist = (PropertyValueCodelist)(componentDictionary[HORSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (topEndMiterCodelist.PropValue == -1)
                    topEndMiterCodelist.PropValue = 1;

                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                (componentDictionary[HORSECTION]).SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                (componentDictionary[HORSECTION]).SetPropertyValue(topBeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                (componentDictionary[HORSECTION]).SetPropertyValue(topEndMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                double[] pipeDiameter = new double[routeCount + 1]; ;
                for (int i = 1; i <= routeCount; i++)
                {
                    PipeObjectInfo pipeinfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                    pipeDiameter[i] = pipeinfo.OutsideDiameter;
                }
                string refPortName = string.Empty;
                double padThickness = 0;
                value = GenericHelper.GetDataByRule("HgrUBoltType", support, out colllection);
                for (int i = 2; i <= routeCount + 1; i++)
                {
                    if (i == 2)
                        refPortName = "Route";
                    else
                    {
                        routeIndex = uBolt - 1;
                        refPortName = "Route_" + i;
                    }
                    string uBoltType=string.Empty;
                    if (colllection != null)
                    {
                        if (colllection[0] == null)
                            uBoltType = (string)colllection[1];
                        else
                            uBoltType = (string)colllection[0];
                    }
                    if (uBoltType == "E")
                    {
                        padThickness = 0.005;
                        (componentDictionary[ubolt[i]]).SetPropertyValue(sectionWidth, "IJOAHgrGenUBoltPadL", "PadL");
                    }
                    (componentDictionary[ubolt[i]]).SetPropertyValue(SectionThickness, "IJOAHgrGenUBoltFlgThick", "FlgThick");
                }
                int frameOrientation = 1;
                if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "This option is not yet implemented. Resetting the value.", "", "UFrameTypeB.cs", 323);
                    support.SetPropertyValue(frameOrientation, "IJOAHgrGenAssyFrmOrient", "FrmOrient");
                }
                double structRouteDistance = GenericAssemblyServices.GetMaximumRouteStructDistance(this, PortDistanceType.Horizontal);
                int distanceRouteIndex = GetMaximumDistanceRouteIndex();

                double horizontalLength = structRouteDistance + overhang;
                string cantiPadPort = string.Empty, sectionPort2 = string.Empty;

                if (showBrace.PropValue == 1)
                {
                    double angle1 = 0, braceLength = 0, bracePadOffset = 0, braceHtOffset = 0, braceHtFactor = 0, angleDegree1 = 0, connectionPlaneOffset = 0, padPlaneOffset = 0, padAxisOffset = 0, padOriginOffset = 0;
                    value = GenericHelper.GetDataByRule("HgrBraceHeightOffset", support, out braceHtFactor);

                    braceLength = ((1 - braceHtFactor) * horizontalLength / Math.Cos((braceAngle * (Math.PI / 180)))) + ((sectionWidth - SectionThickness) * Math.Sin((braceAngle * (Math.PI / 180))));
                    bracePadOffset = (braceLength) * Math.Sin((braceAngle * (Math.PI / 180)));
                    braceHtOffset = braceHtFactor * horizontalLength;

                    string sectionPort1 = string.Empty, bracePort2 = string.Empty, brPadPort = string.Empty;
                    GenericAssemblyServices.ConfigIndex bracePadConfgiIndex = new GenericAssemblyServices.ConfigIndex(), bracePadConfgiIndex1 = new GenericAssemblyServices.ConfigIndex(), bracePadConfgiIndex2 = new GenericAssemblyServices.ConfigIndex();

                    if (Configuration == 1)
                    {
                        bracePadConfgiIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX);
                        bracePadConfgiIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        bracePadConfgiIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);

                        sectionPort1 = "EndCap";
                        sectionPort2 = "BeginCap";
                        bracePort2 = "EndCap";
                        brPadPort = "HgrPort_1";
                        cantiPadPort = "HgrPort_1";

                        padPlaneOffset = 0;
                        padAxisOffset = -bracePadOffset;
                        padOriginOffset = sectionWidth / 3;
                        connectionPlaneOffset = braceHtOffset;
                        angleDegree1 = braceAngle;
                    }
                    else if (Configuration == 2)
                    {
                        bracePadConfgiIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX);
                        bracePadConfgiIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y);
                        bracePadConfgiIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);

                        sectionPort1 = "EndCap";
                        sectionPort2 = "BeginCap";
                        bracePort2 = "BeginCap";
                        brPadPort = "HgrPort_1";
                        cantiPadPort = "HgrPort_1";

                        padPlaneOffset = 0;
                        padAxisOffset = -sectionWidth / 3;
                        padOriginOffset = bracePadOffset;
                        connectionPlaneOffset = -braceHtOffset;
                        angleDegree1 = braceAngle - 90;
                    }
                    else if (Configuration == 3)
                    {
                        sectionPort1 = "BeginCap";
                        sectionPort2 = "EndCap";
                        bracePort2 = "EndCap";
                        brPadPort = "HgrPort_2";
                        cantiPadPort = "HgrPort_2";
                        padPlaneOffset = 0;
                        padAxisOffset = -sectionWidth / 3;
                        padOriginOffset = bracePadOffset;
                        connectionPlaneOffset = braceHtOffset;
                        angleDegree1 = braceAngle;

                        bracePadConfgiIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.Y, Axis.NegativeX);
                        bracePadConfgiIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                        bracePadConfgiIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                    }
                    else if (Configuration == 4)
                    {
                        bracePadConfgiIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                        bracePadConfgiIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.Y, Axis.X);
                        bracePadConfgiIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                        sectionPort1 = "BeginCap";
                        sectionPort2 = "EndCap";
                        bracePort2 = "BeginCap";
                        brPadPort = "HgrPort_2";
                        cantiPadPort = "HgrPort_2";
                        padPlaneOffset = 0;
                        padAxisOffset = -bracePadOffset;
                        padOriginOffset = sectionWidth / 3;
                        connectionPlaneOffset = -braceHtOffset;
                        angleDegree1 = braceAngle - 180;
                    }
                    angle1 = angleDegree1 * (Math.PI / 180);

                    double padThicknessValue;
                    if (includeBracePad.PropValue == 1)
                    {
                        padThicknessValue = (double)((PropertyValueDouble)(componentDictionary[CANTIPAD]).GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                        //Add Joint Between the Plate and the Horizontal Section
                        JointHelper.CreateRigidJoint(BRACEPAD, brPadPort, HORSECTION, sectionPort2, bracePadConfgiIndex.A, bracePadConfgiIndex.B, bracePadConfgiIndex.C, bracePadConfgiIndex.D, padPlaneOffset, padAxisOffset, padOriginOffset);
                    }
                    (componentDictionary[BRACE]).SetPropertyValue(braceLength, "IJOAHgrGenericBrace", "L");
                    (componentDictionary[BRACE]).SetPropertyValue(angle1, "IJOAHgrGenericBrace", "Angle");
                    PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[BRACE].GetPropertyValue("IJOAHgrGenericBrace", "BraceOrient");
                    braceCodelist.PropValue = 2;
                    (componentDictionary[BRACE]).SetPropertyValue(braceCodelist.PropValue, "IJOAHgrGenericBrace", "BraceOrient");
                    (componentDictionary[BRACE]).SetPropertyValue(sectionWidth, "IJOAHgrGenericBrace", "W");
                    (componentDictionary[BRACE]).SetPropertyValue(SectionDepth, "IJOAHgrGenericBrace", "Depth");
                    (componentDictionary[BRACE]).SetPropertyValue(SectionThickness, "IJOAHgrGenericBrace", "FlangeT");
                    //Add a Joint between horizontal section and the point that we want the angle to rotate around.
                    JointHelper.CreateRigidJoint(CONNOBJ, "Connection", HORSECTION, sectionPort1, bracePadConfgiIndex1.A, bracePadConfgiIndex1.B, bracePadConfgiIndex1.C, bracePadConfgiIndex1.D, connectionPlaneOffset, 0, 0);
                    JointHelper.CreateRigidJoint(CONNOBJ, "Connection", BRACE, bracePort2, bracePadConfgiIndex2.A, bracePadConfgiIndex2.B, bracePadConfgiIndex2.C, bracePadConfgiIndex2.D, 0, 0, 0);
                }

                GenericAssemblyServices.ConfigIndex structConfigIndex = new GenericAssemblyServices.ConfigIndex(), routeConfigIndex = new GenericAssemblyServices.ConfigIndex(), uBoltHorSecConfigIndex = new GenericAssemblyServices.ConfigIndex();


                string horizontalPort = string.Empty;
                double uBoltPlaneOffset = 0, uBoltAxisOffset = 0, uBoltOriginOffset = 0;

                if (Configuration == 1)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.Y, Axis.X);
                    else
                    {
                        if (connection == "Slab")
                            structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                    }
                    routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.X);
                    uBoltHorSecConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.X);
                    horizontalPort = "EndCap";
                    uBoltPlaneOffset = -overhang;
                    uBoltAxisOffset = -pipeDiameter[distanceRouteIndex] / 2 - padThickness;
                    uBoltOriginOffset = sectionWidth / 2;
                }
                else if (Configuration == 2)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                    else
                    {
                        if (connection == "Slab")
                            structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                    }
                    routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.X);
                    uBoltHorSecConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.NegativeX);

                    horizontalPort = "EndCap";
                    uBoltPlaneOffset = -pipeDiameter[distanceRouteIndex] / 2 - padThickness;
                    uBoltAxisOffset = -overhang;
                    uBoltOriginOffset = sectionWidth / 2;
                }
                else if (Configuration == 3)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX);
                    else
                    {
                        if (connection == "Slab")
                            structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                    }
                    routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeZX, Axis.X, Axis.X);
                    uBoltHorSecConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.NegativeX);
                    horizontalPort = "BeginCap";
                    uBoltPlaneOffset = -pipeDiameter[distanceRouteIndex] / 2 - padThickness;
                    uBoltAxisOffset = overhang;
                    uBoltOriginOffset = sectionWidth / 2;
                }
                else if (Configuration == 4)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                    else
                    {
                        if (connection == "Slab")
                            structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                    }
                    routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeZX, Axis.X, Axis.X);
                    uBoltHorSecConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.X);
                    horizontalPort = "BeginCap";
                    uBoltPlaneOffset = overhang;
                    uBoltAxisOffset = -pipeDiameter[distanceRouteIndex] / 2 - padThickness;
                    uBoltOriginOffset = sectionWidth / 2;
                }
                double[] distanceArray = new double[routeCount + 2];

                for (int idxUBolt = 2; idxUBolt <= routeCount + 1; idxUBolt++)
                {
                    if (idxUBolt == 2)
                    {
                        refPortName = "Route";
                        routeIndex = 1;
                    }
                    else
                    {
                        refPortName = "Route_" + (idxUBolt - 1);
                        routeIndex = idxUBolt - 1;
                    }

                    distanceArray[routeIndex] = RefPortHelper.DistanceBetweenPorts("Structure", refPortName, PortDistanceType.Horizontal);
                }
                if (includeCantiPad.PropValue == 1)
                {
                    double padThicknessValue = (double)((PropertyValueDouble)(componentDictionary[CANTIPAD]).GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                    (componentDictionary[HORSECTION]).SetPropertyValue(-padThicknessValue, "IJUAHgrOccOverLength", "BeginOverLength");

                    JointHelper.CreateRigidJoint(CANTIPAD, cantiPadPort, HORSECTION, sectionPort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, SectionDepth / 3);
                }
                (componentDictionary[HORSECTION]).SetPropertyValue(horizontalLength, "IJUAHgrOccLength", "Length");

                double angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);
                for (int indexUBolt = 2; indexUBolt <= routeCount + 1; indexUBolt++)
                {
                    if (indexUBolt == 2)
                    {
                        refPortName = "Route";
                        routeIndex = 1;
                        connectionObject = routeCount + 2;
                    }
                    else
                    {
                        refPortName = "Route_" + (indexUBolt - 1);
                        routeIndex = indexUBolt - 1;
                        connectionObject = connectionObject + 1;
                    }

                    double routeAxisOffset = 0;

                    if (Math.Abs(angle) < Math.PI / 2.0)
                        routeAxisOffset = (-pipeDiameter[routeIndex] + pipeDiameter[distanceRouteIndex]) / 2;
                    else
                        routeAxisOffset = (pipeDiameter[routeIndex] - pipeDiameter[distanceRouteIndex]) / 2;
                    if (indexUBolt == 2)
                        //Add joint between Connection Object and the Clamp
                        JointHelper.CreateRigidJoint(connectionPart[connectionObject], "Connection", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                        //Add joint between first Cconnection Object and the Connection Objects on other Routes
                        JointHelper.CreateRigidJoint(connectionPart[connectionObject], "Connection", connectionPart[routeCount + 2], "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, distanceArray[routeIndex] - distanceArray[1], routeAxisOffset, 0);

                    //Add joint between Connection Object and the Clamp
                    JointHelper.CreateRigidJoint(ubolt[indexUBolt], "Route", connectionPart[connectionObject], "Connection", routeConfigIndex.A, routeConfigIndex.B, routeConfigIndex.C, routeConfigIndex.D, -padThickness, 0, 0);
                }
                if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                {
                    if (distanceRouteIndex == 1)
                        uBolt = 2;
                    else
                        uBolt = distanceRouteIndex + 1;

                    //Add Joint Between the Horizontal Section and Steel
                    JointHelper.CreateRigidJoint(HORSECTION, horizontalPort, ubolt[uBolt], "Route", uBoltHorSecConfigIndex.A, uBoltHorSecConfigIndex.B, uBoltHorSecConfigIndex.C, uBoltHorSecConfigIndex.D, uBoltPlaneOffset, uBoltAxisOffset, uBoltOriginOffset);
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

                    structConnections.Add(new ConnectionInfo(HORSECTION, 1));      //partindex, routeindex

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
        //--------------------------------------------------------------
        //BOM Description
        //--------------------------------------------------------------
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                string bomDescription = (string)((PropertyValueString)support.GetPropertyValue("IJOAHgrGenAssyBOMDesc", "BOM_DESC")).PropValue;

                if (bomDescription == "")
                    bomDescription = "Assembly Type B for Single or Multiple Pipes";

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
        /// <summary>
        /// Gets the Pipeindex with maximum route to structure distance.
        /// </summary>
        /// <returns>int value</returns>
        /// <example>
        /// <code>
        /// int maxDistRouteIndex = GetMaximumDistanceRouteIndex();
        /// </code>
        /// </example>   
        public int GetMaximumDistanceRouteIndex()
        {
            double[] horizontalDistance = new double[routeCount + 1];
            string pipePortName = string.Empty;
            for (int i = 1; i <= routeCount; i++)
            {
                if (i == 1)
                    pipePortName = "Route";
                else
                    pipePortName = "Route" + i;
                horizontalDistance[i] = RefPortHelper.DistanceBetweenPorts("Structure", pipePortName, PortDistanceType.Horizontal);
            }
            double maximumRouteStructDistance = 0;
            int pipeIndex = 0;
            for (int i = 1; i <= routeCount; i++)
            {
                if (horizontalDistance[i] > 0)
                    maximumRouteStructDistance = horizontalDistance[i];
            }
            for (int i = 1; i <= routeCount; i++)
            {
                if (HgrCompareDoubleService.cmpdbl(horizontalDistance[i], maximumRouteStructDistance) == true)
                    pipeIndex = i;
            }
            return pipeIndex;
        }
        /// <summary>
        /// Gets the horizontal  Distance Between structures.
        /// </summary>
        /// <returns>double value</returns>
        /// <example>
        /// <code>
        ///    double maxStructRouteDist = GetMaxRouteStructHorDistance();
        /// </code>
        /// </example>   
        public double GetMaxRouteStructHorDistance()
        {
            string pipePortName;
            double[] structhordistance = new double[routeCount + 1];
            for (int i = 1; i <= routeCount; i++)
            {
                if (i == 1)
                    pipePortName = "Route";
                else
                    pipePortName = "Route" + i;

                structhordistance[i] = RefPortHelper.DistanceBetweenPorts("Structure", pipePortName, PortDistanceType.Horizontal);
            }
            double temp;
            for (int i = 1; i <= routeCount; i++)
            {
                for (int j = i + 1; j <= routeCount; j++)
                {
                    if (structhordistance[j] > structhordistance[i])
                    {
                        temp = structhordistance[i];
                        structhordistance[i] = structhordistance[j];
                        structhordistance[j] = temp;
                    }
                }
            }
            return structhordistance[1];
        }
    }
}