//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   UFrameTypeL.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeL
//   Author       : Mahanth
//   Creation Date: 19-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 19-Sep-2013     Mahanth CR-CP-224494  Convert HS_Generic_Assy to C# .Net 
// 31-Oct-2014     PVK     TR-CP-260301	Resolve coverity issues found in August 2014 report
// 22-Jan-2015     PVK     TR-CP-264951  Resolve coverity issues found in November 2014 report
// 29-Apr-2016     PVK	   TR-CP-292882	Resolve the issues found in Generic Assemblies
// 25-May-2016     PVK	   TR-CP-292882	Resolve the issues found in Generic Assemblies
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
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

    public class UFrameTypeL : CustomSupportDefinition
    {
        private const string HORSECTION = "HORSECTION";
        private const string VERTSECTION = "VERTSECTION";
        private const string UBOLT = "UBOLT";
        private const string CONNECTIONOBJECT1 = "CONNECTIONOBJECT1";
        private const string BRACE = "BRACE";
        private const string CONNECTIONOBJECT = "CONNECTIONOBJECT";
        private const string SECPAD = "SECPAD";
        private int numRoutes, indexUBoltBegin, indexUBoltEnd, indexUBolt, indexRoute, fromOrientCodeListVlaue, connection, sectionFromRule, showBrace, overHangOption, includePad;
        private string sectionSize, fromOrient;
        private double overhangLeft, overhangRight, braceAngle;
        Collection<object> colllection;
        private double[] pipeDiameter, pipeOWoInsulation;
        private string[] boltPart, uBoltPartKey;
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

                    showBrace = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrace", "ShowBrace")).PropValue;
                    overHangOption = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyOHOpt", "OverhangOpt")).PropValue;
                    sectionFromRule = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionFromRule")).PropValue;
                    includePad = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyPad", "IncludePad")).PropValue;
                    overhangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyOHLeft", "OverhangLeft")).PropValue;
                    overhangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyOHRight", "OverhangRight")).PropValue;
                    connection = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyConn", "Connection")).PropValue;
                    fromOrientCodeListVlaue = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyFrmOrient", "FrmOrient")).PropValue;
                    MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                    fromOrient = metadataManager.GetCodelistInfo("GenAssyHgrFrmOrient", "UDP").GetCodelistItem(fromOrientCodeListVlaue).DisplayName;
                    PropertyValueCodelist braceAngleCodeList = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrace", "BraceAngle"));
                    braceAngle = Convert.ToDouble(braceAngleCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(braceAngleCodeList.PropValue).ShortDisplayName);
                    //Get Route object
                    numRoutes = SupportHelper.SupportedObjects.Count;
                    Array.Resize(ref pipeDiameter, numRoutes);
                    for (indexRoute = 0; indexRoute < numRoutes; indexRoute++)
                    {
                        PipeObjectInfo pipe = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(indexRoute + 1);
                        pipeDiameter[indexRoute] = pipe.NominalDiameter.Size;
                    }

                    //Get the Section Size
                    if (sectionFromRule == 1)//means True
                    {
                        colllection = new Collection<object>();
                        GenericHelper.GetDataByRule("HgrSectionSize", (BusinessObject)support, out colllection);
                        if (colllection != null)
                        {
                            if (colllection[0] == null)
                                sectionSize = (string)colllection[1];
                            else
                                sectionSize = (string)colllection[0];
                        }
                    }
                    else
                    {
                        PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionSize");
                        sectionSize = sectionSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodelist.PropValue).ShortDisplayName;
                        if (sectionSize.ToUpper() == "NONE" || string.IsNullOrEmpty(sectionSize))
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "UFrameTypeC.cs", 117);
                    }
                    //Get the UBolt
                    Array.Resize(ref boltPart, numRoutes);
                    colllection = new Collection<object>();
                    GenericHelper.GetDataByRule("HgrUBoltSelection", (BusinessObject)support, out colllection);
                    if (colllection != null)
                    {
                        for (int indexUBoltPart = 0; indexUBoltPart < numRoutes; indexUBoltPart++)
                        {
                            if (indexUBoltPart == 0 && colllection[0] == null)
                                boltPart[indexUBoltPart] = (string)colllection[indexUBoltPart + 1];
                            else
                                boltPart[indexUBoltPart] = (string)colllection[indexUBoltPart];
                        }
                    }

                    //Get the Angle Pad
                    string anglePadPart = GenericAssemblyServices.GetDataByConditionString("GenServ_AnglePadDim", "IJUAHgrGenServAnglePad", "PadPart", "IJUAHgrGenServAnglePad", "SectionSize", sectionSize);

                    //Get the Hgr Overhang from the catalog. Check this value with the value on assembly. If
                    //both are not equal, it means that the user modified it. Then use the value given by user
                    double catHgrLeftOH = (double)((PropertyValueDouble)support.SupportDefinition.GetPropertyValue("IJOAHgrGenAssyOHLeft", "OverhangLeft")).PropValue;
                    double catHgrRightOH = (double)((PropertyValueDouble)support.SupportDefinition.GetPropertyValue("IJOAHgrGenAssyOHRight", "OverhangRight")).PropValue;

                    colllection = new Collection<object>();
                    GenericHelper.GetDataByRule("HgrOverhang", (BusinessObject)support, out colllection);
                    if (HgrCompareDoubleService.cmpdbl(overhangLeft, catHgrLeftOH) == true)
                    {
                        if (colllection != null)
                        {
                            if (colllection[0] == null)
                                overhangLeft = (double)colllection[1];
                            else
                                overhangLeft = (double)colllection[0];
                        }
                    }
                    if (HgrCompareDoubleService.cmpdbl(overhangRight, catHgrRightOH) == true)
                    {
                        if (colllection != null)
                        {
                            if (colllection[0] == null)
                                overhangRight = (double)colllection[2];
                            else
                                overhangRight = (double)colllection[1];
                        }
                    }
                    //Create List Of Parts

                    parts.Add(new PartInfo(HORSECTION, sectionSize));
                    parts.Add(new PartInfo(VERTSECTION, sectionSize));

                    indexUBoltBegin = 3;
                    indexUBoltEnd = numRoutes + 2;
                    for (indexUBolt = indexUBoltBegin; indexUBolt <= indexUBoltEnd; indexUBolt++)
                    {
                        if (indexUBolt == indexUBoltBegin)
                            indexRoute = 1;
                        else
                            indexRoute = indexUBolt - indexUBoltBegin + 1;
                        Array.Resize(ref uBoltPartKey, indexUBolt);
                        uBoltPartKey[indexUBolt - 1] = "UBolt" + indexUBolt;
                        parts.Add(new PartInfo(uBoltPartKey[indexUBolt - 1], boltPart[indexRoute - 1]));
                    }

                    if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                        parts.Add(new PartInfo(CONNECTIONOBJECT1, "Log_Conn_Part_1"));
                    if (includePad == 1)
                        parts.Add(new PartInfo(SECPAD, anglePadPart));
                    if (showBrace == 1)
                    {
                        parts.Add(new PartInfo(BRACE, "HgrGen_GenericBrace_1"));
                        parts.Add(new PartInfo(CONNECTIONOBJECT, "Log_Conn_Part_1"));
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
                int indexText;
                for (int index = indexUBoltBegin; index <= indexUBoltEnd; index++)
                {
                    if (index == indexUBoltBegin)
                        indexText = 1;
                    else
                        indexText = index - indexUBoltBegin + 1;
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim" + indexText, componentDictionary[uBoltPartKey[index - 1]], "Route");

                }
                indexText = numRoutes + 1;
                GenericAssemblyServices.CreateDimensionNote(this, "Dim" + indexText, componentDictionary[HORSECTION], "BeginCap");
                indexText = indexText + 1;
                GenericAssemblyServices.CreateDimensionNote(this, "Dim" + indexText, componentDictionary[HORSECTION], "EndCap");
                indexText = indexText + 1;
                GenericAssemblyServices.CreateDimensionNote(this, "Dim" + indexText, componentDictionary[VERTSECTION], "EndCap");
                if (showBrace == 1)
                {
                    indexText = indexText + 1;
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim" + indexText, componentDictionary[BRACE], "BeginCap");
                    indexText = indexText + 1;
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim" + indexText, componentDictionary[BRACE], "EndCap");
                }

                BusinessObject sectionPart = componentDictionary[HORSECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                double sectionWidth = 0.0, sectionDepth = 0.0, sectionThickness = 0.0;
                sectionWidth = crossSection.Width;
                sectionDepth = crossSection.Depth;
                sectionThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                //Get the largest Pipe Dia

                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                Vector vector = routeInfo.Orientation;
                double distance = Math.Sqrt(vector.X * vector.X + vector.Y * vector.Y), slope;
                if (distance < Math3d.DistanceTolerance)
                    slope = Math.PI / 2;
                else
                    slope = Math.Atan(Math.Abs(vector.Z) / vector.Length);

                double pipeAngle = (double)GenericAssemblyServices.GetRouteStructConfigAngle(this, "Route", "Structure", PortAxisType.Z);

                if (double.IsNaN(pipeAngle))
                    pipeAngle = 0;
                double angledOffset = Math.Tan(pipeAngle) * sectionWidth;

                double extraLength = 0;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                        extraLength = angledOffset;
                }
                //Get the distance between Extreme Left and Extreme Right Pipes
                double routeStructAngle = (double)RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                double extremePipesDistance;

                if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && Math.Abs(routeStructAngle) > (Math.PI - 0.001)))  //for pipes which are horizontal and placed vertically
                    extremePipesDistance = GenericAssemblyServices.GetPipesMaximumDistance(this, PortDistanceType.Vertical);
                else
                    extremePipesDistance = GenericAssemblyServices.GetPipesMaximumDistance(this, PortDistanceType.Horizontal);

                //Check if overhangs are small
                double hangerOverHangLeft=0, hangerOverHangRight=0;
                colllection = new Collection<object>();
                GenericHelper.GetDataByRule("HgrOverhang", (BusinessObject)support, out colllection);      //For Straight sections, A is the OH dimension
                if (colllection != null)
                {
                    if (colllection[0] == null)
                    {
                        hangerOverHangLeft = (double)colllection[1];
                        hangerOverHangRight = (double)colllection[1];
                    }
                    else
                    {
                        hangerOverHangLeft = (double)colllection[0];
                        hangerOverHangRight = (double)colllection[0];
                    }
                }
                if (overHangOption == 3)
                {
                    if (!(overhangLeft >= hangerOverHangLeft - Math3d.DistanceTolerance) && (overhangLeft <= hangerOverHangLeft + Math3d.DistanceTolerance))
                    {
                        if (overhangLeft < hangerOverHangLeft + Math3d.DistanceTolerance)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Joints" + ": " + "WARNING: " + "Left overhang is too small. Please check.", "", "UFrameTypeJ.cs", 299);
                    }

                    if (!(overhangRight >= hangerOverHangRight - Math3d.DistanceTolerance) && (overhangRight <= hangerOverHangRight + Math3d.DistanceTolerance))
                    {
                        if (overhangRight < hangerOverHangRight + Math3d.DistanceTolerance)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Joints" + ": " + "WARNING: " + "Left overhang is too small. Please check.", "", "UFrameTypeJ.cs", 305);
                    }
                }
                //Apply the overhang as per the selected overhang option
                if (overHangOption == 1)
                {
                    colllection = new Collection<object>();
                    GenericHelper.GetDataByRule("HgrOverhang", (BusinessObject)support, out colllection);      //For Straight sections, A is the OH dimension
                    if (colllection != null)
                    {
                        if (colllection[0] == null)
                        {
                            overhangRight = (double)colllection[1];
                            overhangLeft = (double)colllection[1];
                        }
                        else
                        {
                            overhangRight = (double)colllection[0];
                            overhangLeft = (double)colllection[0];
                        }
                    }
                }
                else if (overHangOption == 2)
                {
                    if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                    {
                        double distanceRouteLeftStruct = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                        overhangLeft = distanceRouteLeftStruct + sectionWidth / 2;
                    }
                }


                //========================================
                // Set Values of Part Occurance Attributes
                //========================================               
                PropertyValueCodelist topbeginMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (topbeginMiterCodelist.PropValue == -1)
                    topbeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist topendMiterCodelist = (PropertyValueCodelist)(componentDictionary[VERTSECTION]).GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (topendMiterCodelist.PropValue == -1)
                    topendMiterCodelist.PropValue = 1;
                componentDictionary[VERTSECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERTSECTION].SetPropertyValue(topbeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERTSECTION].SetPropertyValue(topendMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");


                //For UBolt
                Collection<object> uBoltPartCollection;
                GenericHelper.GetDataByRule("HgrUBoltType", (BusinessObject)support, out uBoltPartCollection);
                string uBoltType;
                double padThicknessValue = 0;
                if (uBoltPartCollection != null)
                {
                    for (indexUBolt = indexUBoltBegin; indexUBolt <= indexUBoltEnd; indexUBolt++)
                    {
                        if (indexUBolt == indexUBoltBegin)
                            indexRoute = 0;
                        else
                            indexRoute = indexUBolt - indexUBoltBegin;

                        //Check Insulation
                        Array.Resize(ref pipeOWoInsulation, indexRoute + 1);
                        PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(indexRoute + 1);
                        pipeOWoInsulation[indexRoute] = pipeInfo.OutsideDiameter;

                        //Get the UBolt Type
                        //Based on the type, we need to change the vertical offset to include Pad for UBolt Type E
                        if (uBoltPartCollection[indexRoute] == null)
                            uBoltType = (string)uBoltPartCollection[indexRoute + 1];
                        else
                            uBoltType = (string)uBoltPartCollection[indexRoute];

                        if (uBoltType == "E")
                        {
                            padThicknessValue = 0.005; //As per email from China, Pad thickness = 5 mm for UBolt Type E,
                            componentDictionary[uBoltPartKey[indexUBolt - 1]].SetPropertyValue(sectionWidth, "IJOAHgrGenUBoltPadL", "PadL");
                        }
                        componentDictionary[uBoltPartKey[indexUBolt - 1]].SetPropertyValue(sectionThickness, "IJOAHgrGenUBoltFlgThick", "FlgThick");

                    }
                }
                //Set properties on Assembly
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                int size = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sectionSize).Value;
                support.SetPropertyValue(size, "IJOAHgrGenAssySection", "SectionSize");  //DI#129319
                support.SetPropertyValue(overhangLeft, "IJOAHgrGenAssyOHLeft", "OverhangLeft");    //DI#129319
                support.SetPropertyValue(overhangRight, "IJOAHgrGenAssyOHRight", "OverhangRight"); //DI#129319

                //=============
                //Create Joints
                //=============
                //Create a collection to hold the joints
                double distanceRouteStruct;
                int frameOrientation = 1;  // default case 
                if (Math.Abs(routeStructAngle) < (0 - slope + 0.001) && Math.Abs(routeStructAngle) > (0 - slope - 0.001) || (Math.Abs(routeStructAngle) < ((Math.PI - slope) + 0.001) && Math.Abs(routeStructAngle) > ((Math.PI - slope) - 0.001)))
                    distanceRouteStruct = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                else
                    distanceRouteStruct = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                {
                    if (HgrCompareDoubleService.cmpdbl(Math.Abs(slope), 0) == true)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Joints" + ": " + "WARNING: " + "This option is valid only for Sloped Route. Resetting the value.", "", "UFrameTypeL.cs", 389);
                        support.SetPropertyValue(frameOrientation, "IJOAHgrGenAssyFrmOrient", "FrmOrient"); 
                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Joints" + ": " + "WARNING: " + "This option is not yet implemented. Resetting the value.", "", "UFrameTypeL.cs", 392);
                            support.SetPropertyValue(frameOrientation, "IJOAHgrGenAssyFrmOrient", "FrmOrient"); 
                        }
                    }
                }
                double horizontalSectionLength = overhangLeft + extremePipesDistance + overhangRight;
                double braceHTOffFactor, connectionHorSecPlaneOffset, connectionOriginOffset = 0, verticalSectionlenth = 0;

                if ((Configuration == 1) || (Configuration == 2))
                    verticalSectionlenth = distanceRouteStruct - pipeOWoInsulation[0] / 2 - padThicknessValue + extraLength;
                else if ((Configuration == 3) || (Configuration == 4))
                    verticalSectionlenth = distanceRouteStruct + pipeOWoInsulation[0] / 2 + padThicknessValue + extraLength;

                if (showBrace == 1)
                {
                    double angleRadius = braceAngle / 180 * Math.PI;
                    GenericHelper.GetDataByRule("HgrBraceHeightOffset", (BusinessObject)support, out braceHTOffFactor);
                    string sectionPort1 = string.Empty, bracePort2 = string.Empty;
                    double braceLength = ((1 - braceHTOffFactor) * horizontalSectionLength / Math.Cos(angleRadius));
                    double padOffset = braceLength * Math.Sin(angleRadius);
                    double braceHtOffset = braceHTOffFactor * horizontalSectionLength;
                    GenericAssemblyServices.ConfigIndex braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(), braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex();

                    if (Configuration == 1 || Configuration == 2)
                    {
                        if (connection == 1)
                        {
                            braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                            braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.NegativeX);
                            sectionPort1 = "BeginCap";
                            bracePort2 = "EndCap";
                            connectionOriginOffset = sectionThickness;

                        }
                        else if (connection == 2)
                        {
                            braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                            braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeX);
                            sectionPort1 = "BeginCap";
                            bracePort2 = "BeginCap";
                        }
                    }
                    else if (Configuration == 3 || Configuration == 4)
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Joints" + ": " + "WARNING: " + "Brace is invalid for this configuration.", "", "UFrameTypeL.cs", 305);
                    connectionHorSecPlaneOffset = -braceHtOffset;

                    PropertyValueCodelist braceOrientCodeListValue = (PropertyValueCodelist)componentDictionary[BRACE].GetPropertyValue("IJOAHgrGenericBrace", "BraceOrient");
                    if (braceOrientCodeListValue.PropValue == -1)
                        braceOrientCodeListValue.PropValue = 1;
                    componentDictionary[BRACE].SetPropertyValue(braceLength, "IJOAHgrGenericBrace", "L");
                    componentDictionary[BRACE].SetPropertyValue(angleRadius, "IJOAHgrGenericBrace", "Angle");
                    componentDictionary[BRACE].SetPropertyValue(braceOrientCodeListValue.PropValue = 2, "IJOAHgrGenericBrace", "BraceOrient");
                    componentDictionary[BRACE].SetPropertyValue(sectionWidth, "IJOAHgrGenericBrace", "W");
                    componentDictionary[BRACE].SetPropertyValue(sectionDepth, "IJOAHgrGenericBrace", "Depth");
                    componentDictionary[BRACE].SetPropertyValue(sectionThickness, "IJOAHgrGenericBrace", "FlangeT");

                    //Add a Joint between horizontal section and the point that we want the angle to rotate around.
                    JointHelper.CreateRigidJoint(CONNECTIONOBJECT, "Connection", HORSECTION, sectionPort1, braceConfigIndex1.A, braceConfigIndex1.B, braceConfigIndex1.C, braceConfigIndex1.D, connectionHorSecPlaneOffset, 0, connectionOriginOffset);
                    JointHelper.CreateRigidJoint(CONNECTIONOBJECT, "Connection", BRACE, bracePort2, braceConfigIndex2.A, braceConfigIndex2.B, braceConfigIndex2.C, braceConfigIndex2.D, 0, 0, 0);

                }
                componentDictionary[HORSECTION].SetPropertyValue(horizontalSectionLength, "IJUAHgrOccLength", "Length");
                componentDictionary[VERTSECTION].SetPropertyValue(verticalSectionlenth, "IJUAHgrOccLength", "Length");
                double padThickness = 0;
                string padPort = "HgrPort_2";
                if (includePad == 1)
                {
                    padThickness = (double)((PropertyValueDouble)componentDictionary[SECPAD].GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                    componentDictionary[VERTSECTION].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "EndOverLength");
                    JointHelper.CreateRigidJoint(SECPAD, padPort, VERTSECTION, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);

                }
                string referencePortName = string.Empty;
                GenericAssemblyServices.ConfigIndex horizontalConfigIndex = new GenericAssemblyServices.ConfigIndex(), routeConfigIndex = new GenericAssemblyServices.ConfigIndex(), structConfigIndex = new GenericAssemblyServices.ConfigIndex();
                double structAxisOffset = 0, structtOriginOffset = 0, horizontalPlaneOffset = 0, routePlaneOffset = 0, routeAxisOffset = 0, routeOriginOffset = 0;
                string horizontalPort1 = "EndCap", horizontalPort2 = "BeginCap", verticalPort1 = "EndCap", verticalPort2 = "BeginCap";
                if (connection == 1)
                {
                    if (Configuration == 1)
                    {
                        horizontalConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.X);
                        horizontalPlaneOffset = 0;
                        structAxisOffset = -sectionWidth / 2;
                        structtOriginOffset = -overhangLeft;
                        routePlaneOffset = pipeOWoInsulation[0] / 2 + padThicknessValue;
                        routeAxisOffset = extremePipesDistance + overhangRight;
                        routeOriginOffset = -sectionWidth / 2;
                    }
                    else if (Configuration == 2)
                    {
                        horizontalConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX);
                        horizontalPlaneOffset = 0;
                        structAxisOffset = sectionWidth / 2;
                        structtOriginOffset = extremePipesDistance + overhangRight;
                        routePlaneOffset = pipeOWoInsulation[0] / 2 + padThicknessValue;
                        routeAxisOffset = -overhangLeft;
                        routeOriginOffset = sectionWidth / 2;
                    }
                    else if (Configuration == 3)
                    {
                        horizontalConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX);
                        horizontalPlaneOffset = sectionDepth;
                        structAxisOffset = sectionWidth / 2;
                        structtOriginOffset = -overhangLeft;
                        routePlaneOffset = -pipeOWoInsulation[0] / 2 - padThicknessValue;
                        routeAxisOffset = extremePipesDistance + overhangRight;
                        routeOriginOffset = sectionWidth / 2;
                    }
                    else if (Configuration == 4)
                    {
                        horizontalConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeZX, Axis.X, Axis.X);
                        horizontalPlaneOffset = sectionDepth;
                        structAxisOffset = -sectionWidth / 2;
                        structtOriginOffset = extremePipesDistance + overhangRight;
                        routePlaneOffset = -pipeOWoInsulation[0] / 2 - padThicknessValue;
                        routeAxisOffset = -overhangLeft;
                        routeOriginOffset = -sectionWidth / 2;
                    }
                }
                if (connection == 2)
                {
                    if (Configuration == 1)
                    {
                        horizontalConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY);
                        horizontalPlaneOffset = 0;
                        structAxisOffset = sectionWidth / 2;
                        structtOriginOffset = -overhangLeft;
                        routePlaneOffset = pipeOWoInsulation[0] / 2 + padThicknessValue;
                        routeAxisOffset = extremePipesDistance + overhangRight;
                        routeOriginOffset = sectionWidth / 2;
                    }
                    else if (Configuration == 2)
                    {
                        horizontalConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.YZ, Axis.X, Axis.Y);
                        horizontalPlaneOffset = 0;
                        structAxisOffset = -sectionWidth / 2;
                        structtOriginOffset = extremePipesDistance + overhangRight;
                        routePlaneOffset = pipeOWoInsulation[0] / 2 + padThicknessValue;
                        routeAxisOffset = -overhangLeft;
                        routeOriginOffset = -sectionWidth / 2;
                    }
                    else if (Configuration == 3)
                    {
                        horizontalConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.X, Axis.NegativeY);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeY);
                        horizontalPlaneOffset = sectionDepth;
                        structAxisOffset = -sectionWidth / 2;
                        structtOriginOffset = -overhangLeft;
                        routePlaneOffset = -pipeOWoInsulation[0] / 2 - padThicknessValue;
                        routeAxisOffset = extremePipesDistance + overhangRight;
                        routeOriginOffset = -sectionWidth / 2;
                    }
                    else if (Configuration == 4)
                    {
                        horizontalConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.X, Axis.Y);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeY);
                        horizontalPlaneOffset = sectionDepth;
                        structAxisOffset = sectionWidth / 2;
                        structtOriginOffset = extremePipesDistance + overhangRight;
                        routePlaneOffset = -pipeOWoInsulation[0] / 2 - padThicknessValue;
                        routeAxisOffset = -overhangLeft;
                        routeOriginOffset = sectionWidth / 2;
                    }
                }
                if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    JointHelper.CreateRigidJoint(HORSECTION, horizontalPort1, VERTSECTION, verticalPort2, horizontalConfigIndex.A, horizontalConfigIndex.B, horizontalConfigIndex.C, horizontalConfigIndex.D, horizontalPlaneOffset, 0, 0);
                else if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                {
                    if ((Configuration == 1) || (Configuration == 2))
                        JointHelper.CreateRigidJoint(HORSECTION, horizontalPort1, VERTSECTION, verticalPort2, horizontalConfigIndex.A, horizontalConfigIndex.B, horizontalConfigIndex.C, horizontalConfigIndex.D, horizontalPlaneOffset, 0, 0);
                    else if ((Configuration == 3) || (Configuration == 4))
                    {
                        JointHelper.CreateRigidJoint(CONNECTIONOBJECT1, "Connection", VERTSECTION, verticalPort2, horizontalConfigIndex.A, horizontalConfigIndex.B, horizontalConfigIndex.C, horizontalConfigIndex.D, horizontalPlaneOffset, 0, 0);
                        JointHelper.CreateRigidJoint(HORSECTION, horizontalPort1, CONNECTIONOBJECT1, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }
                }

                for (indexUBolt = indexUBoltBegin; indexUBolt <= indexUBoltEnd; indexUBolt++)
                {
                    if (indexUBolt == indexUBoltBegin)
                        referencePortName = "Route";
                    else
                        referencePortName = "Route_" + (indexUBolt - indexUBoltBegin + 1);

                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                        componentDictionary[uBoltPartKey[indexUBolt - 1]].SetPropertyValue(0.0, "IJOAHgrGenericUBoltE", "Angle");
                    else if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                        if ((Configuration == 1) || (Configuration == 2))
                            componentDictionary[uBoltPartKey[indexUBolt - 1]].SetPropertyValue(slope, "IJOAHgrGenericUBoltE", "Angle");
                        else if ((Configuration == 3) || (Configuration == 4))
                            componentDictionary[uBoltPartKey[indexUBolt - 1]].SetPropertyValue(-slope, "IJOAHgrGenericUBoltE", "Angle");

                    GenericAssemblyServices.ConfigIndex routeIndex = new GenericAssemblyServices.ConfigIndex();
                    if ((Configuration == 1) || (Configuration == 2))
                        routeIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                    else if ((Configuration == 3) || (Configuration == 4))
                        routeIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);

                    //Add Joint Between the UBolt and Route '9380
                    JointHelper.CreateRigidJoint(uBoltPartKey[indexUBolt - 1], "Route", "-1", referencePortName, routeIndex.A, routeIndex.B, routeIndex.C, routeIndex.D, 0, 0, 0);
                }

                //What you need to do in here is:
                //   - If we are connecting PERPENDICULAR TO PIPE then we only connect to the pipe
                //   - If we connecting PERPENDICULAR TO THE STRUCTURE then we only connec to the structure
                if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    JointHelper.CreateRigidJoint("-1", "Route", HORSECTION, horizontalPort2, routeConfigIndex.A, routeConfigIndex.B, routeConfigIndex.C, routeConfigIndex.D, routePlaneOffset, routeAxisOffset, routeOriginOffset);
                else if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                {
                    //Add Joint Between Structure and Vertical Section 1
                    JointHelper.CreateRigidJoint("-1", "Structure", VERTSECTION, verticalPort1, structConfigIndex.A, structConfigIndex.B, structConfigIndex.C, structConfigIndex.D, 0, structAxisOffset, structtOriginOffset);

                    //Add Joint Between Structure and Vertical Section 2
                    JointHelper.CreatePlanarJoint("-1", "Structure", VERTSECTION, verticalPort1, Plane.XY, Plane.NegativeXY, 0);
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get As.sembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(HORSECTION, 1));

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


