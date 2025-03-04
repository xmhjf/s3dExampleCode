//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   LFieldSupports.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.UFrameTypeJ
//   Author       : Vijay
//   Creation Date: 19-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   19-Sep-2013     Vijay   CR-CP-224494 Convert HS_Generic_Assy to C# .Net 
//   31-Oct-2014     PVK     TR-CP-260301	Resolve coverity issues found in August 2014 report
//   22-Jan-2015     PVK     TR-CP-264951  Resolve coverity issues found in November 2014 report
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

    public class UFrameTypeJ : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string HORIZONTALSECTION = "HORIZONTALSECTION";
        private const string VERTICALSECTION1 = "VERTICALSECTION1";
        private const string VERTICALSECTION2 = "VERTICALSECTION2";
        private const string LEFTBRACE = "LEFTBRACE";
        private const string RIGHTBRACE = "RIGHTBRACE";
        private const string LEFTCONNOBJ = "LEFTCONNOBJ";
        private const string RIGHTCONNOBJ = "RIGHTCONNOBJ";
        private const string LEFTPAD = "LEFTPAD";
        private const string RIGHTPAD = "RIGHTPAD";
        private const string LEFTBRACEPAD = "LEFTBRACEPAD";
        private const string RIGHTBRACEPAD = "RIGHTBRACEPAD";
        private const string CONNOBJ1 = "CONNOBJ1";
        private const string CONNOBJ2 = "CONNOBJ2";
        private string sectionSize, fromOrient;
        private double braceAngle, overhangLeft, overhangRight, widthOffset, heightOffset;
        private int sectionFromRule, overhangOption, includeLeftPad, includeRightPad, showBrace, includeLeftBrPad, includeRightBrPad, connection, offsetsFromRule;
        static int routeCount = 0, routeIndex = 0;
        string[] uBoltParts = new string[routeCount];
        string[] uboltPartKeys;
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
                    sectionFromRule = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionFromRule")).PropValue;

                    overhangOption = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyOHOpt", "OverhangOpt")).PropValue;
                    includeLeftPad = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyLeftPad", "IncludeLeftPad")).PropValue;
                    includeRightPad = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyRightPad", "IncludeRightPad")).PropValue;
                    showBrace = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrace", "ShowBrace")).PropValue;
                    includeLeftBrPad = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyLeftBrPad", "IncludeLeftBrPad")).PropValue;
                    includeRightBrPad = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyRightBrPad", "IncludeRightBrPad")).PropValue;
                    overhangLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyOHLeft", "OverhangLeft")).PropValue;
                    overhangRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyOHRight", "OverhangRight")).PropValue;
                    connection = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyConn", "Connection")).PropValue;
                    offsetsFromRule = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyFrmOffsets", "OffsetsFromRule")).PropValue;
                    PropertyValueCodelist fromOrientCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyFrmOrient", "FrmOrient");
                    fromOrient = fromOrientCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(fromOrientCodelist.PropValue).ShortDisplayName;
                    PropertyValueCodelist braceAngleCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssyBrace", "BraceAngle");
                    braceAngle = Convert.ToDouble(braceAngleCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(braceAngleCodelist.PropValue).ShortDisplayName);

                    //Get the Width Offset and Height Offsets from rule
                    if (offsetsFromRule == 1)  //1 means TRUE
                    {
                        Collection<object> hangerOverHangCollection;
                        GenericHelper.GetDataByRule("HgrFrmOffsets", (BusinessObject)support, out hangerOverHangCollection);
                        if (hangerOverHangCollection != null)
                        {
                            if (hangerOverHangCollection[0] == null)
                            {
                                widthOffset = (double)hangerOverHangCollection[1];
                                heightOffset = (double)hangerOverHangCollection[2];
                            }
                            else
                            {
                                widthOffset = (double)hangerOverHangCollection[0];
                                heightOffset = (double)hangerOverHangCollection[1];
                            }
                        }
                    }
                    else
                    {
                        widthOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyFrmOffsets", "WidthOffset")).PropValue;
                        heightOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrGenAssyFrmOffsets", "HeightOffset")).PropValue;
                    }

                    //Get Route object
                    routeCount = SupportHelper.SupportedObjects.Count;

                    //Get the Section Size
                    if (sectionFromRule == 1)//means True
                    {
                        Collection<object> hangerOverHangCollection;
                        GenericHelper.GetDataByRule("HgrSectionSize", (BusinessObject)support, out hangerOverHangCollection);
                        if (hangerOverHangCollection != null)
                        {
                            if (hangerOverHangCollection[0] == null)
                                sectionSize = (string)hangerOverHangCollection[1];
                            else
                                sectionSize = (string)hangerOverHangCollection[0];
                        }
                    }
                    else
                    {
                        PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionSize");
                        sectionSize = sectionSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodelist.PropValue).ShortDisplayName;
                        if (sectionSize.ToUpper() == "NONE" || string.IsNullOrEmpty(sectionSize))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Section size is not available.", "", "UFrameTypeJ.cs", 117);
                            return parts;
                        }
                    }

                    Array.Resize(ref uBoltParts, routeCount);
                    Collection<object> uBoltPartCollection;
                    GenericHelper.GetDataByRule("HgrUBoltSelection", (BusinessObject)support, out uBoltPartCollection);
                    if (uBoltPartCollection != null)
                    {
                        for (int i = 0; i < routeCount; i++)
                        {
                            if (i == 0 && uBoltPartCollection[0] == null)
                                uBoltParts[i] = (string)uBoltPartCollection[i + 1];
                            else
                                uBoltParts[i] = (string)uBoltPartCollection[i];
                        }
                    }
                    //Get the Angle Pad
                    string anglePadPart = GenericAssemblyServices.GetDataByConditionString("GenServ_AnglePadDim", "IJUAHgrGenServAnglePad", "PadPart", "IJUAHgrGenServAnglePad", "SectionSize", sectionSize);

                    // Create list of parts
                    parts.Add(new PartInfo(HORIZONTALSECTION, sectionSize));
                    parts.Add(new PartInfo(VERTICALSECTION1, sectionSize));
                    parts.Add(new PartInfo(VERTICALSECTION2, sectionSize));

                    for (int i = 4; i <= routeCount + 3; i++)
                    {
                        if (i == 4)
                            routeIndex = 1;
                        else
                            routeIndex = i - 4 + 1;

                        Array.Resize(ref uboltPartKeys, routeIndex);
                        uboltPartKeys[routeIndex - 1] = "UBOLT" + (routeIndex - 1);
                        parts.Add(new PartInfo(uboltPartKeys[routeIndex - 1], uBoltParts[routeIndex - 1]));
                    }

                    if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                    {
                        parts.Add(new PartInfo(CONNOBJ1, "Log_Conn_Part_1"));      //Rotational Connection Object
                        parts.Add(new PartInfo(CONNOBJ2, "Log_Conn_Part_2"));      //Rotational Connection Object
                    }

                    if (includeLeftPad == 1)
                        parts.Add(new PartInfo(LEFTPAD, anglePadPart));

                    if (includeRightPad == 1)
                        parts.Add(new PartInfo(RIGHTPAD, anglePadPart));

                    if (showBrace == 1)
                    {
                        parts.Add(new PartInfo(LEFTBRACE, "HgrGen_GenericBrace_1"));
                        parts.Add(new PartInfo(LEFTCONNOBJ, "Log_Conn_Part_1")); //Rotational Connection Object
                        parts.Add(new PartInfo(RIGHTBRACE, "HgrGen_GenericBrace_1"));
                        parts.Add(new PartInfo(RIGHTCONNOBJ, "Log_Conn_Part_1")); //Rotational Connection Object
                    }

                    if (includeLeftBrPad == 1)
                    {
                        if (showBrace == 1)
                            parts.Add(new PartInfo(LEFTBRACEPAD, anglePadPart));
                        else if (showBrace == 2)
                        {
                            includeLeftBrPad = 2;
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Brace is not available. Resetting the value.", "", "UFrameTypeJ.cs", 178);
                            return parts;
                        }
                    }

                    if (includeRightBrPad == 1)
                    {
                        if (showBrace == 1)
                            parts.Add(new PartInfo(RIGHTBRACEPAD, anglePadPart));
                        else if (showBrace == 2)
                        {
                            includeRightBrPad = 2;
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Brace is not available. Resetting the value.", "", "UFrameTypeJ.cs", 189);
                            return parts;
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
                int indexText = 0;
                //Auto Dimensioning of Supports
                for (int i = 4; i <= routeCount + 3; i++)
                {
                    if (i == 4)
                        indexText = 1;
                    else
                        indexText = i - 4 + 1;

                    GenericAssemblyServices.CreateDimensionNote(this, "Dim " + indexText, componentDictionary[uboltPartKeys[indexText - 1]], "Route");
                }

                indexText++;
                GenericAssemblyServices.CreateDimensionNote(this, "Dim " + indexText, componentDictionary[HORIZONTALSECTION], "BeginCap");

                indexText++;
                GenericAssemblyServices.CreateDimensionNote(this, "Dim " + indexText, componentDictionary[HORIZONTALSECTION], "EndCap");

                indexText++;
                GenericAssemblyServices.CreateDimensionNote(this, "Dim " + indexText, componentDictionary[VERTICALSECTION1], "EndCap");

                indexText++;
                GenericAssemblyServices.CreateDimensionNote(this, "Dim " + indexText, componentDictionary[VERTICALSECTION2], "BeginCap");

                if (showBrace == 1)
                {
                    indexText++;
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim " + indexText, componentDictionary[LEFTBRACE], "BeginCap");

                    indexText++;
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim " + indexText, componentDictionary[LEFTBRACE], "EndCap");

                    indexText++;
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim " + indexText, componentDictionary[RIGHTBRACE], "BeginCap");

                    indexText++;
                    GenericAssemblyServices.CreateDimensionNote(this, "Dim " + indexText, componentDictionary[RIGHTBRACE], "EndCap");
                }

                //Get Section Structure dimensions
                BusinessObject sectionPart = componentDictionary[HORIZONTALSECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                double sectionWidth = 0.0, sectionDepth = 0.0, sectionThickness = 0.0;
                sectionWidth = crossSection.Width;
                sectionDepth = crossSection.Depth;
                sectionThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                //Width offset and Height offset should be less than 3/4th of the Section dimensions.
                //If not, Give a warning message

                if (widthOffset > 3 * sectionWidth / 4)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Joints" + ": " + "WARNING: " + "Width Offset should be less than 3/4th of the Section Width.", "", "UFrameTypeJ.cs", 263);

                if (heightOffset > 3 * sectionDepth / 4)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Joints" + ": " + "WARNING: " + "Height Offset should be less than 3/4th of the Section Height.", "", "UFrameTypeJ.cs", 266);

                //=======================================
                //Do Something if more than one Structure
                //=======================================
                //get structure count

                bool[] isOffsetApplied = GenericAssemblyServices.GetIsLugEndOffsetApplied(this);
                string[] structPort = GenericAssemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0], rightStructPort = structPort[1];

                //Get the distance between Extreme Left and Extreme Right Pipes
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);

                double extremePipesDistance;

                if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.001) && Math.Abs(routeStructAngle) > (Math.PI - 0.001)))  //for pipes which are horizontal and placed vertically
                    extremePipesDistance = GenericAssemblyServices.GetPipesMaximumDistance(this, PortDistanceType.Vertical);
                else
                    extremePipesDistance = GenericAssemblyServices.GetPipesMaximumDistance(this, PortDistanceType.Horizontal);

                //Check if overhangs are small
                Collection<object> collectionHgrOverHanger;
                double hangerOverHangLeft=0, hangerOverHangRight=0;

                GenericHelper.GetDataByRule("HgrOverhang", (BusinessObject)support, out collectionHgrOverHanger);      //For Straight sections, A is the OH dimension
                if (collectionHgrOverHanger != null)
                {
                    if (collectionHgrOverHanger[0] == null)
                        hangerOverHangLeft = (double)collectionHgrOverHanger[1];
                    else
                        hangerOverHangLeft = (double)collectionHgrOverHanger[0];
                }
                GenericHelper.GetDataByRule("HgrOverhang", (BusinessObject)support, out collectionHgrOverHanger);     //For Straight sections, A is the OH dimension
                if (collectionHgrOverHanger != null)
                {
                    if (collectionHgrOverHanger[0] == null)
                        hangerOverHangRight = (double)collectionHgrOverHanger[1];
                    else
                        hangerOverHangRight = (double)collectionHgrOverHanger[0];
                }
                if (overhangOption == 3)
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
                if (overhangOption == 1)
                {
                    GenericHelper.GetDataByRule("HgrOverhang", (BusinessObject)support, out collectionHgrOverHanger);      //For Straight sections, A is the OH dimension
                    if (collectionHgrOverHanger != null)
                    {
                        if (collectionHgrOverHanger[0] == null)
                            overhangLeft = (double)collectionHgrOverHanger[1];
                        else
                            overhangLeft = (double)collectionHgrOverHanger[0];
                    }
                    GenericHelper.GetDataByRule("HgrOverhang", (BusinessObject)support, out collectionHgrOverHanger);     //For Straight sections, A is the OH dimension
                    if (collectionHgrOverHanger != null)
                    {
                        if (collectionHgrOverHanger[0] == null)
                            overhangRight = (double)collectionHgrOverHanger[1];
                        else
                            overhangRight = (double)collectionHgrOverHanger[0];
                    }
                }
                else if (overhangOption == 2)
                {
                    if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                    {
                        double distRouteLeftStruct = RefPortHelper.DistanceBetweenPorts("Route", leftStructPort, PortDistanceType.Horizontal);
                        double distRouteRightStruct = RefPortHelper.DistanceBetweenPorts("Route", rightStructPort, PortDistanceType.Horizontal);

                        overhangLeft = distRouteLeftStruct + sectionWidth / 2;
                        overhangRight = distRouteRightStruct + sectionWidth / 2;
                    }
                }

                //========================================
                // Set Values of Part Occurance Attributes
                //========================================

                string[] arrayOfKeys = new string[3];
                arrayOfKeys[0] = HORIZONTALSECTION;
                arrayOfKeys[1] = VERTICALSECTION1;
                arrayOfKeys[2] = VERTICALSECTION2;

                for (int i = 0; i < arrayOfKeys.Length; i++)
                {
                    PropertyValueCodelist topbeginMiterCodelist = (PropertyValueCodelist)componentDictionary[arrayOfKeys[0]].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                    if (topbeginMiterCodelist.PropValue == -1)
                        topbeginMiterCodelist.PropValue = 1;
                    PropertyValueCodelist topendMiterCodelist = (PropertyValueCodelist)componentDictionary[arrayOfKeys[0]].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                    if (topendMiterCodelist.PropValue == -1)
                        topendMiterCodelist.PropValue = 1;
                    componentDictionary[arrayOfKeys[i]].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[arrayOfKeys[i]].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                    componentDictionary[arrayOfKeys[i]].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                    componentDictionary[arrayOfKeys[i]].SetPropertyValue(topbeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                    componentDictionary[arrayOfKeys[i]].SetPropertyValue(topendMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");
                }

                double length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                //Get the largest Pipe Dia
                double[] pipeDiameter = new double[routeCount];

                for (int i = 1; i <= routeCount; i++)
                {
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(i);
                    pipeDiameter[i - 1] = pipeInfo.OutsideDiameter - 2 * pipeInfo.InsulationThickness;
                }
                PipeObjectInfo routeInfo1 = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                Vector vector = routeInfo1.Orientation;
                double distance = Math.Sqrt(vector.X * vector.X + vector.Y * vector.Y), slope = 0.0;
                if (distance < Math3d.DistanceTolerance)
                    slope = Math.PI / 2;
                else
                    slope = Math.Atan(Math.Abs(vector.Z) / vector.Length);

                //For sloped Pipe
                double pipeAngle = GenericAssemblyServices.GetRouteStructConfigAngle(this, "Route", "Structure", PortAxisType.Z);
                if (double.IsNaN(pipeAngle))
                    pipeAngle = 0;
                double angledOffset = Math.Tan(pipeAngle) * sectionWidth;

                double extraLength = 0;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                        extraLength = angledOffset;
                }

                //For UBolt
                Collection<object> uBoltPartCollection;
                GenericHelper.GetDataByRule("HgrUBoltType", (BusinessObject)support, out uBoltPartCollection);
                string uBoltType="";
                double padThickness = 0;
                double[] pipeODWoInsulation = new double[routeCount + 3];
                if (uBoltPartCollection != null)
                {
                    for (int i = 4; i <= routeCount + 3; i++)
                    {
                        if (i == 4)
                            routeIndex = 0;
                        else
                            routeIndex = i - 4;

                        //Check Insulation
                        PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex + 1);
                        pipeODWoInsulation[routeIndex] = pipeInfo.OutsideDiameter - 2 * pipeInfo.InsulationThickness;

                        //Get the UBolt Type
                        //Based on the type, we need to change the vertical offset to include Pad for UBolt Type E

                        if (uBoltPartCollection[routeIndex] == null)
                            uBoltType = (string)uBoltPartCollection[routeIndex + 1];
                        else
                            uBoltType = (string)uBoltPartCollection[routeIndex];

                        if (uBoltType == "E")
                        {
                            padThickness = 0.005; //As per email from China, Pad thickness = 5 mm for UBolt Type E,
                            componentDictionary[uboltPartKeys[routeIndex]].SetPropertyValue(sectionWidth, "IJOAHgrGenUBoltPadL", "PadL");
                        }
                        componentDictionary[uboltPartKeys[routeIndex]].SetPropertyValue(sectionThickness, "IJOAHgrGenUBoltFlgThick", "FlgThick");
                        componentDictionary[uboltPartKeys[routeIndex]].SetPropertyValue(pipeODWoInsulation[routeIndex] / 2, "IJOAHgrGenericUBoltE", "PipeRadius");
                    }
                }

                //Set properties on Assembly
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                int secSize = metadataManager.GetCodelistInfo("GenAssyHgrSecSize", "UDP").GetCodelistItem(sectionSize).Value;
                support.SetPropertyValue(secSize, "IJOAHgrGenAssySection", "SectionSize");  //DI#129319
                support.SetPropertyValue(overhangLeft, "IJOAHgrGenAssyOHLeft", "OverhangLeft");    //DI#129319
                support.SetPropertyValue(overhangRight, "IJOAHgrGenAssyOHRight", "OverhangRight"); //DI#129319

                //=============
                //Create Joints
                //=============
                //Create a collection to hold the joints

                componentDictionary[HORIZONTALSECTION].SetPropertyValue(overhangLeft + extremePipesDistance + overhangRight, "IJUAHgrOccLength", "Length");

                double verticlaLength = 0, leftSectionLength = 0, rightSectionLength = 0, routeStructDistance = 0;

                if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                {
                    if (HgrCompareDoubleService.cmpdbl(Math.Abs(slope) , 0)==true)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Joints" + ": " + "WARNING: " + "This option is valid only for Sloped Route. Resetting the value.", "", "UFrameTypeJ.cs", 427);
                        support.SetPropertyValue(1, "IJOAHgrGenAssyFrmOrient", "FrmOrient");    //DI#129319
                        return;
                    }
                    else
                    {
                        if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Joints" + ": " + "WARNING: " + "This option is not yet implemented. Resetting the value.", "", "UFrameTypeJ.cs", 435);
                            support.SetPropertyValue(1, "IJOAHgrGenAssyFrmOrient", "FrmOrient");    //DI#129319
                            return;
                        }
                    }
                }

                double vertLeftStructDistance, verticalRightStructDistance;

                if ((Math.Abs(routeStructAngle) < (0 - slope + 0.001) && Math.Abs(routeStructAngle) > (0 - slope - 0.001)) || (Math.Abs(routeStructAngle) < ((Math.PI - slope) + 0.001) && Math.Abs(routeStructAngle) > ((Math.PI - slope) - 0.001)))
                {
                    routeStructDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                    vertLeftStructDistance = RefPortHelper.DistanceBetweenPorts("Route", leftStructPort, PortDistanceType.Horizontal);
                    verticalRightStructDistance = RefPortHelper.DistanceBetweenPorts("Route", rightStructPort, PortDistanceType.Horizontal);
                }
                else
                {
                    routeStructDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                    vertLeftStructDistance = RefPortHelper.DistanceBetweenPorts("Route", leftStructPort, PortDistanceType.Vertical);
                    verticalRightStructDistance = RefPortHelper.DistanceBetweenPorts("Route", rightStructPort, PortDistanceType.Vertical);
                }

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (Configuration == 1 || Configuration == 2)
                        verticlaLength = routeStructDistance - pipeODWoInsulation[0] / 2 - padThickness + extraLength;
                    else if (Configuration == 3 || Configuration == 4)
                        verticlaLength = routeStructDistance + pipeODWoInsulation[0] / 2 + sectionDepth + padThickness + extraLength;

                    componentDictionary[VERTICALSECTION1].SetPropertyValue(verticlaLength, "IJUAHgrOccLength", "Length");
                    componentDictionary[VERTICALSECTION2].SetPropertyValue(verticlaLength, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    if (Configuration == 1 || Configuration == 2)
                    {
                        leftSectionLength = vertLeftStructDistance - pipeODWoInsulation[0] / 2 - padThickness + extraLength;
                        rightSectionLength = verticalRightStructDistance - pipeODWoInsulation[0] / 2 - padThickness + extraLength;
                    }
                    else if (Configuration == 3 || Configuration == 4)
                    {
                        leftSectionLength = vertLeftStructDistance + pipeODWoInsulation[0] / 2 + sectionDepth + padThickness + extraLength;
                        rightSectionLength = verticalRightStructDistance + pipeODWoInsulation[0] / 2 + sectionDepth + padThickness + extraLength;
                    }

                    componentDictionary[VERTICALSECTION1].SetPropertyValue(leftSectionLength, "IJUAHgrOccLength", "Length");
                    componentDictionary[VERTICALSECTION2].SetPropertyValue(rightSectionLength, "IJUAHgrOccLength", "Length");
                }
                //Add Braces
                double padThicknessValue = 0;
                if (showBrace == 1)
                {
                    GenericAssemblyServices.ConfigIndex braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(), braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(), bracePadConfigIndex = new GenericAssemblyServices.ConfigIndex();
                    string sectionPort1 = string.Empty; string sectionPort2 = string.Empty; string bracePort1 = string.Empty; string bracePort2 = string.Empty;
                    double braceHtOffFactor = 0, leftBraceLength = 0, leftBracePadOffset = 0, leftBraceHtOffset = 0, rightBraceLength = 0, rightBraceHtOffset = 0, rightBracePadOffset = 0, bracePadPlaneOffset1 = 0, bracePadPlaneOffset2 = 0, bracePadAxisOffset1 = 0, bracePadAxisOffset2 = 0, bracePadOriginOffset1 = 0, bracePadOriginOffset2 = 0, leftPlaneOffset = 0, rightPlaneOffset = 0;
                    GenericHelper.GetDataByRule("HgrBraceHeightOffset", (BusinessObject)support, out braceHtOffFactor);
                    double angleRadian1 = braceAngle / 180 * 3.14159265358979;

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        leftBraceLength = (1 - braceHtOffFactor) * verticlaLength / Math.Cos(angleRadian1) + ((sectionWidth - sectionThickness) * Math.Sin(angleRadian1));
                        leftBracePadOffset = (leftBraceLength) * Math.Sin(angleRadian1);
                        leftBraceHtOffset = braceHtOffFactor * verticlaLength;
                        rightBraceLength = leftBraceLength;
                        rightBraceHtOffset = leftBraceHtOffset;
                        rightBracePadOffset = leftBracePadOffset;
                    }
                    else
                    {
                        leftBraceLength = (1 - braceHtOffFactor) * leftSectionLength / Math.Cos(angleRadian1) + ((sectionWidth - sectionThickness) * Math.Sin(angleRadian1));
                        rightBraceLength = (1 - braceHtOffFactor) * rightSectionLength / Math.Cos(angleRadian1) + ((sectionWidth - sectionThickness) * Math.Sin(angleRadian1));
                        leftBracePadOffset = (leftBraceLength) * Math.Sin(angleRadian1);
                        rightBracePadOffset = (rightBraceLength) * Math.Sin(angleRadian1);
                        leftBraceHtOffset = braceHtOffFactor * leftSectionLength;
                        rightBraceHtOffset = braceHtOffFactor * rightSectionLength;
                    }

                    if (Configuration == 1 || Configuration == 4)
                    {
                        braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.Y, Axis.NegativeX);
                        braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);

                        sectionPort1 = "BeginCap";
                        sectionPort2 = "EndCap";
                        bracePort1 = "BeginCap";
                        bracePort2 = "EndCap";

                        bracePadPlaneOffset1 = 0;
                        bracePadPlaneOffset2 = 0;
                        bracePadAxisOffset1 = -sectionWidth / 3;
                        bracePadAxisOffset2 = -sectionWidth / 3;
                        bracePadOriginOffset1 = leftBracePadOffset;
                        bracePadOriginOffset2 = rightBracePadOffset;

                        bracePadConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                    }
                    else if (Configuration == 2 || Configuration == 3)
                    {
                        braceConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                        braceConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);

                        sectionPort1 = "BeginCap";
                        sectionPort2 = "EndCap";
                        bracePort1 = "EndCap";
                        bracePort2 = "BeginCap";

                        bracePadPlaneOffset1 = 0;
                        bracePadPlaneOffset2 = 0;
                        bracePadAxisOffset1 = -leftBracePadOffset;
                        bracePadAxisOffset2 = -rightBracePadOffset;
                        bracePadOriginOffset1 = sectionWidth / 3;
                        bracePadOriginOffset2 = sectionWidth / 3;

                        bracePadConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                    }

                    if (includeLeftBrPad == 1 && includeRightBrPad == 2)
                    {
                        padThicknessValue = (double)((PropertyValueDouble)componentDictionary[LEFTBRACEPAD].GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;

                        //Add Joint Between the Plate and the Vertical Section
                        JointHelper.CreateRigidJoint(LEFTBRACEPAD, "HgrPort_2", VERTICALSECTION1, "EndCap", bracePadConfigIndex.A, bracePadConfigIndex.B, bracePadConfigIndex.C, bracePadConfigIndex.D, bracePadPlaneOffset1, bracePadAxisOffset1, bracePadOriginOffset1);
                    }
                    else if (includeLeftBrPad == 2 && includeRightBrPad == 1)
                    {
                        padThicknessValue = (double)((PropertyValueDouble)componentDictionary[RIGHTBRACEPAD].GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;

                        //Add Joint Between the Plate and the Vertical Section
                        JointHelper.CreateRigidJoint(RIGHTBRACEPAD, "HgrPort_1", VERTICALSECTION2, "BeginCap", bracePadConfigIndex.A, bracePadConfigIndex.B, bracePadConfigIndex.C, bracePadConfigIndex.D, bracePadPlaneOffset2, bracePadAxisOffset2, bracePadOriginOffset2);
                    }
                    else if (includeLeftBrPad == 1 && includeRightBrPad == 1)
                    {
                        padThicknessValue = (double)((PropertyValueDouble)componentDictionary[LEFTBRACEPAD].GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;

                        //Add Joint Between the Plate and the Vertical Section
                        JointHelper.CreateRigidJoint(LEFTBRACEPAD, "HgrPort_2", VERTICALSECTION1, "EndCap", bracePadConfigIndex.A, bracePadConfigIndex.B, bracePadConfigIndex.C, bracePadConfigIndex.D, bracePadPlaneOffset1, bracePadAxisOffset1, bracePadOriginOffset1);

                        //Add Joint Between the Plate and the Vertical Section
                        JointHelper.CreateRigidJoint(RIGHTBRACEPAD, "HgrPort_1", VERTICALSECTION2, "BeginCap", bracePadConfigIndex.A, bracePadConfigIndex.B, bracePadConfigIndex.C, bracePadConfigIndex.D, bracePadPlaneOffset2, bracePadAxisOffset2, bracePadOriginOffset2);
                    }

                    if (includeLeftBrPad == 1)
                        componentDictionary[LEFTBRACE].SetPropertyValue(leftBraceLength - padThicknessValue, "IJOAHgrGenericBrace", "L");
                    else if (includeLeftBrPad == 2)
                        componentDictionary[LEFTBRACE].SetPropertyValue(leftBraceLength, "IJOAHgrGenericBrace", "L");
                    PropertyValueCodelist leftBraceOrientCodelist = (PropertyValueCodelist)componentDictionary[LEFTBRACE].GetPropertyValue("IJOAHgrGenericBrace", "BraceOrient");
                    if (leftBraceOrientCodelist.PropValue == -1)
                        leftBraceOrientCodelist.PropValue = 2;
                    componentDictionary[LEFTBRACE].SetPropertyValue(angleRadian1, "IJOAHgrGenericBrace", "Angle");
                    componentDictionary[LEFTBRACE].SetPropertyValue(leftBraceOrientCodelist.PropValue = 2, "IJOAHgrGenericBrace", "BraceOrient");
                    componentDictionary[LEFTBRACE].SetPropertyValue(sectionWidth, "IJOAHgrGenericBrace", "W");
                    componentDictionary[LEFTBRACE].SetPropertyValue(sectionDepth, "IJOAHgrGenericBrace", "Depth");
                    componentDictionary[LEFTBRACE].SetPropertyValue(sectionThickness, "IJOAHgrGenericBrace", "FlangeT");

                    if (includeRightBrPad == 1)
                        componentDictionary[RIGHTBRACE].SetPropertyValue(rightBraceLength - padThicknessValue, "IJOAHgrGenericBrace", "L");
                    else if (includeRightBrPad == 2)
                        componentDictionary[RIGHTBRACE].SetPropertyValue(rightBraceLength, "IJOAHgrGenericBrace", "L");
                    PropertyValueCodelist rightBraceOrientCodelist = (PropertyValueCodelist)componentDictionary[RIGHTBRACE].GetPropertyValue("IJOAHgrGenericBrace", "BraceOrient");
                    if (rightBraceOrientCodelist.PropValue == -1)
                        rightBraceOrientCodelist.PropValue = 2;
                    componentDictionary[RIGHTBRACE].SetPropertyValue(angleRadian1, "IJOAHgrGenericBrace", "Angle");
                    componentDictionary[RIGHTBRACE].SetPropertyValue(rightBraceOrientCodelist.PropValue = 2, "IJOAHgrGenericBrace", "BraceOrient");
                    componentDictionary[RIGHTBRACE].SetPropertyValue(sectionWidth, "IJOAHgrGenericBrace", "W");
                    componentDictionary[RIGHTBRACE].SetPropertyValue(sectionDepth, "IJOAHgrGenericBrace", "Depth");
                    componentDictionary[RIGHTBRACE].SetPropertyValue(sectionThickness, "IJOAHgrGenericBrace", "FlangeT");

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
                    JointHelper.CreateRigidJoint(LEFTCONNOBJ, "Connection", VERTICALSECTION1, sectionPort1, braceConfigIndex1.A, braceConfigIndex1.B, braceConfigIndex1.C, braceConfigIndex1.D, leftPlaneOffset, 0, 0);

                    JointHelper.CreateRigidJoint(LEFTCONNOBJ, "Connection", LEFTBRACE, bracePort2, braceConfigIndex2.A, braceConfigIndex2.B, braceConfigIndex2.C, braceConfigIndex2.D, 0, 0, 0);

                    //Add a Joint between horizontal section and the point that we want the angle to rotate around.
                    JointHelper.CreateRigidJoint(RIGHTCONNOBJ, "Connection", VERTICALSECTION2, sectionPort2, braceConfigIndex1.A, braceConfigIndex1.B, braceConfigIndex1.C, braceConfigIndex1.D, rightPlaneOffset, 0, 0);

                    JointHelper.CreateRigidJoint(RIGHTCONNOBJ, "Connection", RIGHTBRACE, bracePort1, braceConfigIndex2.A, braceConfigIndex2.B, braceConfigIndex2.C, braceConfigIndex2.D, 0, 0, 0);
                }

                //'****************************************
                GenericAssemblyServices.ConfigIndex horizontalConfigIndex1 = new GenericAssemblyServices.ConfigIndex(), horizontalConfigIndex2 = new GenericAssemblyServices.ConfigIndex(), structConfigIndex = new GenericAssemblyServices.ConfigIndex(), routeConfigIndex = new GenericAssemblyServices.ConfigIndex(), uboltConfigIndex = new GenericAssemblyServices.ConfigIndex();
                double routePlaneOffset = 0, routeAxisOffset = 0, routeOriginOffset = 0, structAxisOffset = 0, structOriginOffset = 0, horizontalPlaneOffset1 = 0, horizontalPlaneOffset2 = 0, horizontalAxisOffset1 = 0, horizontalAxisOffset2 = 0, horizontalOriginOffset1 = 0, horizontalOriginOffset2 = 0;
                string horizontalPort1 = string.Empty; string horizontalPort2 = string.Empty; string verticalPort1 = string.Empty; string verticalPort2 = string.Empty;

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

                if (connection == 1)      //1 means Mitered joint, no Frame Offsets
                {
                    widthOffset = 0;
                    heightOffset = 0;
                    if (Configuration == 1 || Configuration == 2)
                    {
                        horizontalPlaneOffset1 = 0;
                        horizontalPlaneOffset2 = 0;
                        horizontalAxisOffset1 = 0;
                        horizontalAxisOffset2 = 0;
                        horizontalOriginOffset1 = 0;
                        horizontalOriginOffset2 = 0;
                    }
                    else if (Configuration == 3 || Configuration == 4)
                    {
                        horizontalPlaneOffset1 = 0;
                        horizontalPlaneOffset2 = 0;
                        horizontalAxisOffset1 = sectionDepth;
                        horizontalAxisOffset2 = sectionDepth;
                        horizontalOriginOffset1 = 0;
                        horizontalOriginOffset2 = 0;
                    }

                    if (Configuration == 1)
                    {
                        horizontalConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.XY, Axis.X, Axis.X);
                        horizontalConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.X);
                        uboltConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);

                        structAxisOffset = -sectionWidth / 2;
                        structOriginOffset = -overhangLeft;

                        routePlaneOffset = pipeODWoInsulation[0] / 2 + padThickness;
                        routeAxisOffset = -overhangLeft;
                        routeOriginOffset = -sectionWidth / 2;
                    }
                    else if (Configuration == 2)
                    {
                        horizontalConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.ZX, Axis.Z, Axis.X);
                        horizontalConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX); ;
                        uboltConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);

                        structAxisOffset = sectionWidth / 2;
                        structOriginOffset = -overhangRight - extremePipesDistance;

                        routePlaneOffset = pipeODWoInsulation[0] / 2 + padThickness;
                        routeAxisOffset = overhangRight + extremePipesDistance;
                        routeOriginOffset = sectionWidth / 2;
                    }
                    else if (Configuration == 3)
                    {
                        horizontalConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.YZ, Axis.Y, Axis.Y);
                        horizontalConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeYZ, Axis.Y, Axis.Y);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeY);
                        uboltConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);

                        structAxisOffset = sectionWidth / 2;
                        structOriginOffset = -overhangRight - extremePipesDistance;

                        routePlaneOffset = -pipeODWoInsulation[0] / 2 - padThickness;
                        routeAxisOffset = overhangRight + extremePipesDistance;
                        routeOriginOffset = sectionWidth / 2;
                    }
                    else if (Configuration == 4)
                    {
                        horizontalConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeZX, Axis.Y, Axis.X);
                        horizontalConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.ZX, Axis.Y, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y);
                        uboltConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);

                        structAxisOffset = -sectionWidth / 2;
                        structOriginOffset = -overhangLeft;

                        routePlaneOffset = -pipeODWoInsulation[0] / 2 - padThickness;
                        routeAxisOffset = -overhangLeft;
                        routeOriginOffset = -sectionWidth / 2;
                    }
                }
                else if (connection == 2)     //2 means Wrapping Joint
                {
                    componentDictionary[VERTICALSECTION1].SetPropertyValue(heightOffset, "IJUAHgrOccOverLength", "BeginOverLength");
                    componentDictionary[VERTICALSECTION2].SetPropertyValue(heightOffset, "IJUAHgrOccOverLength", "EndOverLength");
                    if (Configuration == 1)
                    {
                        horizontalConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.XY, Axis.Y, Axis.NegativeX);
                        horizontalConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeXY, Axis.Y, Axis.NegativeX);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY);
                        uboltConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);

                        structAxisOffset = sectionWidth / 2;
                        structOriginOffset = -overhangLeft;

                        horizontalPlaneOffset1 = 0;
                        horizontalPlaneOffset2 = 0;
                        horizontalAxisOffset1 = -widthOffset;
                        horizontalAxisOffset2 = widthOffset;
                        horizontalOriginOffset1 = 0;
                        horizontalOriginOffset2 = 0;

                        routePlaneOffset = pipeODWoInsulation[0] / 2 + padThickness;
                        routeAxisOffset = -overhangLeft;
                        routeOriginOffset = sectionWidth / 2;
                    }
                    else if (Configuration == 2)
                    {
                        horizontalConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.X);
                        horizontalConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeX);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.YZ, Axis.X, Axis.Y);
                        uboltConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);

                        structAxisOffset = -sectionWidth / 2;
                        structOriginOffset = -overhangRight - extremePipesDistance;

                        horizontalPlaneOffset1 = 0;
                        horizontalPlaneOffset2 = 0;
                        horizontalAxisOffset1 = 0;
                        horizontalAxisOffset2 = 0;
                        horizontalOriginOffset1 = widthOffset;
                        horizontalOriginOffset2 = -widthOffset;

                        routePlaneOffset = pipeODWoInsulation[0] / 2 + padThickness;
                        routeAxisOffset = overhangRight + extremePipesDistance;
                        routeOriginOffset = -sectionWidth / 2;
                    }
                    else if (Configuration == 3)
                    {
                        horizontalConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY);
                        horizontalConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeY);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeZX, Axis.X, Axis.X);
                        uboltConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);

                        structAxisOffset = -sectionWidth / 2;
                        structOriginOffset = -overhangRight - extremePipesDistance;

                        horizontalPlaneOffset1 = widthOffset;
                        horizontalPlaneOffset2 = -widthOffset;
                        horizontalAxisOffset1 = sectionDepth;
                        horizontalAxisOffset2 = sectionDepth;
                        horizontalOriginOffset1 = 0;
                        horizontalOriginOffset2 = 0;

                        routePlaneOffset = -pipeODWoInsulation[0] / 2 - padThickness;
                        routeAxisOffset = overhangRight + extremePipesDistance;
                        routeOriginOffset = -sectionWidth / 2;
                    }
                    else if (Configuration == 4)
                    {
                        horizontalConfigIndex1 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeYZ, Axis.Z, Axis.NegativeY);
                        horizontalConfigIndex2 = new GenericAssemblyServices.ConfigIndex(Plane.YZ, Plane.NegativeYZ, Axis.Z, Axis.Y);
                        routeConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX);
                        uboltConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                        structConfigIndex = new GenericAssemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);

                        horizontalPlaneOffset1 = 0;
                        horizontalPlaneOffset2 = 0;
                        horizontalAxisOffset1 = sectionDepth;
                        horizontalAxisOffset2 = sectionDepth;
                        horizontalOriginOffset1 = -widthOffset;
                        horizontalOriginOffset2 = widthOffset;

                        structAxisOffset = sectionWidth / 2;
                        structOriginOffset = -overhangLeft;

                        routePlaneOffset = -pipeODWoInsulation[0] / 2 - padThickness;
                        routeAxisOffset = -overhangLeft;
                        routeOriginOffset = sectionWidth / 2;
                    }
                }

                if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                {
                    //Add Joint Between the Horizontal and Vertical Beams
                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, horizontalPort1, VERTICALSECTION1, verticalPort1, horizontalConfigIndex1.A, horizontalConfigIndex1.B, horizontalConfigIndex1.C, horizontalConfigIndex1.D, horizontalPlaneOffset1, horizontalAxisOffset1, horizontalOriginOffset1);

                    //Add Joint Between the Horizontal and Vertical Beams
                    JointHelper.CreateRigidJoint(HORIZONTALSECTION, horizontalPort2, VERTICALSECTION2, verticalPort2, horizontalConfigIndex2.A, horizontalConfigIndex2.B, horizontalConfigIndex2.C, horizontalConfigIndex2.D, horizontalPlaneOffset2, horizontalAxisOffset2, horizontalOriginOffset2);
                }
                else if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                {
                    if (Configuration == 1 || Configuration == 2)
                    {
                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, horizontalPort1, VERTICALSECTION1, verticalPort1, horizontalConfigIndex1.A, horizontalConfigIndex1.B, horizontalConfigIndex1.C, horizontalConfigIndex1.D, horizontalPlaneOffset1, horizontalAxisOffset1, horizontalOriginOffset1);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, horizontalPort2, VERTICALSECTION2, verticalPort2, horizontalConfigIndex2.A, horizontalConfigIndex2.B, horizontalConfigIndex2.C, horizontalConfigIndex2.D, horizontalPlaneOffset2, horizontalAxisOffset2, horizontalOriginOffset2);
                    }
                    else if (Configuration == 3 || Configuration == 4)
                    {
                        //Add joints between Route and Beam
                        JointHelper.CreateRigidJoint(CONNOBJ1, "Connection", VERTICALSECTION1, verticalPort1, horizontalConfigIndex1.A, horizontalConfigIndex1.B, horizontalConfigIndex1.C, horizontalConfigIndex1.D, horizontalPlaneOffset1, horizontalAxisOffset1, horizontalOriginOffset1);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(HORIZONTALSECTION, horizontalPort1, CONNOBJ1, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        //Add joints between Route and Beam
                        JointHelper.CreateRigidJoint(CONNOBJ2, "Connection", VERTICALSECTION2, verticalPort2, horizontalConfigIndex2.A, horizontalConfigIndex2.B, horizontalConfigIndex2.C, horizontalConfigIndex2.D, horizontalPlaneOffset2, horizontalAxisOffset2, horizontalOriginOffset2);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreatePrismaticJoint(HORIZONTALSECTION, horizontalPort2, CONNOBJ2, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    }
                }

                string refPortName = string.Empty;
                for (int i = 4; i <= routeCount + 3; i++)
                {
                    if (i == 4)
                        refPortName = "Route";
                    else
                        refPortName = "Route_" + (i - 4 + 1);

                    if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                        componentDictionary[uboltPartKeys[i - 4]].SetPropertyValue(0.0, "IJOAHgrGenericUBoltE", "Angle");
                    else if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                    {
                        if (Configuration == 1 || Configuration == 2)
                            componentDictionary[uboltPartKeys[i - 4]].SetPropertyValue(slope, "IJOAHgrGenericUBoltE", "Angle");
                        else if (Configuration == 3 || Configuration == 4)
                            componentDictionary[uboltPartKeys[i - 4]].SetPropertyValue(-slope, "IJOAHgrGenericUBoltE", "Angle");
                    }

                    //Add Joint Between the UBolt and Route
                    JointHelper.CreateRigidJoint(uboltPartKeys[i - 4], "Route", "-1", refPortName, uboltConfigIndex.A, uboltConfigIndex.B, uboltConfigIndex.C, uboltConfigIndex.D, 0, 0, 0);
                }

                if (includeLeftPad == 2 && includeRightPad == 1)
                {
                    padThicknessValue = (double)((PropertyValueDouble)componentDictionary[RIGHTPAD].GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                    componentDictionary[VERTICALSECTION1].SetPropertyValue(-padThicknessValue, "IJUAHgrOccOverLength", "EndOverLength");

                    //Add Joint Between the Left and the Vertical Section 1
                    JointHelper.CreateRigidJoint(RIGHTPAD, "HgrPort_2", VERTICALSECTION1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                }
                else if (includeLeftPad == 1 && includeRightPad == 2)
                {
                    padThicknessValue = (double)((PropertyValueDouble)componentDictionary[LEFTPAD].GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                    componentDictionary[VERTICALSECTION2].SetPropertyValue(-padThicknessValue, "IJUAHgrOccOverLength", "BeginOverLength");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(LEFTPAD, "HgrPort_1", VERTICALSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                }
                else if (includeRightPad == 1 && includeLeftPad == 1)
                {
                    padThicknessValue = (double)((PropertyValueDouble)componentDictionary[LEFTPAD].GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                    componentDictionary[VERTICALSECTION1].SetPropertyValue(-padThicknessValue, "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[VERTICALSECTION2].SetPropertyValue(-padThicknessValue, "IJUAHgrOccOverLength", "BeginOverLength");

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(LEFTPAD, "HgrPort_2", VERTICALSECTION1, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);

                    //Add Joint Between the Plate and the Vertical Beam
                    JointHelper.CreateRigidJoint(RIGHTPAD, "HgrPort_1", VERTICALSECTION2, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionDepth / 3);
                }

                // What you need to do in here is:
                //   - If we are connecting PERPENDICULAR TO PIPE then we only connect to the pipe
                //   - If we connecting PERPENDICULAR TO THE STRUCTURE then we only connec to the structure
                if (fromOrient.ToUpper() == "PERPENDICULAR TO PIPE")
                    //Add joints between Route and Beam
                    JointHelper.CreateRigidJoint("-1", "Route", HORIZONTALSECTION, "EndCap", routeConfigIndex.A, routeConfigIndex.B, routeConfigIndex.C, routeConfigIndex.D, routePlaneOffset, routeAxisOffset, routeOriginOffset);
                else if (fromOrient.ToUpper() == "PERPENDICULAR TO STRUCTURE")
                {
                    //Add Joint Between Structure and Vertical Section 1
                    JointHelper.CreateRigidJoint("-1", "Structure", VERTICALSECTION1, "EndCap", structConfigIndex.A, structConfigIndex.B, structConfigIndex.C, structConfigIndex.D, 0, structAxisOffset, structOriginOffset);

                    //Add Joint Between Structure and Vertical Section 2
                    JointHelper.CreatePlanarJoint("-1", "Structure", VERTICALSECTION2, "BeginCap", Plane.XY, Plane.XY, 0);
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
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(HORIZONTALSECTION, 1));

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
                    structConnections.Add(new ConnectionInfo(VERTICALSECTION1, 1));

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
        // BOM Description
        //-------------------------------------------------------------
        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string bomDescription = "";
            try
            {
                string bom = (string)((PropertyValueString)supportOrComponent.GetPropertyValue("IJOAHgrGenAssyBOMDesc", "BOM_DESC")).PropValue;
                if (string.IsNullOrEmpty(bom))
                    bomDescription = "Assembly Type J for Single or Multiple Pipes";
                else
                    bomDescription = bom;
                return bomDescription;
            }
            catch (Exception e)  //General Unhandled exception 
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error BOM Description." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        #endregion
    }
}


