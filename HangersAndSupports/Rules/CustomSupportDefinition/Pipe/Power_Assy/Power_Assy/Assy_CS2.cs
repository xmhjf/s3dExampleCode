//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_CS2.cs
//   Power_Assy,Ingr.SP3D.Content.Support.Rules.Assy_CS2
//   Author       :  Vijay
//   Creation Date:  20.Mar.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-03-2013     Vijay   CR-CP-224472-Initial Creation
//   22/02/2015    PVK      TR-CP-264951  Resolve coverity issues found in November 2014 report 
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle;
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
    [SymbolVersion("1.0.0.0")]
    public class Assy_CS2 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string PipeClamp = "PipeClamp";
        private const string Rod1 = "Rod1";
        private const string Rod2 = "Rod2";
        private const string BeamAtt1 = "BeamAtt1";
        private const string BeamAtt2 = "BeamAtt2";
        private const string EyeNut1 = "EyeNut1";
        private const string EyeNut2 = "EyeNut2";
        private const string EyeNut3 = "EyeNut3";
        private const string EyeNut4 = "EyeNut4";
        private const string ConnObj1 = "ConnObj1";
        private const string ConnObj2 = "ConnObj2";

        private const string spring1 = "Spring1";
        private const string spring2 = "Spring2";
        private const string rod3 = "Rod3";
        private const string rod4 = "Rod4";
        private string clampPart;
        private double clampWidth;
        private double clampHeight;
        private double rodSpacing;
        private long showSpring;
        private long spring1Pos;
        private long spring2Pos;
        private double springHeight;
        private double springDia;
        private string rodSize;

        private long rotAngle;
        private double pipeDia;
        private string connection;

        private const string Steel = "Steel";
        private const string Slab = "Slab";
        private const string SlabSteel = "SlabSteel";
        private const string SteelSlab = "SteelSlab";
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                    clampWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyClamp", "ClampWidth")).PropValue;
                    clampHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyClamp", "ClampHeight")).PropValue;
                    rodSpacing = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyRodSP", "RodSpacing")).PropValue;
                    showSpring = (long)((PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowAssySpring", "ShowSpring")).PropValue;
                    spring1Pos = (long)((PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowAssySpring", "Spring1Pos")).PropValue;
                    spring2Pos = (long)((PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowAssySpring", "Spring2Pos")).PropValue;
                    springHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssySpring", "Height")).PropValue;
                    springDia = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssySpring", "Dia")).PropValue;
                    rotAngle = (long)((PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowAssyRotAngle", "RotAngle")).PropValue;
                    PropertyValueCodelist rodSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowAssyCSRod", "RodSize");
                    int rodSizeValue = (int)rodSizeCodelist.PropValue;
                    rodSize = rodSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)rodSizeValue).ShortDisplayName;
                    double pipeDia1 = 0;
                    string strUnit = string.Empty;

                    pipeDia = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyClampDW", "PipeOutDia")).PropValue;
                    pipeDia = pipeDia / 1000;
                    strUnit = "mm";

                    if (strUnit.ToString() == "mm")
                        pipeDia1 = pipeDia * 1000;

                    PartClass auxTable = (PartClass)cataloghelper.GetPartClass("Assy_CS2PartSel");
                    ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    foreach (BusinessObject classItem in classItems)
                    {
                        if (double.Parse(classItem.GetPropertyValue("IJUAHgrAssyCS2PartSel", "PipeOutDia").ToString()) > pipeDia1 - 0.01 && double.Parse(classItem.GetPropertyValue("IJUAHgrAssyCS2PartSel", "PipeOutDia").ToString()) < pipeDia1 + 0.01 && classItem.GetPropertyValue("IJUAHgrAssyCS2PartSel", "UnitType").ToString().ToLower() == strUnit.ToString().ToLower())
                        {
                            clampPart = (string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAssyCS2PartSel", "PartNo")).ToString();
                        }
                    }
                    parts.Add(new PartInfo(PipeClamp, clampPart));
                    parts.Add(new PartInfo(Rod1, "Anvil_FIG140_3"));
                    parts.Add(new PartInfo(Rod2, "Anvil_FIG140_3"));
                    parts.Add(new PartInfo(BeamAtt1, "Anvil_FIG66_3"));
                    parts.Add(new PartInfo(BeamAtt2, "Anvil_FIG66_3"));
                    parts.Add(new PartInfo(EyeNut1, "Anvil_FIG290_3"));
                    parts.Add(new PartInfo(EyeNut2, "Anvil_FIG290_3"));
                    parts.Add(new PartInfo(EyeNut3, "Anvil_FIG290_3"));
                    parts.Add(new PartInfo(EyeNut4, "Anvil_FIG290_3"));
                    parts.Add(new PartInfo(ConnObj1, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(ConnObj2, "Log_Conn_Part_1"));

                    if (showSpring == 1)
                    {
                        parts.Add(new PartInfo(spring1, "Spring-1"));
                        parts.Add(new PartInfo(spring2, "Spring-1"));
                        parts.Add(new PartInfo(rod3, "Anvil_FIG140_" + rodSize));
                        parts.Add(new PartInfo(rod4, "Anvil_FIG140_" + rodSize));

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
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                Boolean[] bIsOffsetApplied = GetIsLugEndOffsetApplied();
                string[] idxStructPort = GetIndexedStructPortName(bIsOffsetApplied);
                string rightStructPort = idxStructPort[0];
                string leftStructPort = idxStructPort[1];

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                SupportComponent pipe_clamp = componentDictionary[PipeClamp];

                //Determine whether connecting to Steel or a Slab
                if (SupportHelper.SupportingObjects.Count != 0)
                {
                    if (SupportHelper.SupportingObjects.Count > 1)
                    {
                        connection = Steel;
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                        {
                            connection = Slab;          //Two Slabs
                        }
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Member))
                        {
                            connection = SlabSteel;    //Slab then Steel
                        }
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                        {
                            connection = SteelSlab;    //Steel then Slab
                        }
                    }
                    else
                    {
                        if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                        {
                            connection = Steel;
                        }
                        else if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab)
                        {
                            connection = Slab;
                        }
                    }
                }
                else
                    connection = Slab;

                if (PW_IsStructureSlopedAcrossPipe(connection))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "The structure is sloped. Please check.", "", "Assy_CS2.cs", 209);
                    return;
                }


                if (clampWidth <= rodSpacing)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Rod Spacing should be less than the Clamp Width. Please check the value", "", "Assy_CS2.cs", 216);
                    return;
                }

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();

                PartClass assy_CS1_VPartClass = (PartClass)cataloghelper.GetPartClass("Assy_CS1_V");
                ReadOnlyCollection<BusinessObject> parts = assy_CS1_VPartClass.Parts;

                // if user doesn't edit the value, set the default rod spacing to three times the Pipe Outside diameter
                double tempRodSpacing = (double)((PropertyValueDouble)parts[0].GetPropertyValue("IJUAHgrPowAssyRodSP_V", "RodSpacing")).PropValue;

                if (rodSpacing >= (tempRodSpacing - Math3d.DistanceTolerance) && (rodSpacing <= (tempRodSpacing + Math3d.DistanceTolerance)))
                {
                    rodSpacing = 3 * pipeDia;
                    clampWidth = rodSpacing + 0.1;
                }
                pipe_clamp.SetPropertyValue(rodSpacing, "IJUAHgrPowerCS", "C");
                pipe_clamp.SetPropertyValue(clampHeight, "IJUAHgrPowerCS", "ClampHeight");
                pipe_clamp.SetPropertyValue(clampWidth, "IJUAHgrPowerCS", "L");

                //Spring 1 position
                if (showSpring == 1)
                {
                    double eyenutG, eyenutE, springG, beamAttE, beamAttHt, beamAttTakeout, rodDia;

                    SupportComponent beamAtt = componentDictionary[BeamAtt1];
                    BusinessObject beamAtt_1 = beamAtt.GetRelationship("madeFrom", "part").TargetObjects[0];

                    SupportComponent spring_1 = componentDictionary[spring1];
                    SupportComponent spring_2 = componentDictionary[spring2];
                    SupportComponent rod_3 = componentDictionary[rod3];
                    SupportComponent rod_4 = componentDictionary[rod4];
                    beamAttHt = (double)((PropertyValueDouble)beamAtt_1.GetPropertyValue("IJUAHgrAnvil_fig66", "B")).PropValue;
                    beamAttTakeout = (double)((PropertyValueDouble)beamAtt_1.GetPropertyValue("IJUAHgrTake_Out", "TAKE_OUT")).PropValue;
                    beamAttE = (double)((PropertyValueDouble)beamAtt_1.GetPropertyValue("IJUAHgrAnvil_fig66", "E")).PropValue;

                    SupportComponent rod = componentDictionary[Rod1];
                    BusinessObject rod_1 = rod.GetRelationship("madeFrom", "part").TargetObjects[0];

                    rodDia = (double)((PropertyValueDouble)rod_1.GetPropertyValue("IJUAHgrRod_Dia", "ROD_DIA")).PropValue;

                    SupportComponent eye1 = componentDictionary[EyeNut1];
                    BusinessObject eyenut = eye1.GetRelationship("madeFrom", "part").TargetObjects[0];

                    eyenutG = (double)((PropertyValueDouble)eyenut.GetPropertyValue("IJUAHgrAnvil_FIG290", "G")).PropValue;
                    eyenutE = (double)((PropertyValueDouble)eyenut.GetPropertyValue("IJUAHgrAnvil_FIG290", "E")).PropValue;


                    SupportComponent nSpring = componentDictionary[spring1];
                    BusinessObject spri = nSpring.GetRelationship("madeFrom", "part").TargetObjects[0];

                    springG = (double)((PropertyValueDouble)spri.GetPropertyValue("IJUAHgrPowerCSSpring", "G")).PropValue;

                    spring_1.SetPropertyValue(springHeight, "IJUAHgrPowerCSSpring", "B");
                    spring_1.SetPropertyValue(springDia, "IJUAHgrPowerCSSpring", "C");
                    spring_1.SetPropertyValue(springDia + springDia / 5, "IJUAHgrPowerCSSpring", "D");
                    spring_1.SetPropertyValue(springHeight, "IJUAHgrPowTakeOut", "TakeOut");
                    spring_1.SetPropertyValue(rodDia, "IJUAHgrPowerCSSpring", "A");

                    spring_2.SetPropertyValue(springHeight, "IJUAHgrPowerCSSpring", "B");
                    spring_2.SetPropertyValue(springDia, "IJUAHgrPowerCSSpring", "C");
                    spring_2.SetPropertyValue(springDia + springDia / 5, "IJUAHgrPowerCSSpring", "D");
                    spring_2.SetPropertyValue(springHeight, "IJUAHgrPowTakeOut", "TakeOut");
                    spring_2.SetPropertyValue(rodDia, "IJUAHgrPowerCSSpring", "A");

                    double routeStrVertDist, spring1Dist = 0, spring2Dist = 0;
                    routeStrVertDist = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);



                    //Spring 1 position
                    if (spring1Pos == 1)
                    {
                        spring1Dist = eyenutG + springG + rodDia / 2;
                    }
                    else if (spring1Pos == 2)
                    {
                        spring1Dist = routeStrVertDist / 2 - springHeight / 2 - beamAttHt;
                    }
                    else if (spring1Pos == 3)
                    {
                        spring1Dist = routeStrVertDist - (beamAttTakeout + eyenutE - rodDia / 2) + clampHeight / 2 + springG;
                    }

                    //Spring 2 position
                    if (spring2Pos == 1)
                    {
                        spring2Dist = eyenutG + springG + rodDia / 2;
                    }
                    else if (spring2Pos == 2)
                    {
                        spring2Dist = routeStrVertDist / 2 - springHeight / 2 - beamAttHt;
                    }
                    else if (spring2Pos == 3)
                    {
                        spring2Dist = routeStrVertDist - (beamAttTakeout + eyenutE - rodDia / 2) + clampHeight / 2 + springG;
                    }

                    rod_3.SetPropertyValue(spring1Dist, "IJUAHgrOccLength", "Length");
                    rod_4.SetPropertyValue(spring2Dist, "IJUAHgrOccLength", "Length");
                }


                double routeStructConfigAng, structDirAng;
                Plane routePlane1, routePlane2, leftStructPortPlane1, leftStructPortPlane2, rightStructPortPlane1, rightStructPortPlane2;
                routePlane1 = routePlane2 = leftStructPortPlane1 = leftStructPortPlane2 = rightStructPortPlane1 = rightStructPortPlane2 = Plane.XY;

                if (connection == Steel)
                {
                    if (leftStructPort == rightStructPort)
                    {
                        leftStructPort = "Structure";
                        rightStructPort = "Structure";
                        routePlane1 = Plane.ZX;
                        routePlane2 = Plane.YZ;
                        leftStructPortPlane1 = rightStructPortPlane1 = Plane.XY;
                        leftStructPortPlane2 = rightStructPortPlane2 = Plane.NegativeXY;

                    }
                    else
                    {
                        routeStructConfigAng = RefPortHelper.PortConfigurationAngle("Struct_2", "Route", PortAxisType.Y);
                        structDirAng = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, OrientationAlong.Global_X);


                        leftStructPortPlane1 = rightStructPortPlane1 = Plane.XY;
                        leftStructPortPlane2 = rightStructPortPlane2 = Plane.XY;

                        if ((structDirAng > (0 - 0.0001) && structDirAng < (0 + 0.0001) || structDirAng > (Math.PI - 0.0001) && structDirAng < (Math.PI + 0.0001)))
                        {
                            if (routeStructConfigAng < Math.PI / 2)
                            {
                                routePlane1 = Plane.YZ;
                                routePlane2 = Plane.NegativeYZ;
                            }
                            else
                            {
                                routePlane1 = Plane.YZ;
                                routePlane2 = Plane.YZ;
                            }
                        }
                        else
                        {
                            if (routeStructConfigAng > Math.PI / 2)
                            {
                                routePlane1 = Plane.YZ;
                                routePlane2 = Plane.NegativeYZ;
                            }
                            else
                            {
                                routePlane1 = Plane.YZ;
                                routePlane2 = Plane.YZ;
                            }
                        }
                    }
                }
                else if (connection == Slab)
                {
                    leftStructPortPlane1 = rightStructPortPlane1 = Plane.XY;
                    leftStructPortPlane2 = rightStructPortPlane2 = Plane.NegativeXY;
                    leftStructPort = "Structure";
                    rightStructPort = "Structure";
                }
                else if (connection == SlabSteel)
                {
                    leftStructPortPlane1 = Plane.XY;
                    leftStructPortPlane2 = Plane.NegativeXY;
                    rightStructPortPlane1 = Plane.XY;
                    rightStructPortPlane2 = Plane.XY;
                }
                else if (connection == SteelSlab)
                {
                    leftStructPortPlane1 = Plane.XY;
                    leftStructPortPlane2 = Plane.XY;
                    rightStructPortPlane1 = Plane.XY;
                    rightStructPortPlane2 = Plane.NegativeXY;
                }

                if (connection != Steel)
                {
                    if (rotAngle == 1)
                    {
                        routePlane1 = Plane.YZ;
                        routePlane2 = Plane.NegativeYZ;
                    }
                    else if (rotAngle == 2)
                    {
                        routePlane1 = Plane.ZX;
                        routePlane2 = Plane.YZ;
                    }
                }

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support should be placed by Point.", "", "Assy_CS2.cs", 409);
                    return;
                }
                else
                {
                    JointHelper.CreateSphericalJoint(PipeClamp, "Route", "-1", "Route");
                    JointHelper.CreatePlanarJoint(PipeClamp, "Route", "-1", "Structure", routePlane1, routePlane2, 0);
                    //Vertical Joint
                    JointHelper.CreateGlobalAxesAlignedJoint(PipeClamp, "Route", Axis.Z, Axis.Z);
                }

                //Add a Joint between the Pipe Clamps and and Eye Nuts
                JointHelper.CreateRigidJoint(PipeClamp, "LeftPin", EyeNut1, "Eye", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(PipeClamp, "RightPin", EyeNut2, "Eye", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                //Create the Flexible (Prismatic) Joint between the ports of the top rods
                JointHelper.CreatePrismaticJoint(Rod1, "TopExThdRH", Rod1, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                JointHelper.CreatePrismaticJoint(Rod2, "TopExThdRH", Rod2, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                //Add a Joint between the beam attachment and eye nut
                JointHelper.CreateRigidJoint(EyeNut3, "Eye", BeamAtt1, "Pin", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                JointHelper.CreateRigidJoint(EyeNut4, "Eye", BeamAtt2, "Pin", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                //Add a Joint between the Rods and Eye Nuts
                JointHelper.CreateRigidJoint(Rod1, "TopExThdRH", EyeNut1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(Rod2, "TopExThdRH", EyeNut2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                if (showSpring == 1)
                {
                    //Spring 1 joints
                    JointHelper.CreateRigidJoint(Rod1, "BotExThdRH", spring1, "BotInThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(spring1, "TopInThdRH", rod3, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(rod3, "BotExThdRH", EyeNut3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    //Spring 2 joints
                    JointHelper.CreateRigidJoint(Rod2, "BotExThdRH", spring2, "BotInThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(spring2, "TopInThdRH", rod4, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(rod4, "BotExThdRH", EyeNut4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                else
                {
                    JointHelper.CreateRigidJoint(Rod1, "BotExThdRH", EyeNut3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(Rod2, "BotExThdRH", EyeNut4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                if (connection == Steel)
                {
                    if (leftStructPort != rightStructPort)
                    {
                        JointHelper.CreateSphericalJoint(ConnObj1, "Connection", "-1", leftStructPort);
                        JointHelper.CreateSphericalJoint(ConnObj2, "Connection", "-1", rightStructPort);
                        JointHelper.CreateGlobalAxesAlignedJoint(ConnObj1, "Connection", Axis.X, Axis.X);
                        JointHelper.CreateGlobalAxesAlignedJoint(ConnObj2, "Connection", Axis.X, Axis.X);
                        JointHelper.CreatePlanarJoint(BeamAtt1, "Structure", ConnObj1, "Connection", leftStructPortPlane1, leftStructPortPlane2, 0);
                        JointHelper.CreatePlanarJoint(BeamAtt2, "Structure", ConnObj2, "Connection", rightStructPortPlane1, rightStructPortPlane2, 0);
                    }
                    else
                    {
                        JointHelper.CreatePlanarJoint(BeamAtt1, "Structure", "-1", leftStructPort, leftStructPortPlane1, leftStructPortPlane2, 0);
                        JointHelper.CreatePlanarJoint(BeamAtt2, "Structure", "-1", rightStructPort, rightStructPortPlane1, rightStructPortPlane2, 0);
                    }
                }
                else if (connection == SlabSteel)
                {
                    JointHelper.CreatePlanarJoint(BeamAtt2, "Structure", "-1", leftStructPort, leftStructPortPlane1, leftStructPortPlane2, 0);
                    JointHelper.CreateSphericalJoint(ConnObj1, "Connection", "-1", rightStructPort);
                    JointHelper.CreateGlobalAxesAlignedJoint(ConnObj1, "Connection", Axis.X, Axis.X);
                    JointHelper.CreateGlobalAxesAlignedJoint(ConnObj1, "Connection", Axis.Z, Axis.Z);
                    JointHelper.CreatePlanarJoint(BeamAtt1, "Structure", ConnObj1, "Connection", rightStructPortPlane1, rightStructPortPlane2, 0);
                }
                else if (connection == SteelSlab)
                {
                    JointHelper.CreatePlanarJoint(BeamAtt2, "Structure", "-1", rightStructPort, rightStructPortPlane1, rightStructPortPlane2, 0);
                    JointHelper.CreateSphericalJoint(ConnObj1, "Connection", "-1", leftStructPort);
                    JointHelper.CreateGlobalAxesAlignedJoint(ConnObj1, "Connection", Axis.X, Axis.X);
                    JointHelper.CreateGlobalAxesAlignedJoint(ConnObj1, "Connection", Axis.Z, Axis.Z);
                    JointHelper.CreatePlanarJoint(BeamAtt1, "Structure", ConnObj1, "Connection", leftStructPortPlane1, leftStructPortPlane2, 0);
                }
                else if (connection == Slab)
                {
                    JointHelper.CreatePlanarJoint(BeamAtt1, "Structure", "-1", leftStructPort, leftStructPortPlane1, leftStructPortPlane2, 0);
                    JointHelper.CreatePlanarJoint(BeamAtt2, "Structure", "-1", rightStructPort, rightStructPortPlane1, rightStructPortPlane2, 0);
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of Power_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        public Boolean PW_IsStructureSlopedAcrossPipe(string connection, Boolean bOnTurn = false)
        {
            try
            {
                bool PW_IsStructureSlopedAcrossPipe = true;

                double refAngle1;
                double refAngle2;
                double angle1 = 89.99999999999 / 180 * 3.14159265358979;
                double angle2 = 90.0000000001 / 180 * 3.14159265358979;

                refAngle1 = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                refAngle2 = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, OrientationAlong.Global_Z);


                if (connection == Slab)
                {
                    if (refAngle1 > angle1 && refAngle2 < angle2)
                    {
                        PW_IsStructureSlopedAcrossPipe = false;
                    }
                    else
                    {
                        PW_IsStructureSlopedAcrossPipe = true;
                    }
                }
                else
                {
                    if (refAngle2 > angle1 && refAngle2 < angle2)
                    {
                        PW_IsStructureSlopedAcrossPipe = false;
                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct || bOnTurn)
                        {
                            PW_IsStructureSlopedAcrossPipe = true;
                        }
                        else
                        {
                            PW_IsStructureSlopedAcrossPipe = false;
                        }
                    }
                }


                return PW_IsStructureSlopedAcrossPipe;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in PW_IsStructureSlopedAcrossPipe Method of Power_Assy.CS2" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string BOMString = "";
            try
            {
                String BOM_DESC = string.Empty;

                BOM_DESC = oSupportOrComponent.GetPropertyValue("IJOAHgrPowAssyBOMDesc", "BOM_DESC").ToString();
                if (BOM_DESC == "")
                {
                    BOMString = "CS2 Assembly with fixed clamp size (overall width and depth) and rod spacing";
                }
                else
                {
                    BOMString = BOM_DESC;
                }

                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - Assy_CS2" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        #endregion
        #endregion

        /// <summary>
        /// This method checks whether a offset value needs to be explicitly specified on both end of the leg.
        /// </summary>
        /// <returns>Boolean array</returns>
        /// <code>
        ///     Boolean[] bIsOffsetApplied = GetIsLugEndOffsetApplied();  
        ///</code>
        public Boolean[] GetIsLugEndOffsetApplied()
        {
            try
            {
                Collection<BusinessObject> StructureObjects;
                Boolean[] isOffsetApplied = new Boolean[2];

                //first route object is set as primary route object
                StructureObjects = SupportHelper.SupportingObjects;

                isOffsetApplied[0] = true;
                isOffsetApplied[1] = true;
                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                if (StructureObjects != null)
                {
                    if (StructureObjects.Count >= 2)
                    {
                        for (int index = 0; index <= 1; index++)
                        {
                            //check if offset is to be applied
                            const double RouteStructAngle = Math.PI / 180.0;
                            Boolean varRuleApplied = true;

                            if (supportingType == "Steel")
                            {
                                //if angle is within 1 degree, regard as parallel case
                                //Also check for Sloped structure                                
                                MemberPart memberPart = (MemberPart)SupportHelper.SupportingObjects[index];
                                ICurve memberCurve = memberPart.Axis;

                                Vector supportedVector = new Vector();
                                Vector supportingVector = new Vector();

                                if (SupportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.Pipe)
                                {
                                    Position startLocation = new Position(SupportedHelper.SupportedObjectInfo(1).StartLocation);
                                    Position endLocation = new Position(SupportedHelper.SupportedObjectInfo(1).EndLocation);
                                    supportedVector = new Vector(endLocation - startLocation);
                                }
                                if (memberCurve is ILine)
                                {
                                    ILine line = (ILine)memberCurve;
                                    supportingVector = line.Direction;
                                }

                                double angle = GetAngleBetweenVectors(supportingVector, supportedVector);
                                double refAngle1 = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.Y, OrientationAlong.Global_X) - Math.PI / 2;
                                double refAngle2 = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.X, OrientationAlong.Global_X);

                                if (angle < (refAngle1 + 0.001) && angle > (refAngle1 - 0.001))
                                    angle = angle - Math.Abs(refAngle1);
                                else if (angle < (refAngle2 + 0.001) && angle > (refAngle2 - 0.001))
                                    angle = angle - Math.Abs(refAngle2);
                                else
                                    angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                if (Math.Abs(angle) < RouteStructAngle || Math.Abs(angle - Math.PI) < RouteStructAngle)
                                    varRuleApplied = false;
                            }

                            isOffsetApplied[index] = varRuleApplied;
                        }
                    }
                }

                return isOffsetApplied;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in GetIsLugEndOffsetApplied Method of Power_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method returns the direct angle between two vectors in Radians
        /// </summary>
        /// <param name="vector1">The vector1(Vector).</param>
        /// <param name="vector2">The vector2(Vector).</param>
        /// <returns>double</returns>        
        /// <code>
        ///       ContentHelper contentHelper = new ContentHelper();
        ///       double value;
        ///       value = contentHelper. GetAngleBetweenVectors(vector1, vector2 );
        ///</code>
        public double GetAngleBetweenVectors(Vector vector1, Vector vector2)
        {
            try
            {
                double dotProd = (vector1.Dot(vector2) / (vector1.Length * vector2.Length));
                double arcCos = 0.0;
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(dotProd), 1) == false)
                {
                    arcCos = Math.PI / 2 - Math.Atan(dotProd / Math.Sqrt(1 - dotProd * dotProd));
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, -1) == true)
                {
                    arcCos = Math.PI;
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, 1) == true)
                {
                    arcCos = 0;
                }
                return arcCos;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in GetIsLugEndOffsetApplied  Method of Power_Assy.." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }

        }

        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();

                    routeConnections.Add(new ConnectionInfo(PipeClamp, 1));

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

                    structConnections.Add(new ConnectionInfo(PipeClamp, 1));

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

        /// <summary>
        /// This method differentiate multiple structure input object based on relative position.
        /// </summary>
        /// <param name="IsOffsetApplied">IsOffsetApplied-The offset applied.(Boolean[])</param>
        /// <returns>String array</returns>
        /// <code>
        ///     string[] idxStructPort = GetIndexedStructPortName(bIsOffsetApplied);
        /// </code>
        public String[] GetIndexedStructPortName(Boolean[] IsOffsetApplied)
        {
            try
            {
                String[] structurePort = new String[2];
                int structureCount = SupportHelper.SupportingObjects.Count;
                int i;
                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    structurePort[0] = "Structure";
                    structurePort[1] = "Structure";
                }
                else
                {
                    structurePort[0] = "Structure";
                    structurePort[1] = "Struct_2";

                    if (structureCount > 1)
                    {
                        if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                        {
                            for (i = 0; i <= 1; i++)
                            {
                                double angle = 0;
                                if ((supportingType == "Steel") && IsOffsetApplied[i] == false)
                                {
                                    angle = RefPortHelper.PortConfigurationAngle("Route", structurePort[i], PortAxisType.Y);
                                }
                                //the port is the right structure port
                                if (Math.Abs(angle) < Math.PI / 2.0)
                                {
                                    if (i == 0)
                                    {
                                        structurePort[0] = "Struct_2";
                                        structurePort[1] = "Structure";
                                    }
                                }
                                //the port is the left structure port
                                else
                                {
                                    if (i == 1)
                                    {
                                        structurePort[0] = "Struct_2";
                                        structurePort[1] = "Structure";
                                    }
                                }
                            }
                        }
                    }
                    else
                        structurePort[1] = "Structure";
                }
                //switch the OffsetApplied flag
                if (structurePort[0] == "Struct_2")
                {
                    Boolean flag = IsOffsetApplied[0];
                    IsOffsetApplied[0] = IsOffsetApplied[1];
                    IsOffsetApplied[1] = flag;
                }

                return structurePort;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in GetIndexedStructPortName Method of Power_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }



    }
}
