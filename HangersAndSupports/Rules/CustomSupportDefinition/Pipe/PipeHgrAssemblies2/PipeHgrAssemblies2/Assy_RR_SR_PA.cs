//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_RR_SR_PA.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_RR_SR_PA
//   Author       : Vinay
//   Creation Date: 09-Sep-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   01-10-2015      Vinay   DI-CP-276996	Update HS_Assembly2 to use RichSteel
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
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

    public class Assy_RR_SR_PA : CustomSupportDefinition
    {
        // For everything
        private const string PIPE_CLAMP = "PIPE_CLAMP";
        private const string LOW_EYE_NUT = "LOW_EYE_NUT";
        private const string ROD1 = "ROD1";
        private const string NUT1 = "NUT1";

        // For C_CLAMP
        private const string C_CLAMP = "C_CLAMP";
        private const string ROD2 = "ROD2";
        private const string TB = "TB";
        private const string NUT2 = "NUT2";
        private const string NUT3 = "NUT3";

        // For LUG/CLEVIS
        private const string LUG = "LUG";
        private const string CLEVIS = "CLEVIS";
        private const string LUG_ROD2 = "LUG_ROD2";
        private const string LUG_TB = "LUG_TB";
        private const string LUG_NUT2 = "LUG_NUT2";
        private const string LUG_NUT3 = "LUG_NUT3";

        // For BEAM_CLAMP
        private const string BEAM_CLAMP = "BEAM_CLAMP";

        // For ROD_BEAM_ATT
        private const string BEAM_ATT = "BEAM_ATT";
        private const string EYE_NUT = "EYE_NUT";
        private const string ATT_ROD2 = "ATT_ROD2";
        private const string ATT_TB = "ATT_TB";
        private const string ATT_NUT2 = "ATT_NUT2";
        private const string ATT_NUT3 = "ATT_NUT3";

        // For ROD_WASHER
        private const string WASHER = "WASHER";
        private const string NUT4 = "NUT4";
        private const string NUT5 = "NUT5";
        private const string CONNECTION = "CONNECTION";
        private const string WASH_ROD2 = "WASH_ROD2";
        private const string WASH_TB = "WASH_TB";
        private const string WASH_NUT2 = "WASH_NUT2";
        private const string WASH_NUT3 = "WASH_NUT3";

        // For ROD_NUT
        private const string NUT_CONNECTION = "NUT_CONNECTION";
        private const string NUT_NUT4 = "NUT_NUT4";
        private const string NUT_NUT5 = "NUT_NUT5";
        private const string NUT_ROD2 = "NUT_ROD2";
        private const string NUT_TB = "NUT_TB";
        private const string NUT_NUT2 = "NUT_NUT2";
        private const string NUT_NUT3 = "NUT_NUT3";

        private const double CONST_INCH = 1000.0 / 25.4;
        String topType, rodType, pipeAttType;
        int turnbuckle;
        Double botRodLength;

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
                    BusinessObject supportPart = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                    rodType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyRodType", "ROD_TYPE")).PropValue;
                    topType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyTopType", "TOP_TYPE")).PropValue;
                    pipeAttType = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAHSA_PipeAtt", "PipeAtt")).PropValue;
                    turnbuckle = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyTurnbuckle", "TURNBUCKLE")).PropValue;
                    botRodLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyRR", "BOT_ROD_LENGTH")).PropValue;
                    
                        // Set it up for C-Clamp Usage ****************************************************************************************************
                        if (topType == "ROD_C_CLAMP")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(C_CLAMP, "Anv10_C_Clamp_86"));
                                parts.Add(new PartInfo(ROD2, "Anv10_RodETRL"));
                                parts.Add(new PartInfo(TB, "Anv10_Turnbuckle"));
                                parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                                parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(C_CLAMP, "Anv10_C_Clamp_86"));
                            }
                        }

                        // Set it up for Lug/Clevis Usage ****************************************************************************************************
                        if (topType == "ROD_CLEVIS_LUG")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(LUG, "Anv10_ShortStructLug"));
                                parts.Add(new PartInfo(CLEVIS, "Anv10_ClevisWithPin"));
                                parts.Add(new PartInfo(LUG_ROD2, "Anv10_RodETRL"));
                                parts.Add(new PartInfo(LUG_TB, "Anv10_Turnbuckle"));
                                parts.Add(new PartInfo(LUG_NUT2, "Anv10_HexNut"));
                                parts.Add(new PartInfo(LUG_NUT3, "Anv10_HexNut"));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(LUG, "Anv10_ShortStructLug"));
                                parts.Add(new PartInfo(CLEVIS, "Anv10_ClevisWithPin"));
                            }
                        }

                        // Set it up for Rod/Beam Clamp Usage ****************************************************************************************************
                        if (topType == "ROD_BEAM_CLAMP")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(BEAM_CLAMP, "Anv10_MBeamClamp_292"));
                                parts.Add(new PartInfo(ROD2, "Anv10_RodETRL"));
                                parts.Add(new PartInfo(TB, "Anv10_Turnbuckle"));
                                parts.Add(new PartInfo(NUT2, "Anv10_HexNut"));
                                parts.Add(new PartInfo(NUT3, "Anv10_HexNut"));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(BEAM_CLAMP, "Anv10_MBeamClamp_292"));
                            }
                        }

                        // Set it up for ROD_BEAM_ATT Usage ****************************************************************************************************
                        if (topType == "ROD_BEAM_ATT")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(BEAM_ATT, "Anv10_WBABolt"));
                                parts.Add(new PartInfo(EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ATT_ROD2, "Anv10_RodETRL"));
                                parts.Add(new PartInfo(ATT_TB, "Anv10_Turnbuckle"));
                                parts.Add(new PartInfo(ATT_NUT2, "Anv10_HexNut"));
                                parts.Add(new PartInfo(ATT_NUT3, "Anv10_HexNut"));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(BEAM_ATT, "Anv10_WBABolt"));
                                parts.Add(new PartInfo(EYE_NUT, "Anv10_EyeNut"));
                            }
                        }

                        // Set it up for ROD_WASHER Usage ****************************************************************************************************
                        if (topType == "ROD_WASHER")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(WASHER, "Anv10_WasherPlate"));
                                parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
                                parts.Add(new PartInfo(NUT5, "Anv10_HexNut"));
                                parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                                parts.Add(new PartInfo(WASH_ROD2, "Anv10_RodETRL"));
                                parts.Add(new PartInfo(WASH_TB, "Anv10_Turnbuckle"));
                                parts.Add(new PartInfo(WASH_NUT2, "Anv10_HexNut"));
                                parts.Add(new PartInfo(WASH_NUT3, "Anv10_HexNut"));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(WASHER, "Anv10_WasherPlate"));
                                parts.Add(new PartInfo(NUT4, "Anv10_HexNut"));
                                parts.Add(new PartInfo(NUT5, "Anv10_HexNut"));
                                parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                            }
                        }

                        // Set it up for ROD_NUT Usage ****************************************************************************************************
                        if (topType == "ROD_NUT")
                        {
                            if (turnbuckle == 1) // 1 means With Turnbuckle
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(NUT_CONNECTION, "Log_Conn_Part_1"));
                                parts.Add(new PartInfo(NUT_NUT4, "Anv10_HexNut"));
                                parts.Add(new PartInfo(NUT_NUT5, "Anv10_HexNut"));
                                parts.Add(new PartInfo(NUT_ROD2, "Anv10_RodETRL"));
                                parts.Add(new PartInfo(NUT_TB, "Anv10_Turnbuckle"));
                                parts.Add(new PartInfo(NUT_NUT2, "Anv10_HexNut"));
                                parts.Add(new PartInfo(NUT_NUT3, "Anv10_HexNut"));
                            }
                            else
                            {
                                parts.Add(new PartInfo(PIPE_CLAMP, pipeAttType));
                                parts.Add(new PartInfo(LOW_EYE_NUT, "Anv10_EyeNut"));
                                parts.Add(new PartInfo(ROD1, rodType));
                                parts.Add(new PartInfo(NUT1, "Anv10_HexNut"));
                                parts.Add(new PartInfo(NUT_CONNECTION, "Log_Conn_Part_1"));
                                parts.Add(new PartInfo(NUT_NUT4, "Anv10_HexNut"));
                                parts.Add(new PartInfo(NUT_NUT5, "Anv10_HexNut"));
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
                
                // Get interface for accessing items on the collection of Part Occurences
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject lowEyeNutPart = componentDictionary[LOW_EYE_NUT].GetRelationship("madeFrom", "part").TargetObjects[0];

                // Check the Route sloped
                Double routeAngle, nutPosition1, flangeThickness = 0.0;
                Boolean slopedRoute=false;
                String routePort;

                routeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                routeAngle = 90 - routeAngle * 180 / Math.PI;

                if (HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle, 3) , 0)==false)
                    slopedRoute = true;

                if (slopedRoute == true)
                    routePort = "Route";
                else
                    routePort = "RouteAlt";

                double overLength1 = (double)((PropertyValueDouble)lowEyeNutPart.GetPropertyValue("IJUAhsOverLength1", "OverLength1")).PropValue;
                double shapeLength1 = (double)((PropertyValueDouble)lowEyeNutPart.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                nutPosition1 = overLength1 + shapeLength1;

                BusinessObject nut1 = componentDictionary[NUT1].GetRelationship("madeFrom", "part").TargetObjects[0];
                double nutThichness = (double)((PropertyValueDouble)nut1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;
                // Start the Joints here ********************************************************************************************************************
                // Add the intelligence to determine which side of the pipe the steel is when placing by point
                Double byPointAngle1, byPointAngle2 = 0.0;
                Plane planeA, planeB, planeC, planeD;
                Axis axisA, axisB;

                byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

                if (Math.Abs(byPointAngle2) > Math.Round(Math.PI, 7) / 2.0)  // The structure is oriented in the standard direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 7) / 2.0)
                    {
                        planeA = Plane.XY;
                        planeB = Plane.ZX;
                        axisA = Axis.X;
                        axisB = Axis.X;

                        planeC = Plane.XY;
                        planeD = Plane.NegativeZX;
                    }
                    else
                    {
                        planeA = Plane.XY;
                        planeB = Plane.NegativeZX;
                        axisA = Axis.X;
                        axisB = Axis.X;

                        planeC = Plane.XY;
                        planeD = Plane.ZX;
                    }
                }
                else   // The structure is oriented in the opposite direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 7) / 2.0)
                    {
                        planeA = Plane.XY;
                        planeB = Plane.ZX;
                        axisA = Axis.X;
                        axisB = Axis.X;

                        planeC = Plane.XY;
                        planeD = Plane.ZX;
                    }
                    else
                    {
                        planeA = Plane.XY;
                        planeB = Plane.NegativeZX;
                        axisA = Axis.X;
                        axisB = Axis.X;

                        planeC = Plane.XY;
                        planeD = Plane.NegativeZX;
                    }
                }

                // Add a Revolute Joint between Eye Nut and Pipe Clamp
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreateRevoluteJoint(LOW_EYE_NUT, "Eye", PIPE_CLAMP, "Wing", Axis.X, Axis.X);
                else
                    JointHelper.CreateRigidJoint(LOW_EYE_NUT, "Eye", PIPE_CLAMP, "Wing", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                // Add a Vertical Joint to the Rod Z axis

                JointHelper.CreateGlobalAxesAlignedJoint(ROD1, "RodEnd1", Axis.Z, Axis.Z);

                // Create the Flexible (Prismatic) Joint between the ports of the bottom rod

                JointHelper.CreatePrismaticJoint(ROD1, "RodEnd1", ROD1, "RodEnd2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                // Set it up for C-Clamp Usage
                 if (topType == "ROD_C_CLAMP")
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        flangeThickness = 0.02;
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                            flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                        else
                            flangeThickness = 0.02;
                    }
                    // Joint from Pipe Clamp to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);

                        JointHelper.CreatePrismaticJoint("-1", "Structure", C_CLAMP, "Bottom", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, flangeThickness / 2, 0);

                        JointHelper.CreateRevoluteJoint(ROD1, "RodEnd2", C_CLAMP, "Structure", Axis.Z, Axis.Z);
                    }
                    else if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        JointHelper.CreateRigidJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        JointHelper.CreatePrismaticJoint("-1", "Structure", C_CLAMP, "Bottom", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, flangeThickness / 2, 0);

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd2", C_CLAMP, "RodEnd1", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.Z, 0, 0, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(PIPE_CLAMP, "Route", "-1", "Route", planeA, planeB, axisA, axisB, 0, 0, 0);

                        JointHelper.CreatePlanarJoint(C_CLAMP, "Structure", "-1", "Structure", planeC, planeD, -flangeThickness / 2);

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd2", C_CLAMP, "RodEnd1", Plane.ZX, Plane.NegativeYZ, Axis.X, Axis.Y, 0, 0, 0);
                    }

                    if (turnbuckle == 1) // 1 means With Turnbuckle
                    {
                        componentDictionary[ROD2].SetPropertyValue(botRodLength, "IJOAHgrOccLength", "Length");

                        // Add a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(ROD2, "RodEnd1", Axis.Z, Axis.Z);

                        // Add a Rigid Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a revolute Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(ROD2, "RodEnd2", TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        // Add a revolute Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        BusinessObject turnBucklePart1 = componentDictionary[TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition = (opening - rodTakeOut) / 2 + shapeLength;

                        JointHelper.CreateRigidJoint(ROD2, "RodEnd2", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);
                    }
                    else
                    {
                        // Add a Rigid Joint between bottom rod and Low Eye Nut
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        else
                            JointHelper.CreateRevoluteJoint(LOW_EYE_NUT, "RodEnd", ROD1, "RodEnd1", Axis.X, Axis.NegativeY);

                        // Add a Rigid Joint between the bottom nut and the rod
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                    }
                }

                // Set it up for Clevis/Lug Usage
                if (topType == "ROD_CLEVIS_LUG")
                {
                    // Joint from Pipe Clamp to Pipe

                    JointHelper.CreateRigidJoint("-1", routePort, PIPE_CLAMP, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    // Add a Planar Joint between the lug and the Structure

                    JointHelper.CreatePlanarJoint(LUG, "Hole2", "-1", "Structure", Plane.XY, Plane.NegativeXY, 0);

                    // Add a revolute Joint between the lug hole and clevis pin

                    JointHelper.CreateRevoluteJoint(CLEVIS, "Pin", LUG, "Hole1", Axis.X, Axis.X);

                    // Add a Joint between top of the rod and the Clevis
                    if (turnbuckle == 1)
                    {
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint(ROD1, "RodEnd2", CLEVIS, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        else
                            JointHelper.CreateRigidJoint(ROD1, "RodEnd2", CLEVIS, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                    }
                    else
                    {
                        if (Configuration == 2)
                            JointHelper.CreateRigidJoint(ROD1, "RodEnd2", CLEVIS, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        else
                            JointHelper.CreateRigidJoint(ROD1, "RodEnd2", CLEVIS, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                    }
                    if (turnbuckle == 1) // 1 means With Turnbuckle
                    {
                        componentDictionary[LUG_ROD2].SetPropertyValue(botRodLength, "IJOAHgrOccLength", "Length");

                        // Add a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(LUG_ROD2, "RodEnd1", Axis.Z, Axis.Z);

                        // Add a Rigid Joint between bottom rod and Low Eye Nut

                        if (slopedRoute == false)
                        {
                            if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                                JointHelper.CreateRigidJoint(LUG_ROD2, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            else
                                JointHelper.CreateRigidJoint(LUG_ROD2, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        }
                        else
                            JointHelper.CreateRigidJoint(LUG_ROD2, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(LUG_ROD2, "RodEnd2", LUG_TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a revolute Joint between bottom rod and Turnbuckle
                        JointHelper.CreateRigidJoint(LUG_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.Y, Axis.Y, 0, 0, 0);
                            
                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(LUG_ROD2, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        BusinessObject turnBucklePart1 = componentDictionary[LUG_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition = (opening - rodTakeOut) / 2 + shapeLength;

                        JointHelper.CreateRigidJoint(LUG_ROD2, "RodEnd2", LUG_NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LUG_NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);
                    }
                    else
                    {
                        // Add a Rigid Joint between bottom rod and Low Eye Nut

                        if (slopedRoute == false)
                        {
                            if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                                JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                            else
                                JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        }
                        else
                        
                            // Add a Rigid Joint between bottom rod and Low Eye Nut

                            JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                            // Add a Rigid Joint between the bottom nut and the rod

                            JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        
                    }
                }

                // Set it up for Beam Clamp Usage
                if (topType == "ROD_BEAM_CLAMP")
                {
                    // Joint from Pipe Clamp to Pipe and other featured joints
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);

                        JointHelper.CreatePrismaticJoint(BEAM_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0);

                        JointHelper.CreateRevoluteJoint(ROD1, "RodEnd2", BEAM_CLAMP, "RodEnd", Axis.Z, Axis.Z);
                    }
                    else
                    {
                        JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", routePort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);

                        JointHelper.CreatePrismaticJoint(BEAM_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeY, 0, 0);

                        JointHelper.CreateRevoluteJoint(ROD1, "RodEnd2", BEAM_CLAMP, "RodEnd", Axis.Z, Axis.Z);
                    }

                    if (turnbuckle == 1) // 1 means With Turnbuckle
                    {
                        componentDictionary[ROD2].SetPropertyValue(botRodLength, "IJOAHgrOccLength", "Length");

                        // Add a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(ROD2, "RodEnd1", Axis.Z, Axis.Z);

                        // Add a Rigid Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a revolute Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(ROD2, "RodEnd2", TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        // Add a revolute Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD2, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        BusinessObject turnBucklePart1 = componentDictionary[TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition = (opening - rodTakeOut) / 2 + shapeLength;

                        JointHelper.CreateRigidJoint(ROD2, "RodEnd2", NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);
                    }
                    else
                    {
                        // Add a Rigid Joint between bottom rod and Low Eye Nut

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                    }
                 }

                // Set it up for Beam Att/Eye Nut Usage
                if (topType == "ROD_BEAM_ATT")
                {
                    // Joint from Pipe Clamp to Pipe
                    JointHelper.CreateRigidJoint("-1", routePort, PIPE_CLAMP, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    // Add a Planar Joint between the lug and the Structure

                    JointHelper.CreatePlanarJoint(BEAM_ATT, "Structure", "-1", "Structure", Plane.XY, Plane.XY, 0);

                    // Add a revolute Joint between the Eye nut and Beam Attachment

                    JointHelper.CreateRevoluteJoint(EYE_NUT, "Eye", BEAM_ATT, "Pin", Axis.Y, Axis.Y);

                    // Add a Rigid Joint between top of the rod and the eye nut
                    if (Configuration == 1)
                         JointHelper.CreateRigidJoint(ROD1, "RodEnd2", EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd2", EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);


                    if (turnbuckle == 1)
                    {
                        componentDictionary[ATT_ROD2].SetPropertyValue(botRodLength, "IJOAHgrOccLength", "Length");

                        // Add a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(ATT_ROD2, "RodEnd1", Axis.Z, Axis.Z);

                        // Add a Rigid Joint between bottom rod and Low Eye Nut

                        if (slopedRoute == false)
                        {
                            if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                                JointHelper.CreateRigidJoint(ATT_ROD2, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            else
                                JointHelper.CreateRigidJoint(ATT_ROD2, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        }
                        else
                            JointHelper.CreateRigidJoint(ATT_ROD2, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
 
                        // Add a Rigid Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(ATT_ROD2, "RodEnd2", ATT_TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a revolute Joint between bottom rod and Turnbuckle
                        JointHelper.CreateRigidJoint(ATT_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ATT_ROD2, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        BusinessObject turnBucklePart1 = componentDictionary[ATT_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition = (opening - rodTakeOut) / 2 + shapeLength;

                        JointHelper.CreateRigidJoint(ATT_ROD2, "RodEnd2", ATT_NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", ATT_NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);
                    }
                    else
                    {
                        // Add a Rigid Joint between bottom rod and Low Eye Nut
                        if (slopedRoute == false)
                        {
                            if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                                JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            else
                                JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        }
                        else
                       
                            // Add a Rigid Joint between bottom rod and Low Eye Nut

                            JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                            // Add a Rigid Joint between the bottom nut and the rod

                            JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                        
                    }
                }

                // Set it up for Rod/Washer Usage
                if (topType == "ROD_WASHER")
                {
                    // Joint from Pipe Clamp to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRevoluteJoint(ROD1, "RodEnd2", CONNECTION, "Connection", Axis.Z, Axis.Z);

                        JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);

                        JointHelper.CreatePrismaticJoint(WASHER, "Port1", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, 0, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(ROD1, "RodEnd2", CONNECTION, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        JointHelper.CreateRigidJoint(PIPE_CLAMP, "Route", "-1", routePort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        JointHelper.CreatePlanarJoint(WASHER, "Port1", "-1", "Structure", Plane.XY, Plane.NegativeXY, 0);
                    }

                    // Add a Rigid Joint between the bottom nut and the rod

                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.03 + nutThichness, 0, 0);

                    // Add a Rigid Joint between the middle nut and the rod

                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.015 + nutThichness, 0, 0);

                    // Create the rigid joint to locate the connection objecct

                    JointHelper.CreateRigidJoint(CONNECTION, "Connection", WASHER, "Port1", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0.05, 0, 0);

                    if (turnbuckle == 1) // 1 means With Turnbuckle
                    {
                        componentDictionary[WASH_ROD2].SetPropertyValue(botRodLength, "IJOAHgrOccLength", "Length");

                        // Add a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(WASH_ROD2, "RodEnd1", Axis.Z, Axis.Z);

                        // Add a Rigid Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRigidJoint(WASH_ROD2, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a revolute Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(WASH_ROD2, "RodEnd2", WASH_TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        // Add a revolute Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(WASH_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(WASH_ROD2, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        BusinessObject turnBucklePart1 = componentDictionary[WASH_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition = (opening - rodTakeOut) / 2 + shapeLength;

                        JointHelper.CreateRigidJoint(WASH_ROD2, "RodEnd2", WASH_NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", WASH_NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);
                    }
                    else
                    {
                        // Add a Rigid Joint between bottom rod and Low Eye Nut

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                    }
                }

                // Set it up for Rod/NUT Usage
                if (topType == "ROD_NUT")
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        // NO STRUCTURE IS SELCTED, SHOULD GO WITH DEFAULT SIZE
                        flangeThickness = 0.02;
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                            flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                        else
                            flangeThickness = 0.02;
                    }
                     // Joint from Pipe Clamp to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(PIPE_CLAMP, "Route", "-1", routePort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    // Add a Prismatic Joint between the lug and the Structure
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint(NUT_CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, flangeThickness / 2 + 0.05, 0);
                    else
                        JointHelper.CreatePlanarJoint(NUT_CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.XY, flangeThickness / 2.0 + 0.05);

                    // Add a revolute Joint between top of the rod and the C Clamp

                    JointHelper.CreateRevoluteJoint(ROD1, "RodEnd2", NUT_CONNECTION, "Connection", Axis.Z, Axis.Z);

                    // Add a Rigid Joint between the bottom nut and the rod

                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", NUT_NUT4, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, flangeThickness / 2 + 0.05 + nutThichness, 0, 0);

                    // Add a Rigid Joint between the middle nut and the rod

                    JointHelper.CreateRigidJoint(ROD1, "RodEnd2", NUT_NUT5, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, -flangeThickness / 2 + 0.05, 0, 0);

                    if (turnbuckle == 1) // 1 means With Turnbuckle
                    {
                        componentDictionary[NUT_ROD2].SetPropertyValue(botRodLength, "IJOAHgrOccLength", "Length");

                        // Add a Vertical Joint to the Rod Z axis

                        JointHelper.CreateGlobalAxesAlignedJoint(NUT_ROD2, "RodEnd1", Axis.Z, Axis.Z);

                        // Add a Rigid Joint between bottom rod and Clevis Hanger

                        JointHelper.CreateRigidJoint(NUT_ROD2, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a revolute Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(NUT_ROD2, "RodEnd2", NUT_TB, "RodEnd2", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                        // Add a revolute Joint between bottom rod and Turnbuckle

                        JointHelper.CreateRigidJoint(NUT_TB, "RodEnd1", ROD1, "RodEnd1", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(NUT_ROD2, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        BusinessObject turnBucklePart1 = componentDictionary[NUT_TB].GetRelationship("madeFrom", "part").TargetObjects[0];
                        double opening = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsOpening1", "Opening1")).PropValue;
                        double rodTakeOut = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsRodTakeOut", "RodTakeOut")).PropValue;
                        double shapeLength = (double)((PropertyValueDouble)turnBucklePart1.GetPropertyValue("IJUAhsShape1", "Shape1Length")).PropValue;

                        double nutPosition = (opening - rodTakeOut) / 2 + shapeLength;

                        JointHelper.CreateRigidJoint(NUT_ROD2, "RodEnd2", NUT_NUT2, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT_NUT3, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition + nutThichness, 0, 0);
                    }
                    else
                    {
                        // Add a Rigid Joint between bottom rod and Low Eye Nut

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", LOW_EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        // Add a Rigid Joint between the bottom nut and the rod

                        JointHelper.CreateRigidJoint(ROD1, "RodEnd1", NUT1, "Top", Plane.XY, Plane.XY, Axis.X, Axis.X, nutPosition1 + nutThichness, 0, 0);
                    }
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
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(PIPE_CLAMP, 1)); // partindex, routeindex

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

                    if (topType == "ROD_C_CLAMP")
                        structConnections.Add(new ConnectionInfo(C_CLAMP, 1)); // partindex, routeindex
                    if (topType == "ROD_CLEVIS_LUG")
                        structConnections.Add(new ConnectionInfo(LUG, 1));
                    if (topType == "ROD_BEAM_CLAMP")
                        structConnections.Add(new ConnectionInfo(BEAM_CLAMP, 1));
                    if (topType == "ROD_BEAM_ATT")
                        structConnections.Add(new ConnectionInfo(BEAM_ATT, 1));
                    if (topType == "ROD_WASHER")
                        structConnections.Add(new ConnectionInfo(WASHER, 1));
                    if (topType == "ROD_NUT")
                        structConnections.Add(new ConnectionInfo(NUT_CONNECTION, 1));

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
