//+++++++++++++++++++++++++++++ ++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_RR_SR_CL.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_VS_SR_MD
//   Author       : Hema
//   Creation Date: 04.April.2013
//   Description:   Converted HS_Assembly VB Project to C# .Net 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   04.April.2013   Hema     Converted HS_Assembly VB Project to C# .Net
//   22-Jan-2015     PVK      TR-CP-264951  Resolve coverity issues found in November 2014 report 
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
   
    public class Assy_VS_SR_MD : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_VS_SR_DB"
        //----------------------------------------------------------------------------------
        //Constants
        // For everything
        private const string PIPE_CLAMP = "PIPE_CLAMP";
        private const string ROD1 = "ROD1";
        private const string NUT1 = "NUT1";
        private const string CLAMP_EYE_NUT = "CLAMP_EYE_NUT";

        // For C_CLAMP
        private const string C_CLAMP = "C_CLAMP";
        private const string ROD2 = "ROD2";
        private const string SPRING = "SPRING";
        private const string NUT3 = "NUT3";
        private const string NUT4 = "NUT4";

        // For LUG/CLEVIS
        private const string LUG = "LUG";
        private const string CLEVIS = "CLEVIS";
        private const string LUG_ROD2 = "LUG_ROD2";
        private const string LUG_SPRING = "LUG_SPRING";
        private const string LUG_NUT3 = "LUG_NUT3";
        private const string LUG_NUT4 = "LUG_NUT4";

        // For BEAM_CLAMP
        private const string BEAM_CLAMP = "BEAM_CLAMP";

        // For ROD_BEAM_ATT
        private const string BEAM_ATT = "BEAM_ATT";
        private const string EYE_NUT = "EYE_NUT";
        private const string ATT_ROD2 = "ATT_ROD2";
        private const string ATT_SPRING = "ATT_SPRING";
        private const string ATT_NUT3 = "ATT_NUT3";
        private const string ATT_NUT4 = "ATT_NUT4";

        // For ROD_WASHER
        private const string WASHER = "WASHER";
        private const string NUT5 = "NUT5";
        private const string NUT6 = "NUT6";
        private const string CONNECTION = "CONNECTION";
        private const string WASH_ROD2 = "WASH_ROD2";
        private const string WASH_SPRING = "WASH_SPRING";
        private const string WASH_NUT3 = "WASH_NUT3";
        private const string WASH_NUT4 = "WASH_NUT4";

        // For ROD_NUT
        private const string NUT_CONNECTION = "NUT_CONNECTION";
        private const string NUT_NUT5 = "NUT_NUT5";
        private const string NUT_NUT6 = "NUT_NUT6";
        private const string NUT_ROD2 = "NUT_ROD2";
        private const string NUT_SPRING = "NUT_SPRING";
        private const string NUT_NUT3 = "NUT_NUT3";
        private const string NUT_NUT4 = "NUT_NUT4";

        //For ROD_BEAM_ATT_B
        private const string ATT_B_SPRING = "ATT_B_SPRING";
        private const string ATT_B_NUT3 = "ATT_B_NUT3";

        //For ROD_LUG_C
        private const string LUG_C_PIN = "LUG_C_PIN";
        private const string LUG_C_SPRING = "LUG_C_SPRING";
        private const string LUG_C_NUT3 = "LUG_C_NUT3";

        String topType, pinSize;
        
        Double topAssyLength, actualTravel, hotLoad;

        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    //Gets SupportHelper
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                    //Create a new collection to hold the caltalog parts
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    String rodType, beamClampSize, rodSize, springType;

                    rodType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyRodType", "ROD_TYPE")).PropValue;
                    topType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyTopType", "TOP_TYPE")).PropValue;
                    springType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyVS", "SPRING_TYPE")).PropValue;
                    topAssyLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyVS", "TOP_ASSY_LENGTH")).PropValue;
                    rodSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyVS", "ROD_SIZE")).PropValue;
                    beamClampSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyVS", "BEAM_CLAMP_SIZE")).PropValue;

                    //Added these lines to set the attributes in parts properties page - SS - TR 97938
                    actualTravel = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyVS", "WORKING_TRAV")).PropValue;
                    hotLoad = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyVS", "HOT_LOAD")).PropValue;

                    if (rodSize.Equals("8"))
                        pinSize = "8";
                    else
                        pinSize = (Convert.ToInt32(rodSize) - 1).ToString().Trim();

                    // Set it up for C-Clamp Usage ****************************************************************************************************
                    if (topType == "ROD_C_CLAMP")
                    {
                        parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG212"));
                        parts.Add(new PartInfo(ROD1, rodType + rodSize));
                        parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(CLAMP_EYE_NUT, "Anvil_FIG290_" + rodSize));
                        parts.Add(new PartInfo(C_CLAMP, "Anvil_FIG86_" + rodSize));
                        parts.Add(new PartInfo(ROD2, rodType + rodSize));
                        parts.Add(new PartInfo(NUT_SPRING, springType));
                        parts.Add(new PartInfo(NUT3, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(NUT4, "Anvil_HEX_NUT_" + rodSize));
                    }

                    // Set it up for Lug/Clevis Usage ****************************************************************************************************
                    if (topType == "ROD_CLEVIS_LUG")
                    {
                        parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG212"));
                        parts.Add(new PartInfo(ROD1, rodType + rodSize));
                        parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(CLAMP_EYE_NUT, "Anvil_FIG290_" + rodSize));
                        parts.Add(new PartInfo(LUG, "Anvil_FIG55S_" + rodSize));
                        parts.Add(new PartInfo(CLEVIS, "Anvil_FIG299_" + rodSize));
                        parts.Add(new PartInfo(LUG_ROD2, rodType + rodSize));
                        parts.Add(new PartInfo(NUT_SPRING, springType));
                        parts.Add(new PartInfo(LUG_NUT3, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(LUG_NUT4, "Anvil_HEX_NUT_" + rodSize));
                    }

                    // Set it up for Rod/Beam Clamp Usage ****************************************************************************************************
                    if (topType == "ROD_BEAM_CLAMP")
                    {
                        parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG212"));
                        parts.Add(new PartInfo(ROD1, rodType + rodSize));
                        parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(CLAMP_EYE_NUT, "Anvil_FIG290_" + rodSize));
                        parts.Add(new PartInfo(BEAM_CLAMP, "Anvil_FIG292_" + beamClampSize));
                        parts.Add(new PartInfo(ROD2, rodType + rodSize));
                        parts.Add(new PartInfo(NUT_SPRING, springType));
                        parts.Add(new PartInfo(NUT3, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(NUT4, "Anvil_HEX_NUT_" + rodSize));
                    }

                    // Set it up for ROD_BEAM_ATT Usage ****************************************************************************************************
                    if (topType == "ROD_BEAM_ATT")
                    {
                        parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG212"));
                        parts.Add(new PartInfo(ROD1, rodType + rodSize));
                        parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(CLAMP_EYE_NUT, "Anvil_FIG290_" + rodSize));
                        parts.Add(new PartInfo(BEAM_ATT, "Anvil_FIG66_" + rodSize));
                        parts.Add(new PartInfo(EYE_NUT, "Anvil_FIG290_" + rodSize));
                        parts.Add(new PartInfo(ATT_ROD2, rodType + rodSize));
                        parts.Add(new PartInfo(NUT_SPRING, springType));
                        parts.Add(new PartInfo(ATT_NUT3, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(ATT_NUT4, "Anvil_HEX_NUT_" + rodSize));
                    }

                    // Set it up for ROD_WASHER Usage ****************************************************************************************************
                    if (topType == "ROD_WASHER")
                    {
                        parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG212"));
                        parts.Add(new PartInfo(ROD1, rodType + rodSize));
                        parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(CLAMP_EYE_NUT, "Anvil_FIG290_" + rodSize));
                        parts.Add(new PartInfo(WASHER, "Anvil_FIG60_" + rodSize));
                        parts.Add(new PartInfo(NUT5, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(NUT6, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(CONNECTION, "Log_Conn_Part_1"));
                        parts.Add(new PartInfo(WASH_ROD2, rodType + rodSize));
                        parts.Add(new PartInfo(NUT_SPRING, springType));
                        parts.Add(new PartInfo(WASH_NUT3, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(WASH_NUT4, "Anvil_HEX_NUT_" + rodSize));
                    }

                    // Set it up for ROD_NUT Usage ****************************************************************************************************
                    if (topType == "ROD_NUT")
                    {
                        parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG212"));
                        parts.Add(new PartInfo(ROD1, rodType + rodSize));
                        parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(CLAMP_EYE_NUT, "Anvil_FIG290_" + rodSize));
                        parts.Add(new PartInfo(NUT_CONNECTION, "Log_Conn_Part_1"));
                        parts.Add(new PartInfo(NUT_NUT5, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(NUT_NUT6, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(NUT_ROD2, rodType + rodSize));
                        parts.Add(new PartInfo(NUT_SPRING, springType));
                        parts.Add(new PartInfo(NUT_NUT3, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(NUT_NUT4, "Anvil_HEX_NUT_" + rodSize));
                    }

                    // Set it up for ROD_BEAM_ATT_B Usage ****************************************************************************************************
                    if (topType == "ROD_BEAM_ATT_B")
                    {
                        parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG212"));
                        parts.Add(new PartInfo(ROD1, rodType + rodSize));
                        parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(CLAMP_EYE_NUT, "Anvil_FIG290_" + rodSize));
                        parts.Add(new PartInfo(BEAM_ATT, "Anvil_FIG66_" + rodSize));
                        parts.Add(new PartInfo(NUT_SPRING, springType));
                        parts.Add(new PartInfo(ATT_NUT3, "Anvil_HEX_NUT_" + rodSize));
                    }

                    // Set it up for Lug/Clevis Usage ****************************************************************************************************
                    if (topType == "ROD_LUG_C")
                    {
                        parts.Add(new PartInfo(PIPE_CLAMP, "Anvil_FIG212"));
                        parts.Add(new PartInfo(ROD1, rodType + rodSize));
                        parts.Add(new PartInfo(NUT1, "Anvil_HEX_NUT_" + rodSize));
                        parts.Add(new PartInfo(CLAMP_EYE_NUT, "Anvil_FIG290_" + rodSize));
                        parts.Add(new PartInfo(LUG, "Anvil_FIG55S_" + rodSize));
                        parts.Add(new PartInfo(LUG_C_PIN, "Anvil_FIG291_" + pinSize));
                        parts.Add(new PartInfo(NUT_SPRING, springType));
                        parts.Add(new PartInfo(LUG_C_NUT3, "Anvil_HEX_NUT_" + rodSize));
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
                //Get SupportHelper
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                //=== ========== == ==== ===========
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                //Get interface for accessing items on the collection of Part Occurences
                //Added these lines to get the codelisted value for travel direction - SS - TR 97938
                BusinessObject pipeClampPart = (componentDictionary[PIPE_CLAMP]).GetRelationship("madeFrom", "part").TargetObjects[0];
                int tempDir = ((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyVS", "DIR")).PropValue;

                //======================================================
                //Get the Current Location in the Route Connection Cycle
                //======================================================

                //Check the Route sloped
                Boolean slopedRoute = false;
                double byPointAngle1, byPointAngle2, topRodLength = 0, flangeThickness = 0.0, flangeWidthValue = 0, routeAngle;
                string routePort;

                routeAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                routeAngle = 90 - (routeAngle) * 180 / Math.PI;

               if (HgrCompareDoubleService.cmpdbl(Math.Round(routeAngle, 3) , 0)==false)
                    slopedRoute = true;

                if (slopedRoute == true)
                    routePort = "Route";
                else
                    routePort = "RouteAlt";

                //====== ======
                //Create Joints
                //====== ======

                //Add the intelligence to determine which side of the pipe the steel is when placing by point
                byPointAngle1 = RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y);
                byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

                Plane[] planearry = new Plane[2];

                if (Math.Abs(byPointAngle2) > Math.Round(Math.PI, 7) / 2.0)  //The structure is oriented in the standard direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 7) / 2.0)
                    {
                        planearry[0] = Plane.XY;
                        planearry[1] = Plane.ZX;
                    }
                    else
                    {
                        planearry[0] = Plane.XY;
                        planearry[1] = Plane.NegativeZX;
                    }
                }
                else   //The structure is oriented in the opposite direction
                {
                    if (Math.Abs(byPointAngle1) < Math.Round(Math.PI, 7) / 2.0)
                    {
                        planearry[0] = Plane.XY;
                        planearry[1] = Plane.NegativeZX;
                    }
                    else
                    {
                        planearry[0] = Plane.XY;
                        planearry[1] = Plane.ZX;
                    }
                }

                //Setting the HOT_LOAD,WORKING_TRAV,DIR values
                //Added these lines to set the attributes in parts properties page - SS - TR 97938

                String[] collectionOfInterfaces = new String[] { "IJOAHgrAnvil_FIG82A", "IJOAHgrAnvil_FIG82B", "IJOAHgrAnvil_FIG82C", "IJOAHgrAnvil_FIG98A", "IJOAHgrAnvil_FIG98B", "IJOAHgrAnvil_FIG98C", "IJOAHgrAnvil_FIGB268A", "IJOAHgrAnvil_FIGB268B", "IJOAHgrAnvil_FIGB268C" };
                CommonAssembly commonAssembly = new CommonAssembly();

                (componentDictionary[NUT_SPRING]).SetPropertyValue(hotLoad, "IJOAHgrHotLoad", "HOT_LOAD");
                commonAssembly.SetPropertyFromObject(componentDictionary[NUT_SPRING], collectionOfInterfaces, "WORKING_TRAV", actualTravel);
                commonAssembly.SetPropertyFromObject(componentDictionary[NUT_SPRING], collectionOfInterfaces, "DIR", tempDir);

                //Start the Joints here ********************************************************************************************************************

                //Add a Revolute Joint between Eye Nut and Pipe Clamp
                JointHelper.CreateRevoluteJoint(CLAMP_EYE_NUT, "Eye", PIPE_CLAMP, "Pin", Axis.X, Axis.Y);

                //Add a Vertical Joint to the Rod Z axis
                JointHelper.CreateGlobalAxesAlignedJoint(ROD1, "TopExThdRH", Axis.Z, Axis.Z);

                //Set it up for C-Clamp Usage
                if (topType == "ROD_C_CLAMP")
                {
                    //Raise the error for Place By reference
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "C-Clamp Top Type can not be placed by 'Place By Reference' Command", "", "Assy_VS_SR_MD", 1);
                        return;
                    }
                    if ((SupportHelper.SupportingObjects.Count != 0))
                        flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;

                    (componentDictionary[C_CLAMP]).SetPropertyValue(flangeThickness, "IJOAHgrAnvil_FIG86", "FLANGE_T");

                    topRodLength = topAssyLength;

                    //Joint from Pipe Clamp to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                        JointHelper.CreatePrismaticJoint(C_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, 0);
                        JointHelper.CreateRevoluteJoint(ROD1, "TopExThdRH", C_CLAMP, "InThdRH", Axis.Z, Axis.Z);
                    }
                    else
                    {
                        JointHelper.CreateRevoluteJoint(PIPE_CLAMP, "Route", "-1", "Route", Axis.X, Axis.X);
                        JointHelper.CreatePlanarJoint(C_CLAMP, "Structure", "-1", "Structure", planearry[0], planearry[1], 0);
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", C_CLAMP, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                    }

                    //Create the Flexible (Prismatic) Joint between the ports of the bottom rod
                    JointHelper.CreatePrismaticJoint(ROD2, "TopExThdRH", ROD2, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                    //Add a Vertical Joint to the Rod Z axis
                    JointHelper.CreateGlobalAxesAlignedJoint(ROD2, "TopExThdRH", Axis.Z, Axis.Z);

                    //Add a Rigid Joint between bottom rod and Eye nut
                    JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a revolute Joint between bottom rod and Turnbuckle
                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", NUT_SPRING, "BotInThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a revolute Joint between bottom rod and Turnbuckle
                    JointHelper.CreateRigidJoint(NUT_SPRING, "TopInThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    (componentDictionary[ROD1]).SetPropertyValue(topRodLength, "IJUAHgrOccLength", "Length");

                    //Add a Rigid Joint between the bottom nuts and the rods
                    JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                }

                //Set it up for Clevis/Lug Usage
                if (topType == "ROD_CLEVIS_LUG")
                {
                    BusinessObject lugPart = (componentDictionary[LUG]).GetRelationship("madeFrom", "part").TargetObjects[0];
                    BusinessObject clevisPart = (componentDictionary[CLEVIS]).GetRelationship("madeFrom", "part").TargetObjects[0];

                    topRodLength = topAssyLength - ((double)((PropertyValueDouble)lugPart.GetPropertyValue("IJUAHgrAnvil_FIG55S", "H")).PropValue + (double)((PropertyValueDouble)lugPart.GetPropertyValue("IJUAHgrAnvil_FIG55S", "F")).PropValue / 2.0) - ((double)((PropertyValueDouble)clevisPart.GetPropertyValue("IJUAHgrAnvil_fig299", "A")).PropValue - (double)((PropertyValueDouble)clevisPart.GetPropertyValue("IJUAHgrRod_Dia", "ROD_DIA")).PropValue / 2.0 - (double)((PropertyValueDouble)clevisPart.GetPropertyValue("IJUAHgrAnvil_fig299", "P")).PropValue / 2.0);

                    //Joint from Pipe Clamp to Pipe
                    JointHelper.CreateRigidJoint("-1", routePort, PIPE_CLAMP, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Planar Joint between the lug and the Structure
                    JointHelper.CreatePlanarJoint(LUG, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, 0);

                    //Add a Vertical Joint to the Rod Z axis
                    JointHelper.CreateGlobalAxesAlignedJoint(LUG_ROD2, "TopExThdRH", Axis.Z, Axis.Z);

                    //Create the Flexible (Prismatic) Joint between the ports of the bottom rod
                    JointHelper.CreatePrismaticJoint(LUG_ROD2, "TopExThdRH", LUG_ROD2, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                    //Add a revolute Joint between the lug hole and clevis pin
                    if (Configuration == 2)
                        JointHelper.CreateRevoluteJoint(CLEVIS, "Pin", LUG, "Hole", Axis.X, Axis.Y);
                    else
                        JointHelper.CreateRevoluteJoint(CLEVIS, "Pin", LUG, "Hole", Axis.Y, Axis.NegativeX);

                    //Add a Joint between top of the rod and the Clevis
                    if (Configuration == 1)
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", CLEVIS, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", CLEVIS, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                    //Add a Rigid Joint between bottom rod and Eye nut
                    if (slopedRoute == false)
                        if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                            JointHelper.CreateRigidJoint(LUG_ROD2, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        else
                            JointHelper.CreateRigidJoint(LUG_ROD2, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(LUG_ROD2, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                    //Add a Rigid Joint between bottom rod and Turnbuckle
                    JointHelper.CreateRigidJoint(LUG_ROD2, "TopExThdRH", NUT_SPRING, "BotInThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Rigid Joint between bottom rod and Turnbuckle
                    JointHelper.CreateRigidJoint(NUT_SPRING, "TopInThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    (componentDictionary[ROD1]).SetPropertyValue(topRodLength, "IJUAHgrOccLength", "Length");

                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(LUG_ROD2, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(LUG_ROD2, "TopExThdRH", LUG_NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", LUG_NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                }

                //Set it up for Beam Clamp Usage
                if (topType == "ROD_BEAM_CLAMP")
                {
                    //We need to get the Flange Width.  However, it comes in metres, so we change it to inches.  Then we Int it to get rid of decimals.
                    //Then we add one so that the Flange Width of the clamp is bigger than the actual flange width, not smaller
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        flangeWidthValue = 3;
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                            flangeWidthValue = (int)(SupportingHelper.SupportingObjectInfo(1).Width * 1000.0 / 25.4) + 1.0;
                    }

                    //The lowest allowable FLANGE_W is 3 and the max is 15 
                    if (flangeWidthValue < 3)
                        flangeWidthValue = 3;
                    else
                        if (flangeWidthValue > 15)
                            flangeWidthValue = 15;

                    //We need to pass Flange Width into the Beam Clamp as a codelist Index so we get the Index from the Actual below
                    topRodLength = topAssyLength;

                    CodelistItem beamClampFWCodelist = (componentDictionary[BEAM_CLAMP]).GetPropertyValue("IJOAHgrAnvil_FIG292", "FLANGE_WIDTH").PropertyInfo.CodeListInfo.GetCodelistItem(flangeWidthValue.ToString().Trim());
                    (componentDictionary[BEAM_CLAMP]).SetPropertyValue(beamClampFWCodelist.Value, "IJOAHgrAnvil_FIG292", "FLANGE_WIDTH");

                    //Joint from Pipe Clamp to Pipe
                    //Changed to Prismatic joint for TR 87508
                    JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);

                    //Add a revolute Joint between top of the rod and the C Clamp
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint(BEAM_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, 0);
                    else
                        //Changed to Spherical joint for TR 87508
                        JointHelper.CreatePrismaticJoint(BEAM_CLAMP, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.NegativeY, 0, 0);

                    //Joint from Pipe Clamp to Pipe and other featured joints
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRevoluteJoint(ROD1, "TopExThdRH", BEAM_CLAMP, "InThdRH", Axis.Z, Axis.Z);
                    else
                        JointHelper.CreateSphericalJoint(ROD1, "TopExThdRH", BEAM_CLAMP, "InThdRH");

                    //Create the Flexible (Prismatic) Joint between the ports of the bottom rod
                    JointHelper.CreatePrismaticJoint(ROD2, "TopExThdRH", ROD2, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                    //Add a Vertical Joint to the Rod Z axis
                    JointHelper.CreateGlobalAxesAlignedJoint(ROD2, "TopExThdRH", Axis.Z, Axis.Z);

                    //Add a Rigid Joint between bottom rod and Clevis Hanger
                    JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Rigid Joint between bottom rod and Turnbuckle
                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", NUT_SPRING, "BotInThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Rigid Joint between bottom rod and Turnbuckle
                    JointHelper.CreateRigidJoint(NUT_SPRING, "TopInThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    (componentDictionary[ROD1]).SetPropertyValue(topRodLength, "IJUAHgrOccLength", "Length");

                    //Add a Rigid Joint between the bottom nuts and the rods
                    JointHelper.CreateRigidJoint(ROD2, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD2, "TopExThdRH", NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                }

                //Set it up for Beam Att/Eye Nut Usage
                if (topType == "ROD_BEAM_ATT")
                {
                    BusinessObject beamAttPart = (componentDictionary[BEAM_ATT]).GetRelationship("madeFrom", "part").TargetObjects[0];
                    BusinessObject eyeNutPart = (componentDictionary[EYE_NUT]).GetRelationship("madeFrom", "part").TargetObjects[0];

                    topRodLength = topAssyLength - ((double)((PropertyValueDouble)beamAttPart.GetPropertyValue("IJUAHgrAnvil_fig66", "E_PRIME")).PropValue - (double)((PropertyValueDouble)beamAttPart.GetPropertyValue("IJUAHgrAnvil_fig66", "H")).PropValue / 2.0) - ((double)((PropertyValueDouble)eyeNutPart.GetPropertyValue("IJUAHgrAnvil_FIG290", "E")).PropValue - (double)((PropertyValueDouble)eyeNutPart.GetPropertyValue("IJUAHgrRod_Dia", "ROD_DIA")).PropValue / 2.0);

                    (componentDictionary[ROD1]).SetPropertyValue(topRodLength, "IJUAHgrOccLength", "Length");

                    //Joint from Pipe Clamp to Pipe
                    JointHelper.CreateRigidJoint("-1", routePort, PIPE_CLAMP, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Planar Joint between the lug and the Structure
                    JointHelper.CreatePlanarJoint(BEAM_ATT, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, 0);

                    //Add a revolute Joint between the lug hole and clevis pin
                    if (Configuration == 1)
                        JointHelper.CreateRevoluteJoint(EYE_NUT, "Eye", BEAM_ATT, "Pin", Axis.X, Axis.Y);
                    else
                        JointHelper.CreateRevoluteJoint(EYE_NUT, "Eye", BEAM_ATT, "Pin", Axis.Y, Axis.NegativeX);

                    //Add a revolute Joint between top of the rod and the eye nut
                    if (Configuration == 2)
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                    //Add a Rigid Joint between bottom rod and Turnbuckle
                    JointHelper.CreateRigidJoint(NUT_SPRING, "TopInThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Rigid Joint between bottom rod and Turnbuckle
                    JointHelper.CreateRigidJoint(ATT_ROD2, "TopExThdRH", NUT_SPRING, "BotInThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Create the Flexible (Prismatic) Joint between the ports of the bottom rod
                    JointHelper.CreatePrismaticJoint(ATT_ROD2, "TopExThdRH", ATT_ROD2, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                    //Add a Vertical Joint to the Rod Z axis
                    JointHelper.CreateGlobalAxesAlignedJoint(ATT_ROD2, "TopExThdRH", Axis.Z, Axis.Z);

                    //Add a Rigid Joint between bottom rod and Clevis Hanger
                    if (slopedRoute == false)
                        if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                            JointHelper.CreateRigidJoint(ATT_ROD2, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        else
                            JointHelper.CreateRigidJoint(ATT_ROD2, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(LUG_ROD2, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ATT_ROD2, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ATT_ROD2, "TopExThdRH", ATT_NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", ATT_NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                }

                //Set it up for Rod/Washer Usage
                if (topType == "ROD_WASHER")
                {
                    topRodLength = topAssyLength;

                    (componentDictionary[ROD1]).SetPropertyValue(topRodLength, "IJUAHgrOccLength", "Length");

                    //Joint from Pipe Clamp to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreateRevoluteJoint(ROD1, "TopExThdRH", CONNECTION, "Connection", Axis.Z, Axis.Z);
                        JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                        JointHelper.CreatePrismaticJoint(WASHER, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, 0, 0);
                        JointHelper.CreateRigidJoint(CONNECTION, "Connection", WASHER, "OppStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.05, 0, 0);
                    }
                    else
                    {
                        JointHelper.CreatePlanarJoint(CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.XY, 0.05);
                        JointHelper.CreateSphericalJoint(ROD1, "TopExThdRH", CONNECTION, "Connection");
                        JointHelper.CreateRigidJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateRigidJoint(WASHER, "Structure", ROD1, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.05, 0, 0);
                    }
                    //Create the Flexible (Prismatic) Joint between the ports of the bottom rod
                    JointHelper.CreatePrismaticJoint(WASH_ROD2, "TopExThdRH", WASH_ROD2, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                    //Add a Vertical Joint to the Rod Z axis
                    JointHelper.CreateGlobalAxesAlignedJoint(WASH_ROD2, "TopExThdRH", Axis.Z, Axis.Z);

                    //Add a Rigid Joint between bottom rod and Clevis Hanger
                    JointHelper.CreateRigidJoint(WASH_ROD2, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Rigid Joint between bottom rod and Turnbuckle
                    JointHelper.CreateRigidJoint(WASH_ROD2, "TopExThdRH", NUT_SPRING, "BotInThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Rigid Joint between bottom rod and Turnbuckle
                    JointHelper.CreateRigidJoint(NUT_SPRING, "TopInThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(WASH_ROD2, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(WASH_ROD2, "TopExThdRH", WASH_NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", WASH_NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT5, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT6, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);
                }

                //Set it up for Rod/NUT Usage
                if (topType == "ROD_NUT")
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        flangeThickness = 0;
                    else
                    {
                        if ((SupportHelper.SupportingObjects.Count != 0))
                            flangeThickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                    }

                    topRodLength = topAssyLength;

                    (componentDictionary[ROD1]).SetPropertyValue(topRodLength, "IJUAHgrOccLength", "Length");

                    //Joint from Pipe Clamp to Pipe
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint(PIPE_CLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    else
                        JointHelper.CreateRevoluteJoint(PIPE_CLAMP, "Route", "-1", "Route", Axis.X, Axis.X);

                    //Need to use a connection to lift the rod above the bottom of the structure
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint(NUT_CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.X, -flangeThickness / 2.0 - 0.05, 0, 0);
                    else
                        JointHelper.CreatePlanarJoint(NUT_CONNECTION, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, 0);

                    //Create the Flexible (Prismatic) Joint between the ports of the top rod
                    JointHelper.CreatePrismaticJoint(NUT_ROD2, "TopExThdRH", NUT_ROD2, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                    //Add a Vertical Joint to the Rod Z axis
                    JointHelper.CreateGlobalAxesAlignedJoint(NUT_ROD2, "TopExThdRH", Axis.Z, Axis.Z);

                    if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                        //Add a planar Joint between top of the rod where it connects to the structure connection
                        JointHelper.CreatePlanarJoint(ROD1, "TopExThdRH", NUT_CONNECTION, "Connection", Plane.XY, Plane.XY, 0);
                    else
                        JointHelper.CreateSphericalJoint(ROD1, "TopExThdRH", NUT_CONNECTION, "Connection");

                    //Add a Revolute Joint between bottom rod and Clevis Hanger
                    JointHelper.CreateRigidJoint(NUT_ROD2, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRevoluteJoint(NUT_ROD2, "TopExThdRH", NUT_SPRING, "BotInThdRH", Axis.Z, Axis.Z);
                    else
                        //Add a revolute Joint between bottom rod and Turnbuckle
                        JointHelper.CreateRigidJoint(NUT_ROD2, "TopExThdRH", NUT_SPRING, "BotInThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    JointHelper.CreateRigidJoint(NUT_SPRING, "TopInThdRH", ROD1, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(NUT_ROD2, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    JointHelper.CreateRigidJoint(NUT_ROD2, "TopExThdRH", NUT_NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT_NUT4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_NUT5, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.03, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_NUT6, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.015, 0, 0);
                }

                //Set it up for Beam Att/Type B Spring Usage
                if (topType == "ROD_BEAM_ATT_B")
                {
                    //Joint from Pipe Clamp to Pipe
                    JointHelper.CreateRigidJoint("-1", routePort, PIPE_CLAMP, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Prismatic Joint between the lug and the Structure
                    JointHelper.CreatePlanarJoint(BEAM_ATT, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, 0);

                    //Add a revolute Joint between the Hole and clevis pin
                    if (Configuration == 1)
                        JointHelper.CreateRevoluteJoint(NUT_SPRING, "Hole", BEAM_ATT, "Pin", Axis.X, Axis.X);
                    else
                        JointHelper.CreateRevoluteJoint(NUT_SPRING, "Hole", BEAM_ATT, "Pin", Axis.Y, Axis.Y);

                    //Add a revolute Joint between bottom rod and Turnbuckle
                    if (Configuration == 1)
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_SPRING, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_SPRING, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                    //Create the Flexible (Prismatic) Joint between the ports of the top rod
                    JointHelper.CreatePrismaticJoint(ROD1, "TopExThdRH", ROD1, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                    //Add a Rigid Joint between bottom rod and Clevis Hanger
                    if (slopedRoute == false)
                        if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                            JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        else
                            JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", ATT_NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                }

                //Set it up for Lug/Type C Spring Usage
                if (topType == "ROD_LUG_C")
                {
                    //Joint from Pipe Clamp to Pipe
                    JointHelper.CreateRigidJoint("-1", routePort, PIPE_CLAMP, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a Prismatic Joint between the lug and the Structure
                    JointHelper.CreatePlanarJoint(LUG, "Structure", "-1", "Structure", Plane.XY, Plane.NegativeXY, 0);

                    //Add a Rigid Joint between bottom rod and Clevis Hanger
                    if (slopedRoute == false)
                        if (!(SupportHelper.PlacementType == PlacementType.PlaceByStruct))
                            JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        else
                            JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", CLAMP_EYE_NUT, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    if (Configuration == 1)
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_SPRING, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", NUT_SPRING, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a rigid Joint between top of the rod and the C Clamp
                    JointHelper.CreateRigidJoint(NUT_SPRING, "Hole", LUG_C_PIN, "TopPin", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add a revolute Joint between the lug hole and clevis pin
                    if (Configuration == 1)
                        JointHelper.CreateRevoluteJoint(LUG_C_PIN, "BotPin", LUG, "Hole", Axis.X, Axis.Y);
                    else
                        JointHelper.CreateRevoluteJoint(LUG_C_PIN, "BotPin", LUG, "Hole", Axis.Y, Axis.NegativeX);

                    //Create the Flexible (Prismatic) Joint between the ports of the top rod
                    JointHelper.CreatePrismaticJoint(ROD1, "TopExThdRH", ROD1, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                    //Add a Rigid Joint between the bottom nut and the rod
                    JointHelper.CreateRigidJoint(ROD1, "BotExThdRH", NUT1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(ROD1, "TopExThdRH", LUG_C_NUT3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -0.0508, 0, 0);
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }


        // -----------------------------------------------------------------------------------
        // Get Max Route Connection Value
        // -----------------------------------------------------------------------------------
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

                    routeConnections.Add(new ConnectionInfo(PIPE_CLAMP, 1));  //partindex, routeindex

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
                    //Create a collection to hold the ALL structure connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    //partindex, structindex
                    if (topType == "ROD_C_CLAMP")
                        structConnections.Add(new ConnectionInfo(C_CLAMP, 1));
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
                    if (topType == "ROD_BEAM_ATT_B")
                        structConnections.Add(new ConnectionInfo(BEAM_ATT, 1));
                    if (topType == "ROD_LUG_C")
                        structConnections.Add(new ConnectionInfo(LUG, 1));

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






