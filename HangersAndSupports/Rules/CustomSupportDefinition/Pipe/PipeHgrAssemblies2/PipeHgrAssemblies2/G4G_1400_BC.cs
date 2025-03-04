//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   G4G_1400_BC.cs
//   PipeHgrAssemblies2,Ingr.SP3D.Content.Support.Rules.G4G_1400_BC
//   Author       : Vinay
//   Creation Date: 09-Sep-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   27-11-2015      Vinay   DI-CP-276798	Replace the use of any HS_Utility parts
//   07-04-2016      Vinay   TR-CP-292269 	Unable to place HS_G4G_1401_03_H1/H2 under HS_Asssembly_V2 by ‘place by ref’ 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class G4G_1400_BC : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblies2,Ingr.SP3D.Content.Support.Rules.G4G_1400_BC"
        //----------------------------------------------------------------------------------
        //Constants
        private const string BEAM_ATT = "WELDEDBEAMATTACHMENT";
        private const string EYE_NUT = "EYENUT1";
        private const string ROD = "ROD";
        private const string PIPE_LUG = "PIPELUG";
        private const string Clevis = "CLEVIS";

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
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                    PropertyValueCodelist rodDia = ((PropertyValueCodelist)support.GetPropertyValue("IJUAHSA_RodDia", "RodDiameter"));
                    string rodDiametr1 = rodDia.PropertyInfo.CodeListInfo.GetCodelistItem(rodDia.PropValue).DisplayName;

                    // WBAAttachment 
                    parts.Add(new PartInfo(BEAM_ATT, "S3Dhs_WBABolt-" + rodDiametr1));

                    //TopEyenut
                    parts.Add(new PartInfo(EYE_NUT, "S3Dhs_EyeNut-" + rodDiametr1));

                    //Flexible Rod
                    parts.Add(new PartInfo(ROD, "S3Dhs_RodCT-" + rodDiametr1));

                    //Pipe Clamp
                    parts.Add(new PartInfo(PIPE_LUG, "S3dhs_PipeLug - " + rodDiametr1));

                    //Clevis
                    parts.Add(new PartInfo(Clevis, "S3Dhs_ClevisWithpin-" + rodDiametr1));


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
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                //Support the Route Connection Button
                return 2;
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = routeInfo.OutsideDiameter;
                //=== ========== == ==== ===========
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                //====== ======
                //Create Joints
                //====== ======
                // Add a Revolute Joint between Clevis and Pipe Lug
                JointHelper.CreateRevoluteJoint(Clevis, "Pin", PIPE_LUG, "Hole1", Axis.X, Axis.X);

                // Add a Vertical Joint to the Rod Z axis

                JointHelper.CreateGlobalAxesAlignedJoint(ROD, "RodEnd1", Axis.Z, Axis.Z);

                // Create the Flexible (Prismatic) Joint between the ports of the bottom rod

                JointHelper.CreatePrismaticJoint(ROD, "RodEnd1", ROD, "RodEnd2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                // Joint from Pipe Lug to Pipe
                JointHelper.CreateRigidJoint("-1", "Route", PIPE_LUG, "Hole2", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, pipeDiameter / 2, 0, 0);

                // Add a Planar Joint between the lug and the Structure

                JointHelper.CreatePlanarJoint(BEAM_ATT, "Structure", "-1", "Structure", Plane.XY, Plane.XY, 0);

                // Add a revolute Joint between the Eye nut and Beam Attachment

                JointHelper.CreateRevoluteJoint(EYE_NUT, "Eye", BEAM_ATT, "Pin", Axis.Y, Axis.Y);

                // Add a Rigid Joint between top of the rod and the eye nut
                if (Configuration == 2)
                    JointHelper.CreateRigidJoint(ROD, "RodEnd2", EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                else
                    JointHelper.CreateRigidJoint(ROD, "RodEnd2", EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);


                JointHelper.CreateRigidJoint(ROD, "RodEnd1", Clevis, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
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

                    //Add the PipeClampConnections to the Route Collection of Connections.
                    routeConnections.Add(new ConnectionInfo(PIPE_LUG, 1));

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
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    //Add the PipeClampConnections to the Route Collection of Connections.
                    //"G4G_1461_01" (i.e. the beam clamp).  It is the first part
                    //G4G_1451_06 Connects to First Structure Input (i.e. the beam or plate)
                    structConnections.Add(new ConnectionInfo(BEAM_ATT, 1));

                    //Return the collection of Route connection information.
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

