//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SingleCnHg.cs
//   CabletrayAssy2,Ingr.SP3D.Content.Support.Rules.ClipHoldClampSingleCnHg
//   Author       : Vinay
//   Creation Date:  20/11/2015
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Support.Middle.Hidden;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class SingleCnHg : CustomSupportDefinition
    {
        private const string BEAM_ATT = "WELDEDBEAMATTACHMENT";
        private const string EYE_NUT = "EYENUT1";
        private const string ROD = "ROD"; 
        private const string CTSINGLECNHG = "CTSINGLECNHG";
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    SupportComponentUtils supportComponentUtils = new SupportComponentUtils();
                    PartClass CTClipHoldClamp = (PartClass)catalogBaseHelper.GetPartClass("CTSingleCnHg");
                    string partselection = CTClipHoldClamp.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
                    Part ctHoldSideClampPart = supportComponentUtils.GetPartFromPartClass("CTSingleCnHg", partselection, support);
                    Part g4gPart1 = supportComponentUtils.GetPartFromPartClass("G4G_1461_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize", support);

                    PropertyValueCodelist rodDia = ((PropertyValueCodelist)support.GetPropertyValue("IJUAHSA_RodDia", "RodDiameter"));
                    string rodDiametr1 = rodDia.PropertyInfo.CodeListInfo.GetCodelistItem(rodDia.PropValue).DisplayName;

                    // WBAAttachment 
                    parts.Add(new PartInfo(BEAM_ATT, "S3Dhs_WBABolt-" + rodDiametr1));

                    //Flexible Rod
                    parts.Add(new PartInfo(ROD, "S3Dhs_RodCT-" + rodDiametr1));

                    //TopEyenut
                    parts.Add(new PartInfo(EYE_NUT, "S3Dhs_EyeNut-" + rodDiametr1));

                    //Channel
                    parts.Add(new PartInfo(CTSINGLECNHG, ctHoldSideClampPart.ToString()));


                    return parts;
                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
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
                //=== ========== == ==== ===========
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========    
                CableTrayObjectInfo cableInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                double depth = cableInfo.Depth;
                double width = cableInfo.Width;
                double radius = cableInfo.BendRadius;
                if (width <= 0 || depth <= 0)
                {
                    width = radius * 2;
                    depth = radius * 2;
                }
   
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                (componentDictionary[CTSINGLECNHG]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTSINGLECNHG]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");

                //======================================================
                //Create Joints
                //======================================================
                // Add a Planar Joint between the Beam Attachment and the structure
                JointHelper.CreatePlanarJoint(BEAM_ATT, "Structure", "-1", "Structure", Plane.XY, Plane.XY, 0);

                // Add a revolute Joint between the Eye nut and Beam Attachment
                JointHelper.CreateRevoluteJoint(EYE_NUT, "Eye", BEAM_ATT, "Pin", Axis.Y, Axis.Y);

                //Add a Rigid Joint between Rod and Eyenut
                if (Configuration == 1 || Configuration == 2)
                    JointHelper.CreateRigidJoint(ROD, "RodEnd2", EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                else
                    JointHelper.CreateRigidJoint(ROD, "RodEnd2", EYE_NUT, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                // Add a Vertical Joint to the Rod Z axis
                JointHelper.CreateGlobalAxesAlignedJoint(ROD, "RodEnd1", Axis.Z, Axis.Z);

                // Add the Flexible (Prismatic) Joint between the ports of the bottom rod
                JointHelper.CreatePrismaticJoint(ROD, "RodEnd1", ROD, "RodEnd2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                //Add a Rigid Joint between cable tray Clamp and Bottom of Rod
                JointHelper.CreateRigidJoint(ROD, "RodEnd1", CTSINGLECNHG, "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                //Add a Rigid Joint between cable tray Clamp and Route
                if (Configuration == 1 || Configuration == 4)
                {
                    JointHelper.CreateRigidJoint(CTSINGLECNHG, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                else if (Configuration == 3 || Configuration == 2)
                {
                    JointHelper.CreateRigidJoint(CTSINGLECNHG, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);
                }
            }
            catch (Exception exception)
            {
                Type myType = this.GetType();
                CmnException exception1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                throw exception1;
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

                    routeConnections.Add(new ConnectionInfo(CTSINGLECNHG, 1)); // partindex, routeindex
                    routeConnections.Add(new ConnectionInfo(CTSINGLECNHG, 1)); // partindex, routeindex

                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
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

                    structConnections.Add(new ConnectionInfo(BEAM_ATT, 1)); // partindex, structureindex


                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
                }
            }
        }

    }
}