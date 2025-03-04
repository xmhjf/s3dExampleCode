//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SingleCnHg.cs
//   CabletrayAssy,Ingr.SP3D.Content.Support.Rules.SingleCnHg
//   Author       :  MK
//   Creation Date:  04/07/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  04/07/2013       MK        CR-CP-224477 - Converted CabletrayAssemblies to C# .Net
//  30/04/2015       Chethan   TR-CP-271643  Rec exc. minidumps created while placing support by ref. on cabletray  
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
        private const string G4G1461011 = "G4G1461011"; //HngSupHgrAssemblySymbols.ThreadTop  (Beam Clamp)
        private const string G4G1460011 = "G4G1460011"; //HngSupHgrAssemblySymbols.ThreadFlex (Flexible Rod)
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
                    PartClass G4G146101 = (PartClass)catalogBaseHelper.GetPartClass("G4G_1460_01");
                    Part g4gPart2 = null;
                    string g4gPart2Part = string.Empty;
                    foreach (BusinessObject businessObject in G4G146101.Parts)
                    {
                        if (((double)((PropertyValueDouble)businessObject.GetPropertyValue("IJUAHgrDia", "RodSizeDia")).PropValue) >= (((double)((PropertyValueDouble)ctHoldSideClampPart.GetPropertyValue("IJUAHgrCTSgCnHg", "RodDiameter")).PropValue)))
                        {
                            g4gPart2 = (Part)businessObject;
                            g4gPart2Part = g4gPart2.ToString();
                            break;
                        }
                    }
                    parts.Add(new PartInfo(CTSINGLECNHG, ctHoldSideClampPart.ToString()));
                    parts.Add(new PartInfo(G4G1461011, g4gPart1.ToString()));
                    parts.Add(new PartInfo(G4G1460011, g4gPart2Part));

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
                double beamWidth;
                if (width <= 0 || depth <= 0)
                {
                    width = radius * 2;
                    depth = radius * 2;
                }
   
                    if(SupportHelper.SupportingObjects.Count>0)
                    {
                        beamWidth = SupportingHelper.SupportingObjectInfo(1).Width;    
                    }
                    else
                    {
                        beamWidth = 0;
                    }                                                  

                // If unable to retrieve BeamWidth, Structure could be a slab.
                // Use the cable tray width to set the clamp dimensions.

                if (beamWidth <= 0)
                    beamWidth = width;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                (componentDictionary[CTSINGLECNHG]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTSINGLECNHG]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[G4G1461011]).SetPropertyValue((0.75 * beamWidth), "IJUAHgrOccGeometry", "Width");

                //======================================================
                //Get the Current Location in the Route Connection Cycle
                //======================================================


                //======================================================
                //Create Joints
                //======================================================

                if (Configuration == 1 || Configuration == 4)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint(CTSINGLECNHG, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(CTSINGLECNHG, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        
                }
                else if (Configuration == 3 || Configuration == 2)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint(CTSINGLECNHG, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(CTSINGLECNHG, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);
                }

                //Add a Spherical Joint between cable tray Clamp and Bottom of Rod
                JointHelper.CreateSphericalJoint(CTSINGLECNHG, "Structure", G4G1460011, "RodBottom");

                //Add a Spherical Joint between Beam Clamp and Top of Rod
                JointHelper.CreateSphericalJoint(G4G1460011, "RodTop", G4G1461011, "Rod");

                //Create the Flexible (Cylindrical) Joint between the ports of the rod
                JointHelper.CreateCylindricalJoint(G4G1460011, "RodTop", G4G1460011, "RodBottom", Axis.Z, Axis.Z,0);

                if (Configuration == 1 || Configuration == 2)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        //Beam Structure ... Add a Prismatic Joint between Beam and Beam Clamp
                        JointHelper.CreatePrismaticJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    else
                        //Plate Structure ... Add a Translation Joint Between Plate and Beam Clamp
                        JointHelper.CreateRigidJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                else if (Configuration == 3 || Configuration == 4)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        //Beam Structure ... Add a Prismatic Joint between Beam and Beam Clamp
                        JointHelper.CreatePrismaticJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0);
                    else
                        //Plate Structure ... Add a Translation Joint Between Plate and Beam Clamp
                        JointHelper.CreateRigidJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0, 0);
                }
                //Add a Vertical Joint to the Rod Z axis
                JointHelper.CreateGlobalAxesAlignedJoint(G4G1460011, "RodBottom", Axis.Z, Axis.Z);

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

                    structConnections.Add(new ConnectionInfo(G4G1461011, 1)); // partindex, routeindex


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