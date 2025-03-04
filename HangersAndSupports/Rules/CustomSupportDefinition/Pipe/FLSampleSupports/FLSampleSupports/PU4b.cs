//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PU4b.cs
//   FLSampleSupports,Ingr.SP3D.Content.Support.Rules.PU4b
//   Author       : Rajeswari
//   Creation Date: 30-Aug-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 30-Aug-2013  Rajeswari CR-CP-224478 Convert FlSample_Supports to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;

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

    public class PU4b : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string HANGERBEAM = "HANGERBEAM";
        double boundingBoxWidth = 0, boundingBoxHeight = 0;
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

                    string sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrFl_SecSize", "SecSize")).PropValue;
                    string hangerBeam = (string)((PropertyValueString)support.GetPropertyValue("IJUAHangerBeam", "HangerBeam")).PropValue;

                    parts.Add(new PartInfo(HANGERBEAM, hangerBeam + sectionSize));
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

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)     //It is placed by point only
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Can not place-by-structure please place-by-point.", "", "PU4b.cs", 76);
                    return;
                }

                double hangerOffsetLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHangerOffset", "Left")).PropValue;
                double hangerOffsetRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHangerOffset", "Right")).PropValue;
                double shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrShoeH", "ShoeHeight")).PropValue;

                // ==========================
                // 1. Load standard bounding box definition
                // ==========================
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;

                // ==========================
                // 3. retrieve dimension of the bounding box
                // ==========================
                // Get route box geometry
                //  ____________________
                // |                    |
                // |  ROUTE BOX BOUND   | dHeight
                // |____________________|
                //        dWidth

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                boundingBoxWidth = boundingBox.Width;
                boundingBoxHeight = boundingBox.Height;

                if (boundingBoxWidth + hangerOffsetLeft + hangerOffsetRight > 1.8)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Length can not exceed 1.8m.", "", "PU4b.cs", 110);
                    return;
                }

                if (Math.Round(boundingBoxWidth, 3) > Math.Round(boundingBoxHeight, 3))
                {
                    componentDictionary[HANGERBEAM].SetPropertyValue(boundingBoxWidth + hangerOffsetLeft + hangerOffsetRight, "IJUAHgrOccLength", "Length");
                    if (Configuration == 1 || Configuration == 5)
                        JointHelper.CreateRigidJoint("-1", "BBR_High", HANGERBEAM, "EndCap", Plane.XY, Plane.ZX, Axis.X, Axis.X, shoeHeight, -boundingBoxWidth - hangerOffsetRight, -0.045);
                    else if (Configuration == 2 || Configuration == 6)
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HANGERBEAM, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -shoeHeight, -hangerOffsetLeft, -0.045);
                    else if (Configuration == 3 || Configuration == 7)
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HANGERBEAM, "EndCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -shoeHeight, -hangerOffsetLeft, 0.045);
                    else if (Configuration == 4 || Configuration == 8)
                        JointHelper.CreateRigidJoint("-1", "BBR_High", HANGERBEAM, "BeginCap", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX, shoeHeight, -boundingBoxWidth - hangerOffsetRight, 0.045);
                }
                else if (Math.Round(boundingBoxWidth, 3) < Math.Round(boundingBoxHeight, 3))
                {
                    componentDictionary[HANGERBEAM].SetPropertyValue(boundingBoxHeight + hangerOffsetLeft + hangerOffsetRight, "IJUAHgrOccLength", "Length");
                    if (Configuration == 1 || Configuration == 5)
                        JointHelper.CreateRigidJoint("-1", "BBR_High", HANGERBEAM, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, -hangerOffsetRight, shoeHeight, -0.045);
                    else if (Configuration == 2 || Configuration == 6)
                        JointHelper.CreateRigidJoint("-1", "BBR_High", HANGERBEAM, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, boundingBoxHeight + hangerOffsetRight, shoeHeight, 0.045);
                    else if (Configuration == 3 || Configuration == 7)
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HANGERBEAM, "BeginCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, boundingBoxHeight + hangerOffsetRight, shoeHeight, -0.045);
                    else if (Configuration == 4 || Configuration == 8)
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HANGERBEAM, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, -hangerOffsetRight, -shoeHeight, 0.045);
                }
                else
                {
                    // if there is only 1 pipe then you need to toggle eight times
                    componentDictionary[HANGERBEAM].SetPropertyValue(boundingBoxWidth + hangerOffsetLeft + hangerOffsetRight, "IJUAHgrOccLength", "Length");

                    if (Configuration == 1)
                        JointHelper.CreateRigidJoint("-1", "BBR_High", HANGERBEAM, "EndCap", Plane.XY, Plane.ZX, Axis.X, Axis.X, shoeHeight, -boundingBoxWidth - hangerOffsetRight, -0.045);
                    else if (Configuration == 2)
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HANGERBEAM, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -shoeHeight, -hangerOffsetLeft, -0.045);
                    else if (Configuration == 3)
                        JointHelper.CreateRigidJoint("-1", "BBR_High", HANGERBEAM, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, hangerOffsetRight, shoeHeight, -0.045);
                    else if (Configuration == 4)
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HANGERBEAM, "BeginCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, boundingBoxWidth + hangerOffsetRight, -shoeHeight, -0.045);
                    else if (Configuration == 5)
                        JointHelper.CreateRigidJoint("-1", "BBR_High", HANGERBEAM, "EndCap", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX, shoeHeight, hangerOffsetRight, 0.045);
                    else if (Configuration == 6)
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HANGERBEAM, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -shoeHeight, boundingBoxWidth + hangerOffsetLeft, 0.045);
                    else if (Configuration == 7)
                        JointHelper.CreateRigidJoint("-1", "BBR_High", HANGERBEAM, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, -boundingBoxWidth - hangerOffsetRight, shoeHeight, 0.045);
                    else if (Configuration == 8)
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HANGERBEAM, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, -hangerOffsetRight, -shoeHeight, 0.045);
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
                return 8;
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
                    routeConnections.Add(new ConnectionInfo(HANGERBEAM, 1));

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
                    structConnections.Add(new ConnectionInfo(HANGERBEAM, 1));

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
        #region "ICustomHgrBOMDescription Members"
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            return "PU4b - Pick-Up (Stabilizer for Vertical Lines) Pipe Sizes 6in and Smaller.";
        }
        #endregion
        /// <summary>
        /// This method returns the Mirrored configuration Value
        /// </summary>
        /// <param name="CurrentMirrorToggleValue">int - Toggle Value.</param>
        /// <param name="eMirrorPlane">MirrorPlane - eMirrorPlane.</param>
        /// <returns>int</returns>        
        /// <code>
        ///     MirroredConfiguration(int CurrentMirrorToggleValue, MirrorPlane eMirrorPlane);
        ///</code>
        public override int MirroredConfiguration(int CurrentMirrorToggleValue, MirrorPlane eMirrorPlane)
        {
            try
            {
                if (Math.Round(boundingBoxWidth, 3) > Math.Round(boundingBoxHeight, 3))
                {
                    if (eMirrorPlane == MirrorPlane.XYPlane)
                    {
                        if (CurrentMirrorToggleValue == 1 || CurrentMirrorToggleValue == 5)
                            return 2;
                        else if (CurrentMirrorToggleValue == 2 || CurrentMirrorToggleValue == 6)
                            return 1;
                        else if (CurrentMirrorToggleValue == 3 || CurrentMirrorToggleValue == 7)
                            return 4;
                        else if (CurrentMirrorToggleValue == 4 || CurrentMirrorToggleValue == 8)
                            return 3;
                    }
                    else if (eMirrorPlane == MirrorPlane.YZPlane)
                    {
                        if (CurrentMirrorToggleValue == 1 || CurrentMirrorToggleValue == 5)
                            return 4;
                        else if (CurrentMirrorToggleValue == 2 || CurrentMirrorToggleValue == 6)
                            return 3;
                        else if (CurrentMirrorToggleValue == 3 || CurrentMirrorToggleValue == 7)
                            return 2;
                        else if (CurrentMirrorToggleValue == 4 || CurrentMirrorToggleValue == 8)
                            return 1;
                    }
                }
                else if (Math.Round(boundingBoxWidth, 3) < Math.Round(boundingBoxHeight, 3))
                {
                    if (eMirrorPlane == MirrorPlane.XZPlane)
                    {
                        if (CurrentMirrorToggleValue == 1 || CurrentMirrorToggleValue == 5)
                            return 3;
                        else if (CurrentMirrorToggleValue == 2 || CurrentMirrorToggleValue == 6)
                            return 4;
                        else if (CurrentMirrorToggleValue == 3 || CurrentMirrorToggleValue == 7)
                            return 1;
                        else if (CurrentMirrorToggleValue == 4 || CurrentMirrorToggleValue == 8)
                            return 2;
                    }
                    else if (eMirrorPlane == MirrorPlane.YZPlane)
                    {
                        if (CurrentMirrorToggleValue == 1 || CurrentMirrorToggleValue == 5)
                            return 2;
                        else if (CurrentMirrorToggleValue == 2 || CurrentMirrorToggleValue == 6)
                            return 1;
                        else if (CurrentMirrorToggleValue == 3 || CurrentMirrorToggleValue == 7)
                            return 4;
                        else if (CurrentMirrorToggleValue == 4 || CurrentMirrorToggleValue == 8)
                            return 3;
                    }
                }
                else
                {
                    if (eMirrorPlane == MirrorPlane.YZPlane)
                    {
                        if (CurrentMirrorToggleValue == 1)
                            return 5;
                        else if (CurrentMirrorToggleValue == 2)
                            return 6;
                        else if (CurrentMirrorToggleValue == 3)
                            return 7;
                        else if (CurrentMirrorToggleValue == 4)
                            return 8;
                        else if (CurrentMirrorToggleValue == 5)
                            return 1;
                        else if (CurrentMirrorToggleValue == 6)
                            return 2;
                        else if (CurrentMirrorToggleValue == 7)
                            return 3;
                        else if (CurrentMirrorToggleValue == 8)
                            return 4;
                    }
                    else if (eMirrorPlane == MirrorPlane.XYPlane)
                    {
                        if (CurrentMirrorToggleValue == 1)
                            return 2;
                        else if (CurrentMirrorToggleValue == 2)
                            return 1;
                        else if (CurrentMirrorToggleValue == 5)
                            return 6;
                        else if (CurrentMirrorToggleValue == 6)
                            return 5;
                    }
                    else if (eMirrorPlane == MirrorPlane.XZPlane)
                    {
                        if (CurrentMirrorToggleValue == 3)
                            return 4;
                        else if (CurrentMirrorToggleValue == 4)
                            return 3;
                        else if (CurrentMirrorToggleValue == 7)
                            return 8;
                        else if (CurrentMirrorToggleValue == 8)
                            return 7;
                    }
                }
                return CurrentMirrorToggleValue;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Mirrored Configuration." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}
