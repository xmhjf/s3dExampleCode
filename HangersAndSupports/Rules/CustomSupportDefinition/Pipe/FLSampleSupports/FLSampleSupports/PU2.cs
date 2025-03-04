//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PU2.cs
//   FLSampleSupports,Ingr.SP3D.Content.Support.Rules.PU2
//   Author       : Rajeswari
//   Creation Date: 29-Aug-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 29-Aug-2013  Rajeswari CR-CP-224478 Convert FlSample_Supports to C# .Net 
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
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class PU2 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string PLATE1 = "PLATE1";
        private const string PLATE2 = "PLATE2";
        private const string HANGERBEAM = "HANGERBEAM";

        string sectionSize;
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

                    string plate = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Plate", "Plate")).PropValue;
                    sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrFl_SecSize", "SecSize")).PropValue;
                    string hangerBeam = (string)((PropertyValueString)support.GetPropertyValue("IJUAHangerBeam", "HangerBeam")).PropValue;

                    parts.Add(new PartInfo(HANGERBEAM, hangerBeam + sectionSize));
                    parts.Add(new PartInfo(PLATE1, plate));
                    parts.Add(new PartInfo(PLATE2, plate));
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
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Can not place-by-structure please place-by-point.", "", "PU2.cs", 84);
                    return;
                }

                double hangerOffsetLeft = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHangerOffset", "Left")).PropValue;
                double hangerOffsetRight = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrHangerOffset", "Right")).PropValue;
                double plateT = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateT")).PropValue;
                double plateW = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateW")).PropValue;
                double plateH = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateH")).PropValue;
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
                double boundingBoxWidth = boundingBox.Width;

                if (boundingBoxWidth > 1.2)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Length can not exceed 1.2m.", "", "PU2.cs", 119);
                    return;
                }

                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                if (routeInfo.InsulationThickness > 0)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARININGMESSAGE", "ConfigureSupport" + ": " + "WARINING: " + "Not placable for Insulated pipe.", "", "PU2.cs", 125);

                string sectionStandard = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrFl_SecStandard", "SecStandard")).PropValue;
                string sectionType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrFl_SecType", "SecType")).PropValue;

                CatalogStructHelper catalogHelper = new CatalogStructHelper();
                CrossSection steelData = catalogHelper.GetCrossSection(sectionStandard, sectionType, sectionSize);
                Double hangerWidth = steelData.Width;

                componentDictionary[HANGERBEAM].SetPropertyValue(boundingBoxWidth + hangerOffsetLeft + hangerOffsetRight, "IJUAHgrOccLength", "Length");
                componentDictionary[PLATE1].SetPropertyValue(plateW, "IJOAHgrUtility_VARIABLE_BOX", "WIDTH");
                componentDictionary[PLATE1].SetPropertyValue(plateH, "IJOAHgrUtility_VARIABLE_BOX", "DEPTH");
                componentDictionary[PLATE1].SetPropertyValue(plateT, "IJUAHgrOccLength", "Length");
                componentDictionary[PLATE1].SetPropertyValue("50x50x6 THK Carbon Steel Plate.", "IJOAHgrUtility_VARIABLE_BOX", "BOM_DESC");
                componentDictionary[PLATE2].SetPropertyValue(plateW, "IJOAHgrUtility_VARIABLE_BOX", "WIDTH");
                componentDictionary[PLATE2].SetPropertyValue(plateH, "IJOAHgrUtility_VARIABLE_BOX", "DEPTH");
                componentDictionary[PLATE2].SetPropertyValue(plateT, "IJUAHgrOccLength", "Length");
                componentDictionary[PLATE2].SetPropertyValue("50x50x6 THK Carbon Steel Plate.", "IJOAHgrUtility_VARIABLE_BOX", "BOM_DESC");

                if (Configuration == 1)
                    JointHelper.CreateRigidJoint("-1", "BBRV_Low", HANGERBEAM, "EndCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, -shoeHeight, -hangerOffsetRight, 0);
                else if (Configuration == 2)
                    JointHelper.CreateRigidJoint("-1", "BBRV_Low", HANGERBEAM, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -shoeHeight, -hangerOffsetRight, 0);

                JointHelper.CreateRigidJoint(HANGERBEAM, "BeginCap", PLATE1, "EndOther", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, hangerWidth / 2);
                JointHelper.CreateRigidJoint(HANGERBEAM, "EndCap", PLATE2, "EndOther", Plane.XY, Plane.XY, Axis.X, Axis.X, plateT, 0, hangerWidth / 2);
                    
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
            return "PU2 - Pick-Up from Single non-Insulated Line.";
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
                if (eMirrorPlane == MirrorPlane.YZPlane)
                {
                    if (CurrentMirrorToggleValue == 1)
                        return 2;
                    else if (CurrentMirrorToggleValue == 2)
                        return 1;
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
