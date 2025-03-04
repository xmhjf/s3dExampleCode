//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_DM1.cs
//   Power_Assy,Ingr.SP3D.Content.Support.Rules.Assy_DM1
//   Author       :  Pradeep
//   Creation Date:  28.Mar.2013
//   Description: 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   28-03-2013     Pradeep    CR-CP-224472-Initial Creation
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;

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
    public class Assy_DM1 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string PipeStrut = "PipeStrut";
        private string config;
        private double X;
        private double Y;
        private double length;
        private string size;
        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {

            get
            {
                Collection<PartInfo> parts = new Collection<PartInfo>();
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    X = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowerAssyDM1", "X")).PropValue;
                    Y = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowerAssyDM1", "Y")).PropValue;
                    length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                    size = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrPowerAssyDM1", "PartSize")).ToString();
                    config = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrPowerAssyDM1", "CONFIG")).ToString();
                    string partNumber = "";
                    CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                    PartClass auxTable = (PartClass)cataloghelper.GetPartClass("Assy_DM1PartSel");
                    ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    foreach (BusinessObject classItem in classItems)
                    {
                        if (classItem.GetPropertyValue("IJUAHgrAssyDM1PartSel", "PartSize").ToString() == size && classItem.GetPropertyValue("IJUAHgrAssyDM1PartSel", "CONFIG").ToString() == config)
                        {
                            partNumber = (string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAssyDM1PartSel", "PartNo")).ToString();
                        }
                    }
                    string tempPartStr = partNumber.Substring(13);
                    parts.Add(new PartInfo(PipeStrut, "Anvil_Fig211_" + tempPartStr));
                    return parts;
                }
                catch(Exception e)
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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                //Load standard bounding box definition
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox BBX;
                int routeconnectionValue = Configuration;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                SupportComponent Pipestruct = componentDictionary[PipeStrut];


                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                }
                else
                {
                    BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                }
                double BBXWidth = BBX.Width;
                double BBXHeight = BBX.Height;
                double length1 = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Direct);
                //'=================================
                //'Set the attributes of the part(s)
                //'=================================
                Pipestruct.SetPropertyValue(X, "IJOAHgrAnvil_FIG211", "X");
                Pipestruct.SetPropertyValue(Y, "IJOAHgrAnvil_FIG211", "Y");
                Pipestruct.SetPropertyValue(length1, "IJUAHgrOccLength", "Length");
                //=============
                //'Create Joints
                //'=============
                //'Create a collection to hold the joints
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    JointHelper.CreateRigidJoint(PipeStrut, "Route", "-1", "BBSR_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -BBXHeight / 2, -BBXWidth / 2, 0);
                }
                else
                {
                    JointHelper.CreateRigidJoint(PipeStrut, "Route", "-1", "BBR_Low", Plane.XY, Plane.XY, Axis.X, Axis.X, -BBXHeight / 2, -BBXWidth / 2, 0);
                }
            }
            catch { }
        }
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                
                try
                {
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(PipeStrut, 1));
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
                //Create a collection to hold ALL the Structure Connection information
                Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                structConnections.Add(new ConnectionInfo("PipeStrut", 1));
                return structConnections;

            }
        }
        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject supportOrComponent)
        {
            string BOMString = "";
            try
            {
                String BOM_DESC = string.Empty;
                BOM_DESC = supportOrComponent.GetPropertyValue("IJOAHgrPowAssyBOMDesc", "BOM_DESC").ToString();

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDia = pipeInfo.OutsideDiameter;
                NominalDiameter pipeNominalDiameter = new NominalDiameter();
                pipeNominalDiameter = pipeInfo.NominalDiameter;
                string strUnit = pipeNominalDiameter.Units;
                if (strUnit == "mm")
                {
                    pipeDia = pipeDia * 1000;
                }
                else if (strUnit == "in")
                {
                    pipeDia = pipeDia * 39.37008;
                }

                if (BOM_DESC == "")
                {
                    BOMString = "DM1 Assembly Size A Sway Strut Assembly (Option 3) for Pipe Size " + pipeDia + "" + strUnit;
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
                CmnException e1 = new CmnException("Error in BOM of Assembly - Assy_DM1" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}
