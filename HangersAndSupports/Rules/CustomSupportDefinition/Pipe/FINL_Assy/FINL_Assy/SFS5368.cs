//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5368.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5368
//   Author       :  Vijay
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     Vijay   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
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
    [SymbolVersion("1.0.0.0")]
    public class SFS5368 : CustomSupportDefinition, ICustomHgrBOMDescription
    {

        private const string STOPPER = "Stopper_5368";
        private const string ROUTECONNOBJECT = "RouteConnObject_5368";

        public Boolean Override { get; set; }
        public int StopperType { get; set; }
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    if (!Override)
                        StopperType = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLStopperType", "Type")).PropValue;

                    parts.Add(new PartInfo(STOPPER, "FINLCmp_SFS5368"));
                    parts.Add(new PartInfo(ROUTECONNOBJECT, "Log_Conn_Part_1"));       //connect obj on route so we can rotate a whole assy

                    return parts;       //Get the collection of parts
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
                return 2;
            }
        }

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                PropertyValueCodelist stopperTypeCodelist = (PropertyValueCodelist)(componentDictionary[STOPPER]).GetPropertyValue("IJOAFINL_StopperType", "StopperType");
                if (stopperTypeCodelist.PropValue == -1)
                    stopperTypeCodelist.PropValue = 1;

                if (StopperType == 1)
                    stopperTypeCodelist.PropValue = 1;
                else
                    stopperTypeCodelist.PropValue = 2;

                componentDictionary[STOPPER].SetPropertyValue(stopperTypeCodelist.PropValue, "IJOAFINL_StopperType", "StopperType");


                //this will be overriden by super assembly
                if (!Override)     //by default it is false
                    JointHelper.CreateRigidJoint("-1", "Route", ROUTECONNOBJECT, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                JointHelper.CreateRigidJoint(ROUTECONNOBJECT, "Connection", STOPPER, "Route", Plane.YZ, Plane.YZ, Axis.Z, Axis.NegativeZ, 0, 0, 0);
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of SFS5368." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                    if (Override == true)
                        routeConnections.Add(new ConnectionInfo(ROUTECONNOBJECT, 1));      //partindex, routeindex
                    else
                        routeConnections.Add(new ConnectionInfo(STOPPER, 1));       //partindex, routeindex

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
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    structConnections.Add(new ConnectionInfo(STOPPER, 1));      //partindex, structindex

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

        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string BOMString = "";
            try
            {
                string size, material;
                double pipeDiameter = ((PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1)).NominalDiameter.Size;

                if (pipeDiameter < 110)
                    size = "1";
                else if (pipeDiameter > 110 && pipeDiameter < 310)
                    size = "2";
                else
                    size = "3";

                material = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJUAFINLMaterial", "Material")).PropValue;

                BOMString = "Stopper SFS 5368 " + StopperType + size + " " + material;

                return BOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        #endregion


    }
}
