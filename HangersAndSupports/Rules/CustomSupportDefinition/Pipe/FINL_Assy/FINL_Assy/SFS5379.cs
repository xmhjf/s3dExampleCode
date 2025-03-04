//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5379.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5379
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
    [SymbolVersion("1.0.0.0")]
    public class SFS5379 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        double normalPipeDiameterMetric;

        private const string PLATE1 = "Plate1_5379";
        private const string PLATE2 = "Plate2_5379";
        private const string ROUTECONNOBJECT = "RouteConnObject_5379";
        private const string LEFTLEG = "LeftLeg_5379";
        private const string RIGHTLEG = "RightLeg_5379";
        public int Index { get; set; }
        public Boolean Override { get; set; }
        public int Clamps { get; set; }
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    //To get Pipe Nom Dia
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                    if (pipeInfo.NominalDiameter.Units != "mm")
                        normalPipeDiameterMetric = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, UnitName.NPD_INCH, UnitName.NPD_MILLIMETER) / 1000;
                    else
                        normalPipeDiameterMetric = pipeInfo.NominalDiameter.Size / 1000;

                    NominalDiameter minNominalDiameter = new NominalDiameter();
                    minNominalDiameter.Size = 20;
                    minNominalDiameter.Units = "mm";
                    NominalDiameter maxNominalDiameter = new NominalDiameter();
                    maxNominalDiameter.Size = 500;
                    maxNominalDiameter.Units = "mm";

                    //check valid pipe size
                    if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, null) == false)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5379.cs", 74);
                        return parts;
                    }

                    if (normalPipeDiameterMetric < 200.0 / 1000)       //T shape
                    {
                        parts.Add(new PartInfo(PLATE1 + "_" + Index, "Utility_USER_FIXED_BOX_1"));
                        parts.Add(new PartInfo(PLATE2 + "_" + Index, "Utility_USER_FIXED_BOX_1"));
                        parts.Add(new PartInfo(ROUTECONNOBJECT + "_" + Index, "Log_Conn_Part_1"));           //connect obj on route so we can rotate a whole assy
                    }
                    else
                    {
                        parts.Add(new PartInfo(PLATE1 + "_" + Index, "Utility_USER_FIXED_BOX_1"));        //Second Clamp
                        parts.Add(new PartInfo(PLATE2 + "_" + Index, "Utility_USER_FIXED_BOX_1"));        //Second Clamp
                        parts.Add(new PartInfo(ROUTECONNOBJECT + "_" + Index, "Log_Conn_Part_1"));           //connect obj on route so we can rotate a whole assy
                        parts.Add(new PartInfo(LEFTLEG + "_" + Index, "Utility_GENERIC_L_1"));            //First Clamp
                        parts.Add(new PartInfo(RIGHTLEG + "_" + Index, "Utility_GENERIC_L_1"));           //Second Clamp
                    }

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
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                double pipeOR = 0, plateThickness = 0, plateDepth = 0;

                //Getting Dimension Information
                double shoeH = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5379_Dim", "IJUAFINLSrv_SFS5379_Dim", "Shoe_H", "IJUAFINLSrv_SFS5379_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double length = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5379_Dim", "IJUAFINLSrv_SFS5379_Dim", "L", "IJUAFINLSrv_SFS5379_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double E = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5379_Dim", "IJUAFINLSrv_SFS5379_Dim", "E", "IJUAFINLSrv_SFS5379_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double A = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5379_Dim", "IJUAFINLSrv_SFS5379_Dim", "A", "IJUAFINLSrv_SFS5379_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double B = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5379_Dim", "IJUAFINLSrv_SFS5379_Dim", "B", "IJUAFINLSrv_SFS5379_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double S = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5379_Dim", "IJUAFINLSrv_SFS5379_Dim", "S", "IJUAFINLSrv_SFS5379_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double H = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5379_Dim", "IJUAFINLSrv_SFS5379_Dim", "H", "IJUAFINLSrv_SFS5379_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));

                pipeOR = H - shoeH;

                //Setting attribute of two plates:  L, WIDTH, DEPTH, BOM_DESC
                string plate1Bom, plate2Bom;

                if (normalPipeDiameterMetric < 200.0 / 1000)     //T Shape
                {
                    plateThickness = 8.0 / 1000;
                    plateDepth = 60.0 / 1000;

                    plate1Bom = "Plate SFS 2022 " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, shoeH - plateThickness, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateThickness, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                    componentDictionary[PLATE1 + "_" + Index].SetPropertyValue(shoeH - plateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                    componentDictionary[PLATE1 + "_" + Index].SetPropertyValue(plateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                    componentDictionary[PLATE1 + "_" + Index].SetPropertyValue(length, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                    componentDictionary[PLATE1 + "_" + Index].SetPropertyValue(plate1Bom, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");

                    plate2Bom = "Plate SFS 2022 " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, shoeH - plateThickness, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateThickness, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                    componentDictionary[PLATE2 + "_" + Index].SetPropertyValue(plateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                    componentDictionary[PLATE2 + "_" + Index].SetPropertyValue(plateDepth, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                    componentDictionary[PLATE2 + "_" + Index].SetPropertyValue(length, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                    componentDictionary[PLATE2 + "_" + Index].SetPropertyValue(plate2Bom, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");

                    //this will be overriden by super assembly
                    if (!Override)     //by default it is false
                        JointHelper.CreateRigidJoint("-1", "Route", ROUTECONNOBJECT + "_" + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    //Add the joints between route and vertical plate
                    JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", PLATE1 + "_" + Index, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, pipeOR, 0, 0);

                    //Add the joints between vertical plate and horizontal plate
                    JointHelper.CreateRigidJoint(PLATE1 + "_" + Index, "EndOther", PLATE2 + "_" + Index, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                else
                {
                    length = 350.00 / 1000.00;
                    plateThickness = 8.00 / 1000.00;
                    plateDepth = 80.00 / 1000.00;

                    plate1Bom = "Plate SFS 2022 " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateDepth, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, E, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateThickness, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                    componentDictionary[PLATE1 + "_" + Index].SetPropertyValue(plateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                    componentDictionary[PLATE1 + "_" + Index].SetPropertyValue(E, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                    componentDictionary[PLATE1 + "_" + Index].SetPropertyValue(plateDepth, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                    componentDictionary[PLATE1 + "_" + Index].SetPropertyValue(plate1Bom, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");

                    plate2Bom = "Plate SFS 2022 " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateDepth, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, E, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateThickness, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                    componentDictionary[PLATE2 + "_" + Index].SetPropertyValue(plateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                    componentDictionary[PLATE2 + "_" + Index].SetPropertyValue(E, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                    componentDictionary[PLATE2 + "_" + Index].SetPropertyValue(plateDepth, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                    componentDictionary[PLATE2 + "_" + Index].SetPropertyValue(plate2Bom, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");

                    //Setting attribute of two plates L sections
                    string leftLegBom = "L" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, B, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, A, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, S, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                    componentDictionary[LEFTLEG + "_" + Index].SetPropertyValue(length, "IJOAHgrUtility_GENERIC_L", "L");
                    componentDictionary[LEFTLEG + "_" + Index].SetPropertyValue(B, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                    componentDictionary[LEFTLEG + "_" + Index].SetPropertyValue(A, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                    componentDictionary[LEFTLEG + "_" + Index].SetPropertyValue(S, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                    componentDictionary[LEFTLEG + "_" + Index].SetPropertyValue(leftLegBom, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");

                    string rightLegBom = "L" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, B, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, A, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, S, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                    componentDictionary[RIGHTLEG + "_" + Index].SetPropertyValue(length, "IJOAHgrUtility_GENERIC_L", "L");
                    componentDictionary[RIGHTLEG + "_" + Index].SetPropertyValue(B, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                    componentDictionary[RIGHTLEG + "_" + Index].SetPropertyValue(A, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                    componentDictionary[RIGHTLEG + "_" + Index].SetPropertyValue(S, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                    componentDictionary[RIGHTLEG + "_" + Index].SetPropertyValue(rightLegBom, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");

                    //this will be overriden by super assembly
                    if (Override == false)     //by default it is false
                        JointHelper.CreateRigidJoint("-1", "Route", ROUTECONNOBJECT + "_" + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    double offset = Math.Sqrt(pipeOR * pipeOR - (E / 2) * (E / 2));

                    //Add the joints between route and vertical plate
                    JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", PLATE1 + "_" + Index, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, offset + A - plateThickness - 10.0 / 1000, 0, length / 2 - plateDepth / 2);

                    //Add the joints between vertical plate and horizontal plate
                    JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", PLATE2 + "_" + Index, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, offset + A - plateThickness - 10.0 / 1000, 0, -length / 2 + plateDepth / 2);

                    //Add the joints between route and vertical plate
                    JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", LEFTLEG + "_" + Index, "Structure", Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeY, -length / 2, offset + A, E / 2);

                    //Add the joints between vertical plate and horizontal plate
                    JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", RIGHTLEG + "_" + Index, "Structure", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.Y, offset + A, length / 2, -E / 2);
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of Finl_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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

                    routeConnections.Add(new ConnectionInfo(ROUTECONNOBJECT + "_" + Index, 1));      //partindex, routeindex

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
                    if (normalPipeDiameterMetric < 200.0 / 1000)
                        structConnections.Add(new ConnectionInfo(PLATE1 + "_" + Index, 1));      //partindex, structindex
                    else
                    {
                        structConnections.Add(new ConnectionInfo(LEFTLEG + "_" + Index, 1));      //partindex, structindex
                        structConnections.Add(new ConnectionInfo(RIGHTLEG + "_" + Index, 1));     //partindex, structindex
                    }

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
                //To get Pipe Nom Dia
                double pipeDiameter = ((PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1)).NominalDiameter.Size;

                BOMString = "Pipe slide SFS 5379 DN " + Convert.ToString(pipeDiameter);

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
