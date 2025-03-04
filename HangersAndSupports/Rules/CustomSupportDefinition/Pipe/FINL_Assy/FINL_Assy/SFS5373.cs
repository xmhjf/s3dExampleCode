//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5373.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5373
//   Author       :  Vijay
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     Vijay   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
//  11-Dec-2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
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
    [SymbolVersion("1.0.0.0")]
    public class SFS5373 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        double rotation, normalPipeDiameterMetric;
        string material;

        private const string VERPLATE = "VerPlate_5373";
        private const string HORPLATE = "HorPlate_5373";
        private const string CONNOBJECT = "ConnObj_5373";
        private const string PIPECLAMP1 = "PipeClamp1_5373";
        private const string PIPECLAMP2 = "PipeClamp2_5373";
        private const string ROUTECONNOBJECT = "RouteConnObj_5373";
        public int Index { get; set; }
        public int Clamps { get; set; }
        public Boolean Override { get; set; }

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    //only if it is not by super assy    
                    if (!Override)
                    {
                        if (support.SupportsInterface("IJUAFINLClamps1"))
                            Clamps = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLClamps1", "Clamps")).PropValue;
                        else
                            Clamps = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLClamps", "Clamps")).PropValue;
                    }
                    if (support.SupportsInterface("IJUAFINLRot1"))
                        rotation = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLRot1", "Rot")).PropValue;
                    else
                        rotation = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLRot", "Rot")).PropValue;
                    material = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLMaterial", "Material")).PropValue;

                    //To get Pipe Nom Dia
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                    if (pipeInfo.NominalDiameter.Units != "mm")
                        normalPipeDiameterMetric = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, UnitName.NPD_INCH, UnitName.NPD_MILLIMETER) / 1000;
                    else
                        normalPipeDiameterMetric = pipeInfo.NominalDiameter.Size / 1000;

                    NominalDiameter minNominalDiameter = new NominalDiameter();
                    minNominalDiameter.Size = 10;
                    minNominalDiameter.Units = "mm";
                    NominalDiameter maxNominalDiameter = new NominalDiameter();
                    maxNominalDiameter.Size = 150;
                    maxNominalDiameter.Units = "mm";
                    NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 90 }, "mm");

                    //check valid pipe size
                    if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5373.cs", 91);

                    if (Clamps == 1)       //with clamps
                    {
                        if (normalPipeDiameterMetric > (49.0 /1000.0))        //with two clamps
                        {
                            parts.Add(new PartInfo(VERPLATE + "_" + Index, "Utility_USER_FIXED_BOX_1"));      //Second Clamp
                            parts.Add(new PartInfo(HORPLATE + "_" + Index, "Utility_USER_FIXED_BOX_1"));      //Second Clamp
                            parts.Add(new PartInfo(CONNOBJECT + "_" + Index, "Log_Conn_Part_1"));               //Rotational Connection Object
                            parts.Add(new PartInfo(ROUTECONNOBJECT + "_" + Index, "Log_Conn_Part_1"));           //connect obj on route so we can rotate a whole assy
                            parts.Add(new PartInfo(PIPECLAMP1 + "_" + Index, "FINLCmp_SFS5370"));             //First Clamp
                            parts.Add(new PartInfo(PIPECLAMP2 + "_" + Index, "FINLCmp_SFS5370"));             //Second Clamp
                        }
                        else            //only one clamp
                        {
                            parts.Add(new PartInfo(VERPLATE + "_" + Index, "Utility_USER_FIXED_BOX_1"));      //Second Clamp
                            parts.Add(new PartInfo(HORPLATE + "_" + Index, "Utility_USER_FIXED_BOX_1"));      //Second Clamp
                            parts.Add(new PartInfo(CONNOBJECT + "_" + Index, "Log_Conn_Part_1"));               //Rotational Connection Object
                            parts.Add(new PartInfo(ROUTECONNOBJECT + "_" + Index, "Log_Conn_Part_1"));           //connect obj on route so we can rotate a whole assy
                            parts.Add(new PartInfo(PIPECLAMP1 + "_" + Index, "FINLCmp_SFS5370"));             //First Clamp
                        }
                    }
                    else
                    {
                        parts.Add(new PartInfo(VERPLATE + "_" + Index, "Utility_USER_FIXED_BOX_1"));      //Second Clamp
                        parts.Add(new PartInfo(HORPLATE + "_" + Index, "Utility_USER_FIXED_BOX_1"));      //Second Clamp
                        parts.Add(new PartInfo(CONNOBJECT + "_" + Index, "Log_Conn_Part_1"));               //Rotational Connection Object
                        parts.Add(new PartInfo(ROUTECONNOBJECT + "_" + Index, "Log_Conn_Part_1"));           //connect obj on route so we can rotate a whole assy
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
                double plateThickness = 8.0 / 1000;

                //Getting Dimension Information
                //Interface name: IJUAFINLSrv_SFS5373_Dim:: Pipe_Nom_Dia_m, Shoe_H Shoe_W, B ,L, Weight, Max_Load
                double shoeHeight = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5373_Dim", "IJUAFINLSrv_SFS5373_Dim", "Shoe_H", "IJUAFINLSrv_SFS5373_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double B = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5373_Dim", "IJUAFINLSrv_SFS5373_Dim", "B", "IJUAFINLSrv_SFS5373_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double L = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5373_Dim", "IJUAFINLSrv_SFS5373_Dim", "L", "IJUAFINLSrv_SFS5373_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));

                //Getting Clamp Thickeness, width, and innerdia using multiple interface query                
                double clampThickness = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5370", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);
                double clampWidth = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5370", "IJUAFINL_Width", "Width", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);
                double clampInnerDiameter = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5370", "IJUAHgrPipe_Dia", "PIPE_DIA", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);

                string verticalPlateBom = "Plate SFS 2022 " + Convert.ToString(Math.Round((shoeHeight - clampThickness - plateThickness) * 1000, 0)) + "x" + Convert.ToString((L * 1000)).Split('.').GetValue(0) + "x" + Convert.ToString(Math.Round(plateThickness * 1000, 0));
                componentDictionary[VERPLATE + "_" + Index].SetPropertyValue(shoeHeight - clampThickness - plateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                componentDictionary[VERPLATE + "_" + Index].SetPropertyValue(plateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                componentDictionary[VERPLATE + "_" + Index].SetPropertyValue(L, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                componentDictionary[VERPLATE + "_" + Index].SetPropertyValue(verticalPlateBom, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");

                string horizontalPlateBom = "Plate SFS 2022 " + Convert.ToString(Math.Round(B * 1000, 0)) + "x" + Convert.ToString(Math.Round(L * 1000, 0)) + "x" + Convert.ToString(Math.Round(plateThickness * 1000, 0));
                componentDictionary[HORPLATE + "_" + Index].SetPropertyValue(plateThickness, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                componentDictionary[HORPLATE + "_" + Index].SetPropertyValue(B, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                componentDictionary[HORPLATE + "_" + Index].SetPropertyValue(L, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                componentDictionary[HORPLATE + "_" + Index].SetPropertyValue(horizontalPlateBom, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");

                //this will be overriden by super assembly
                if (!Override)     //by default it is false
                    JointHelper.CreateRigidJoint("-1", "Route", ROUTECONNOBJECT + "_" + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", VERPLATE + "_" + Index, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, clampInnerDiameter / 2 + clampThickness, 0, 0);

                //Add the joints between vertical plate and horizontal plate
                JointHelper.CreateRigidJoint(VERPLATE + "_" + Index, "EndOther", HORPLATE + "_" + Index, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                if (Clamps == 1)      //with one or two clamps
                {
                    double verticalAngle = rotation * 180 / Math.PI;
                    double hyp = 0.1;      //This is just an arbitrary number.
                    double X = Math.Sin((verticalAngle / 180 * Math.PI) - (90.0 / 180 * Math.PI)) * hyp;
                    double Y = Math.Cos((verticalAngle / 180 * Math.PI) - (90.0 / 180 * Math.PI)) * hyp;

                    //except for 45
                    if ((Math.Abs(verticalAngle) > 40 + 0.0001) && Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(Math.Abs(verticalAngle)), 45)==false)
                        
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Clamp rotation cannot exceed 40 degrees.", "", "SFS5373.cs", 194);

                    //if toggled, mirror the support
                    if (Configuration == 2)
                        verticalAngle = -verticalAngle;

                    //Add a joint between Connection 1 and the clamp pin so the clamp can spin
                    JointHelper.CreateRevoluteJoint(CONNOBJECT + "_" + Index, "Connection", PIPECLAMP1 + "_" + Index, "Pin", Axis.X, Axis.Y);

                    //Add a joint between Route and the Clamp center so the clamp can spin
                    JointHelper.CreateCylindricalJoint(ROUTECONNOBJECT + "_" + Index, "Connection", PIPECLAMP1 + "_" + Index, "Route", Axis.X, Axis.Y, 0);

                    //Add a flexible joint to the clamp
                    JointHelper.CreateCylindricalJoint(PIPECLAMP1 + "_" + Index, "Route", PIPECLAMP1 + "_" + Index, "Pin", Axis.Z, Axis.Z, 0);

                    if (normalPipeDiameterMetric > 49 / 1000.0)      //with two clamps
                    {
                        //set material attribute
                        componentDictionary[PIPECLAMP1 + "_" + Index].SetPropertyValue(material, "IJOAFINL_Material", "Material");
                        componentDictionary[PIPECLAMP2 + "_" + Index].SetPropertyValue(material, "IJOAFINL_Material", "Material");

                        //Add the joints to the clamps
                        JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", CONNOBJECT + "_" + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, Y, X, -L / 2 + clampWidth / 2);

                        JointHelper.CreateRigidJoint(PIPECLAMP1 + "_" + Index, "Route", PIPECLAMP2 + "_" + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, L - clampWidth, 0);
                    }
                    else        //only one clamp in the middle
                    {
                        //set material attribute
                        componentDictionary[PIPECLAMP1 + "_" + Index].SetPropertyValue(material, "IJOAFINL_Material", "Material");
                        JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", CONNOBJECT + "_" + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, Y, X, 0);
                    }
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

                    structConnections.Add(new ConnectionInfo(HORPLATE + "_" + Index, 1));     //partindex, structindex

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
                double rotation;

                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                rotation = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLRot", "Rot")).PropValue;
                //To get Pipe Nom Dia
                double pipeDiameter = ((PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1)).NominalDiameter.Size;

                double angle = rotation * 180.0 / Math.PI;

                if (Symbols.HgrCompareDoubleService.cmpdbl(rotation, 0) == true)
                    BOMString = "Pipe slide SFS 5373 DN " + Convert.ToString(pipeDiameter);
                else
                    BOMString = "Pipe slide SFS 5373 DN " + Convert.ToString(pipeDiameter) + " -" + Convert.ToString(Math.Round(angle, 0)) + Convert.ToString(176);

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
