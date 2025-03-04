//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5378.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5378
//   Author       :  Vijay
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     Vijay     CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
//   11-Dec-2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    public class SFS5378 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        double rotation, normalPipeDiameterMetric;
        string material;

        private const string CONNOBJECT = "ConnObj_5378";
        private const string LEFTLEG = "LeftLeg_5378";
        private const string RIGHTLEG = "RightLeg_5378";
        private const string CSECTION1 = "CSection1_5378";
        private const string CSECTION2 = "CSection2_5378";
        private const string CSECTION3 = "CSection3_5378";
        private const string CSECTION4 = "CSection4_5378";
        private const string CSECTION5 = "CSection5_5378";      //hor c-section
        private const string CSECTION6 = "CSection6_5378";      //hor c-section
        private const string ROUTECONNOBJECT = "RouteConnObject_5378";
        private const string PIPECLAMP1 = "PipeClamp1_5378";
        private const string PIPECLAMP2 = "PipeClamp2_5378";
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
                    minNominalDiameter.Size = 600;
                    minNominalDiameter.Units = "mm";
                    NominalDiameter maxNominalDiameter = new NominalDiameter();
                    maxNominalDiameter.Size = 1200;
                    maxNominalDiameter.Units = "mm";
                    NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 650, 750, 850, 950, 1100 }, "mm");

                    //check valid pipe size
                    if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5378.cs", 99);
                        return parts;
                    }

                    if (Clamps == 1)       //with clamps
                    {
                        parts.Add(new PartInfo(CONNOBJECT + "_" + Index, "Log_Conn_Part_1"));               //Rotational Connection Object
                        parts.Add(new PartInfo(LEFTLEG + "_" + Index, "Utility_GENERIC_L_1"));            //First Clamp
                        parts.Add(new PartInfo(RIGHTLEG + "_" + Index, "Utility_GENERIC_L_1"));           //Second Clamp
                        parts.Add(new PartInfo(CSECTION1 + "_" + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION2 + "_" + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION3 + "_" + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION4 + "_" + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION5 + "_" + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION6 + "_" + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(ROUTECONNOBJECT + "_" + Index, "Log_Conn_Part_1"));           //connect obj on route so we can rotate a whole assy
                        parts.Add(new PartInfo(PIPECLAMP1 + "_" + Index, "FINLCmp_SFS5372"));             //First Clamp
                        parts.Add(new PartInfo(PIPECLAMP2 + "_" + Index, "FINLCmp_SFS5372"));             //Second Clamp
                    }
                    else            //with out clamp
                    {
                        parts.Add(new PartInfo(CONNOBJECT + "_" + Index, "Log_Conn_Part_1"));               //Rotational Connection Object
                        parts.Add(new PartInfo(LEFTLEG + "_" + Index, "Utility_GENERIC_L_1"));            //First Clamp
                        parts.Add(new PartInfo(RIGHTLEG + "_" + Index, "Utility_GENERIC_L_1"));           //Second Clamp
                        parts.Add(new PartInfo(CSECTION1 + "_" + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION2 + "_" + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION3 + "_" + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION4 + "_" + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION5 + "_" + Index, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION6 + "_" + Index, "Utility_GENERIC_C_1"));
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
                double length = 500.00 / 1000.00;

                //Getting Dimension Information
                //Interface name: IJUAFINLSrv_SFS5378_Dim
                double shoeH = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5378_Dim", "IJUAFINLSrv_SFS5378_Dim", "Shoe_H", "IJUAFINLSrv_SFS5378_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double E = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5378_Dim", "IJUAFINLSrv_SFS5378_Dim", "E", "IJUAFINLSrv_SFS5378_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double A = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5378_Dim", "IJUAFINLSrv_SFS5378_Dim", "A", "IJUAFINLSrv_SFS5378_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double B = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5378_Dim", "IJUAFINLSrv_SFS5378_Dim", "B", "IJUAFINLSrv_SFS5378_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));
                double S = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5378_Dim", "IJUAFINLSrv_SFS5378_Dim", "S", "IJUAFINLSrv_SFS5378_Dim", "Pipe_Nom_Dia_m", Convert.ToString(normalPipeDiameterMetric * 1000));

                //Getting Clamp Thickeness, width, and innerdia using multiple interface query
                //Interface name: IJUAFINLSrv_SFS5378_Dim
                double clampWidth = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5372", "IJUAFINL_B", "B", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);
                double clampInnerDiameter = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5372", "IJUAHgrPipe_Dia", "PIPE_DIA", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);
                double DPartA = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5372", "IJUAFINL_A", "A", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);
                double DPartM = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5372", "IJUAFINL_M", "M", "IJUAFINL_PipeND_mm", "PipeND", normalPipeDiameterMetric - 0.003, normalPipeDiameterMetric + 0.003);

                //Setting attribute of two plates L sections: Utility_GENERIC_L L, WIDTH, DEPTH, THICKNESS, BOM_DESC
                string leftLegBom = "L" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, A, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, B, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, S, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[LEFTLEG + "_" + Index].SetPropertyValue(length, "IJOAHgrUtility_GENERIC_L", "L");
                componentDictionary[LEFTLEG + "_" + Index].SetPropertyValue(B, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                componentDictionary[LEFTLEG + "_" + Index].SetPropertyValue(A, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                componentDictionary[LEFTLEG + "_" + Index].SetPropertyValue(S, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                componentDictionary[LEFTLEG + "_" + Index].SetPropertyValue(leftLegBom, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");

                string rightLegBom = "L" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, A, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, B, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, S, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[RIGHTLEG + "_" + Index].SetPropertyValue(length, "IJOAHgrUtility_GENERIC_L", "L");
                componentDictionary[RIGHTLEG + "_" + Index].SetPropertyValue(B, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                componentDictionary[RIGHTLEG + "_" + Index].SetPropertyValue(A, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                componentDictionary[RIGHTLEG + "_" + Index].SetPropertyValue(S, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                componentDictionary[RIGHTLEG + "_" + Index].SetPropertyValue(rightLegBom, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");

                //add two C shape guides  -UPN100
                //Get steel Data

                double widthC, flangeTC, depthC, webTC;
                FINLAssemblyServices.GetCrossSectionDimensions("EURO-OTUA-2002", "C", "UPN100", out widthC, out flangeTC, out depthC, out webTC);

                double offset = Math.Sqrt((DPartA / 2 - DPartM / 2) * (DPartA / 2 - DPartM / 2) - (E / 2 + S + widthC) * (E / 2 + S + widthC));
                double lengthC = clampInnerDiameter / 2 + shoeH - A + 50.0 / 1000 - offset;

                string CSectionBom = "UPN100 Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthC, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[CSECTION1 + "_" + Index].SetPropertyValue(lengthC, "IJOAHgrUtility_L", "L");
                componentDictionary[CSECTION1 + "_" + Index].SetPropertyValue(widthC, "IJOAHgrUtility_Width", "Width");
                componentDictionary[CSECTION1 + "_" + Index].SetPropertyValue(flangeTC, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[CSECTION1 + "_" + Index].SetPropertyValue(webTC, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[CSECTION1 + "_" + Index].SetPropertyValue(depthC, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[CSECTION1 + "_" + Index].SetPropertyValue(CSectionBom, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                componentDictionary[CSECTION2 + "_" + Index].SetPropertyValue(lengthC, "IJOAHgrUtility_L", "L");
                componentDictionary[CSECTION2 + "_" + Index].SetPropertyValue(widthC, "IJOAHgrUtility_Width", "Width");
                componentDictionary[CSECTION2 + "_" + Index].SetPropertyValue(flangeTC, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[CSECTION2 + "_" + Index].SetPropertyValue(webTC, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[CSECTION2 + "_" + Index].SetPropertyValue(depthC, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[CSECTION2 + "_" + Index].SetPropertyValue(CSectionBom, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                componentDictionary[CSECTION3 + "_" + Index].SetPropertyValue(lengthC, "IJOAHgrUtility_L", "L");
                componentDictionary[CSECTION3 + "_" + Index].SetPropertyValue(widthC, "IJOAHgrUtility_Width", "Width");
                componentDictionary[CSECTION3 + "_" + Index].SetPropertyValue(flangeTC, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[CSECTION3 + "_" + Index].SetPropertyValue(webTC, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[CSECTION3 + "_" + Index].SetPropertyValue(depthC, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[CSECTION3 + "_" + Index].SetPropertyValue(CSectionBom, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                componentDictionary[CSECTION4 + "_" + Index].SetPropertyValue(lengthC, "IJOAHgrUtility_L", "L");
                componentDictionary[CSECTION4 + "_" + Index].SetPropertyValue(widthC, "IJOAHgrUtility_Width", "Width");
                componentDictionary[CSECTION4 + "_" + Index].SetPropertyValue(flangeTC, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[CSECTION4 + "_" + Index].SetPropertyValue(webTC, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[CSECTION4 + "_" + Index].SetPropertyValue(depthC, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[CSECTION4 + "_" + Index].SetPropertyValue(CSectionBom, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                //horizontal CSection
                CSectionBom = "UPN100 Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, E, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);       //update the CBom

                componentDictionary[CSECTION5 + "_" + Index].SetPropertyValue(E, "IJOAHgrUtility_L", "L");
                componentDictionary[CSECTION5 + "_" + Index].SetPropertyValue(widthC, "IJOAHgrUtility_Width", "Width");
                componentDictionary[CSECTION5 + "_" + Index].SetPropertyValue(flangeTC, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[CSECTION5 + "_" + Index].SetPropertyValue(webTC, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[CSECTION5 + "_" + Index].SetPropertyValue(depthC, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[CSECTION5 + "_" + Index].SetPropertyValue(CSectionBom, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                componentDictionary[CSECTION6 + "_" + Index].SetPropertyValue(E, "IJOAHgrUtility_L", "L");
                componentDictionary[CSECTION6 + "_" + Index].SetPropertyValue(widthC, "IJOAHgrUtility_Width", "Width");
                componentDictionary[CSECTION6 + "_" + Index].SetPropertyValue(flangeTC, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[CSECTION6 + "_" + Index].SetPropertyValue(webTC, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[CSECTION6 + "_" + Index].SetPropertyValue(depthC, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[CSECTION6 + "_" + Index].SetPropertyValue(CSectionBom, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                //this will be overriden by super assembly
                if (!Override)     //by default it is false
                    JointHelper.CreateRigidJoint("-1", "Route", ROUTECONNOBJECT + "_" + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                //Add four UPN50s
                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", CSECTION1 + "_" + Index, "EndCap", Plane.XY, Plane.NegativeYZ, Axis.Y, Axis.Z, offset, length / 2 + depthC / 2 - clampWidth / 2 - 30.0 / 1000, E / 2 + S);

                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", CSECTION2 + "_" + Index, "EndCap", Plane.XY, Plane.NegativeYZ, Axis.Y, Axis.Z, offset, -length / 2 + depthC / 2 + clampWidth / 2 + 30.0 / 1000, E / 2 + S);

                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", CSECTION3 + "_" + Index, "EndCap", Plane.XY, Plane.NegativeYZ, Axis.Y, Axis.NegativeZ, offset, length / 2 - depthC / 2 - clampWidth / 2 - 30.0 / 1000, -E / 2 - S);

                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", CSECTION4 + "_" + Index, "EndCap", Plane.XY, Plane.NegativeYZ, Axis.Y, Axis.NegativeZ, offset, -length / 2 - depthC / 2 + clampWidth / 2 + 30.0 / 1000, -E / 2 - S);

                //Hor CSections
                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", CSECTION5 + "_" + Index, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, clampInnerDiameter / 2 + shoeH - widthC - 40.0 / 1000, E / 2, length / 2 + depthC / 2 - clampWidth / 2);

                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", CSECTION6 + "_" + Index, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, clampInnerDiameter / 2 + shoeH - widthC - 40.0 / 1000, E / 2, -length / 2 + depthC / 2 + clampWidth / 2);

                //Add the joints between
                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", LEFTLEG + "_" + Index, "Structure", Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeY, -length / 2, offset + A + lengthC - 50.0 / 1000, E / 2);

                //Add the joints between
                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", RIGHTLEG + "_" + Index, "Structure", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.Y, offset + A + lengthC - 50.0 / 1000, length / 2, -E / 2);

                //two clamps
                if (Clamps == 1)
                {
                    //set material attribute
                    componentDictionary[PIPECLAMP1 + "_" + Index].SetPropertyValue(material, "IJOAFINL_Material", "Material");
                    componentDictionary[PIPECLAMP2 + "_" + Index].SetPropertyValue(material, "IJOAFINL_Material", "Material");

                    double verticalAngle = rotation * 180 / Math.PI;
                    double hyp = 0.1;      //This is just an arbitrary number.
                    double X = Math.Sin((verticalAngle / 180 * Math.PI) - (90.0 / 180 * Math.PI)) * hyp;
                    double Y = Math.Cos((verticalAngle / 180 * Math.PI) - (90.0 / 180 * Math.PI)) * hyp;

                    //except for 45
                    if ((Math.Abs(verticalAngle) > 40 + 0.0001) && Symbols.HgrCompareDoubleService.cmpdbl(Math.Round(Math.Abs(verticalAngle)), 45) == false)
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Clamp rotation cannot exceed 40 degrees.", "", "SFS5376.cs", 286);

                    //if toggled, mirror the support
                    if (Configuration == 2)
                        verticalAngle = -verticalAngle;

                    //Add a joint between Connection 1 and the clamp pin so the clamp can spin
                    JointHelper.CreateRevoluteJoint(CONNOBJECT + "_" + Index, "Connection", PIPECLAMP1 + "_" + Index, "Pin", Axis.X, Axis.Y);

                    //Add a joint between Route and the Clamp center so the clamp can spin
                    JointHelper.CreateCylindricalJoint(ROUTECONNOBJECT + "_" + Index, "Connection", PIPECLAMP1 + "_" + Index, "Route", Axis.X, Axis.Y, 0);

                    //Add a flexible joint to the clamp
                    JointHelper.CreateCylindricalJoint(PIPECLAMP1 + "_" + Index, "Route", PIPECLAMP1 + "_" + Index, "Pin", Axis.Z, Axis.Z, 0);

                    //Add the joints to the clamps
                    JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", CONNOBJECT + "_" + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, Y, X, -length / 2 + clampWidth / 2 + 30.0 / 1000);

                    JointHelper.CreateRigidJoint(PIPECLAMP1 + "_" + Index, "Route", PIPECLAMP2 + "_" + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, length - clampWidth - 60.0 / 1000, 0);
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

                    structConnections.Add(new ConnectionInfo(LEFTLEG + "_" + Index, 1));      //partindex, structindex
                    structConnections.Add(new ConnectionInfo(RIGHTLEG + "_" + Index, 1));     //partindex, structindex

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
                double rotationbom = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAFINLRot", "Rot")).PropValue;

                //To get Pipe Nom Dia
                double pipeDiameter = ((PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1)).NominalDiameter.Size;

                double angle = rotationbom * 180.0 / Math.PI;

                if (Symbols.HgrCompareDoubleService.cmpdbl(rotationbom, 0) == true)
                    BOMString = "Pipe slide SFS 5378 DN " + Convert.ToString(pipeDiameter);
                else
                    BOMString = "Pipe slide SFS 5378 DN " + Convert.ToString(pipeDiameter) + " -" + Convert.ToString(Math.Round(angle, 0)) + Convert.ToString(176);

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
