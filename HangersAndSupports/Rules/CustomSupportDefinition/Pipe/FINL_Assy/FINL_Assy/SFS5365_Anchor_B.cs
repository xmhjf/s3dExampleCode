//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5365_Anchor_B.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5365_Anchor_B
//   Author       :  BS
//   Creation Date:  04/07/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  04/07/2013        BS  CR-CP-224485- Converted HS_FINL_Assy to C# .Net
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;


namespace Ingr.SP3D.Content.Support.Rules
{
    public class SFS5365_Anchor_B : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        CustomSupportDefinition shoe1, shoe2, connObj1, connObj2, stopper, l1, l2;

        string slideType = string.Empty, sSectionSize = string.Empty;
        int stopperType, clmaps;
        double shoeWidthFromDB, npdMetric, shoeHeightFromDB, shoeLength, shoeInset, clampInnerDiameter, stopperLength, gap, widthL, thicknessL, depthL, sectionL, existSteelW;
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                    slideType = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLSlideType", "SlideType")).PropValue;
                    stopperType = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLStopperType", "Type")).PropValue;
                    clmaps = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLClamps1", "Clamps")).PropValue;
                    gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLGap2", "Gap2")).PropValue;

                    shoe1 = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + slideType.Trim(), support);
                    shoe2 = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + slideType.Trim(), support);
                    l1 = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + "Utility_L_Wrapper", support);
                    l2 = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + "Utility_L_Wrapper", support);

                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    npdMetric = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
                    shoeWidthFromDB = FINLAssemblyServices.GetDataByCondition("FINLSrv_" + slideType + "_Dim", "IJUAFINLSrv_" + slideType + "_Dim", "Shoe_W", "IJUAFINLSrv_" + slideType + "_Dim", "Pipe_Nom_Dia_m", Convert.ToString(npdMetric * 1000));
                    shoeHeightFromDB = FINLAssemblyServices.GetDataByCondition("FINLSrv_" + slideType + "_Dim", "IJUAFINLSrv_" + slideType + "_Dim", "Shoe_H", "IJUAFINLSrv_" + slideType + "_Dim", "Pipe_Nom_Dia_m", Convert.ToString(npdMetric * 1000));

                    if (slideType.Equals("SFS5374"))
                        shoeLength = 350.00 / 1000.00;

                    else if (slideType.Equals("SFS5373") || slideType.Equals("SFS5376") || slideType.Equals("SFS5379") || slideType.Equals("SFS5858") || slideType.Equals("SFS5859"))
                        shoeLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_" + slideType + "_Dim", "IJUAFINLSrv_" + slideType + "_Dim", "L", "IJUAFINLSrv_" + slideType + "_Dim", "Pipe_Nom_Dia_m", Convert.ToString(npdMetric * 1000));

                    else if (slideType.Equals("SFS5375") || slideType.Equals("SFS5377") || slideType.Equals("SFS5378") || slideType.Equals("SFS5860"))
                        shoeLength = 500.00 / 1000.00;


                    if (slideType.Equals("SFS5375") || slideType.Equals("SFS5378"))
                        shoeInset = 30.00 / 1000.00;
                    else
                        shoeInset = 0;

                    sSectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLSectionSize", "SectionSize")).PropValue;
                    FINLAssemblyServices.GetCrossSectionDimensions("Euro", "L", sSectionSize, out widthL, out thicknessL, out sectionL, out depthL);
                    sectionL = shoeWidthFromDB;

                    FINLAssemblyServices.SetValueOnPropertyType(shoe2, "Override", true);
                    FINLAssemblyServices.SetValueOnPropertyType(shoe2, "Clamps", 2);

                    FINLAssemblyServices.SetValueOnPropertyType(l1, "Override", true);
                    FINLAssemblyServices.SetValueOnPropertyType(l1, "SectionSize", sSectionSize);
                    FINLAssemblyServices.SetValueOnPropertyType(l1, "SectionL", sectionL);

                    FINLAssemblyServices.SetValueOnPropertyType(l2, "Override", true);
                    FINLAssemblyServices.SetValueOnPropertyType(l2, "SectionSize", sSectionSize);
                    FINLAssemblyServices.SetValueOnPropertyType(l2, "SectionL", sectionL);

                    String partClass = string.Empty;
                    if (slideType.Equals("SFS5375") || slideType.Equals("SFS5378"))
                        partClass = "FINLCmp_SFS5372";
                    else if (slideType.Equals("SFS5860"))
                        partClass = "FINLCmp_SFS5856";
                    else
                        partClass = "FINLCmp_SFS5370";
                    //multiple data lookup
                    //Clamp inner dia
                    clampInnerDiameter = FINLAssemblyServices.GetDataByCondition(partClass, "IJUAHgrPipe_Dia", "PIPE_DIA", "IJUAFINL_PipeND_mm", "PipeND", npdMetric - 0.003, npdMetric + 0.003);
                    FINLAssemblyServices.SetValueOnPropertyType(l1, "Index", 0);
                    FINLAssemblyServices.SetValueOnPropertyType(l2, "Index", 1);
                    FINLAssemblyServices.SetValueOnPropertyType(shoe1, "Index", 0);
                    FINLAssemblyServices.SetValueOnPropertyType(shoe2, "Index", 1);
                    Collection<PartInfo> parts = new Collection<PartInfo>(shoe1.Parts.Concat(shoe2.Parts).Concat(l1.Parts).Concat(l2.Parts).ToList());

                    if (npdMetric >= 50.00 / 1000.00 && (!slideType.Equals("SFS5379")) && clmaps == 1)
                    {
                        //Stopper Length from SFS5368
                        stopperLength = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5368", "IJUAFINL_L", "L", "IJUAFINL_PipeND_mm", "PipeND", npdMetric - 0.003, npdMetric + 0.003);
                        connObj1 = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules.Conn_Obj_Wrapper", support);
                        connObj2 = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules.Conn_Obj_Wrapper", support);
                        stopper = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5368", support);

                        FINLAssemblyServices.SetValueOnPropertyType(connObj1, "Override", true);
                        FINLAssemblyServices.SetValueOnPropertyType(connObj1, "Index", 1);
                        FINLAssemblyServices.SetValueOnPropertyType(connObj2, "Override", true);
                        FINLAssemblyServices.SetValueOnPropertyType(connObj2, "Index", 2);
                        FINLAssemblyServices.SetValueOnPropertyType(stopper, "Override", true);
                        FINLAssemblyServices.SetValueOnPropertyType(stopper, "StopperType", stopperType);

                        parts = new Collection<PartInfo>(parts.Concat(stopper.Parts).Concat(connObj1.Parts).Concat(connObj2.Parts).ToList());

                    }
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
                    routeConnections = shoe1.SupportedConnections;
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
                    structConnections = shoe1.SupportingConnections;
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

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                ReadOnlyCollection<object> shoe1Joints = FINLAssemblyServices.GetSubAssemblyJoints(this, shoe1, oSupCompColl);
                ReadOnlyCollection<object> shoe2Joints = FINLAssemblyServices.GetSubAssemblyJoints(this, shoe2, oSupCompColl);
                ReadOnlyCollection<object> l1Joints = FINLAssemblyServices.GetSubAssemblyJoints(this, l1, oSupCompColl);
                ReadOnlyCollection<object> l2Joints = FINLAssemblyServices.GetSubAssemblyJoints(this, l2, oSupCompColl);
                JointHelper.m_oCollOfJoints = new Collection<object>(shoe1Joints.Concat(shoe2Joints).Concat(l1Joints).Concat(l2Joints).ToList());

                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                if (supportingType == "Slab")
                    existSteelW = 2 * gap;
                else if (supportingType == "Steel")
                    existSteelW = SupportingHelper.SupportingObjectInfo(1).Width + 2 * gap;


                if (existSteelW > shoeLength)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Existing steel width should not exceed shoe1 length.", "", "SFS5365_Anchor_B.cs", 203);
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Existing steel width should not exceed shoe2 length.", "", "SFS5365_Anchor_B.cs", 204);
                }
                Collection<ConnectionInfo> iSubAssyRouteConShoe1 = shoe1.SupportedConnections;
                Collection<ConnectionInfo> iSubAssyRouteConShoe2 = shoe2.SupportedConnections;
                //upper shoe
                JointHelper.CreateRigidJoint("-1", "Route", iSubAssyRouteConShoe2[0].PartKey, "Connection", Plane.YZ, Plane.YZ, Axis.Z, Axis.NegativeZ, 0, 0, 0);

                Collection<ConnectionInfo> iSubAssyRouteConL1 = l1.SupportedConnections;
                Collection<ConnectionInfo> iSubAssyRouteConL2 = l2.SupportedConnections;
                JointHelper.CreateRigidJoint("-1", "Route", iSubAssyRouteConL1[0].PartKey, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, clampInnerDiameter / 2 + shoeHeightFromDB, -shoeWidthFromDB / 2, existSteelW / 2);
                JointHelper.CreateRigidJoint("-1", "Route", iSubAssyRouteConL2[0].PartKey, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.Y, clampInnerDiameter / 2 + shoeHeightFromDB, shoeWidthFromDB / 2, -existSteelW / 2);

                if (npdMetric >= 50.0 / 1000 && (!slideType.Equals("SFS5379")) && clmaps == 1)
                {
                    ReadOnlyCollection<object> connObj1Joints = FINLAssemblyServices.GetSubAssemblyJoints(this, connObj1, oSupCompColl);
                    ReadOnlyCollection<object> connObj2Joints = FINLAssemblyServices.GetSubAssemblyJoints(this, connObj2, oSupCompColl);
                    ReadOnlyCollection<object> stopperJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, stopper, oSupCompColl);
                    JointHelper.m_oCollOfJoints = new Collection<object>(JointHelper.m_oCollOfJoints.Concat(connObj1Joints).Concat(connObj2Joints).Concat(stopperJoints).ToList());

                    Collection<ConnectionInfo> subAssyRoutConObj1 = connObj1.SupportedConnections;
                    Collection<ConnectionInfo> subAssyRoutConObj2 = connObj2.SupportedConnections;
                    Collection<ConnectionInfo> subAssyRoutConStopper = stopper.SupportedConnections;

                    double varAngle = -45, sign = 1, hyp = 0.001, X, Y;
                    if (Configuration == 2)
                    {
                        varAngle = varAngle - 180;
                        sign = -sign;
                    }
                    X = Math.Sin((varAngle - 90) / 180 * Math.PI) * hyp;
                    Y = Math.Cos((varAngle - 90) / 180 * Math.PI) * hyp;


                    JointHelper.CreateRevoluteJoint(subAssyRoutConObj1[0].PartKey, "Connection", subAssyRoutConObj2[0].PartKey, "Connection", Axis.X, Axis.X);
                    JointHelper.CreateCylindricalJoint("-1", "Route", subAssyRoutConStopper[0].PartKey, "Connection", Axis.X, Axis.X, 0);
                    JointHelper.CreateCylindricalJoint(subAssyRoutConStopper[0].PartKey, "Connection", subAssyRoutConObj2[0].PartKey, "Connection", Axis.Z, Axis.Z, 0);
                    JointHelper.CreateRigidJoint("-1", "Route", subAssyRoutConObj1[0].PartKey, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, X, Y, sign * (shoeLength / 2 + stopperLength / 2 - shoeInset));
                }

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of SFS5365_Anchor_B." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string BOMString = "";
            try
            {
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                return BOMString = "Anchor B SFS 5365 DN " + pipeInfo.NominalDiameter.Size;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5365_Anchor_B" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion


    }

}